VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBilling 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pembuatan Invoice"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   14085
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdTutup2 
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
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   1800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer2 
      Height          =   8055
      Left            =   13560
      TabIndex        =   108
      Top             =   1320
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
   Begin VB.Timer TimerFiskal 
      Interval        =   2000
      Left            =   0
      Top             =   0
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
      Left            =   15360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   720
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
      Left            =   14280
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   720
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   14280
      TabIndex        =   107
      Top             =   240
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
   Begin VB.Frame frmLengkap 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   6120
      TabIndex        =   116
      Top             =   10080
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CheckBox opLengkap 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lainnya"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   120
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox opLengkap 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SPK"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   119
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton btnSimpan 
         BackColor       =   &H00FF0000&
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
         Left            =   1200
         Picture         =   "frmBilling.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   1200
         Width           =   1245
      End
      Begin VB.CommandButton btnCancel 
         BackColor       =   &H00FF0000&
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
         Left            =   2520
         Picture         =   "frmBilling.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   1200
         Width           =   1245
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   115
         Left            =   1200
         TabIndex        =   121
         Top             =   600
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
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
         Index           =   116
         Left            =   1200
         TabIndex        =   122
         Top             =   840
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
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
      Begin VB.Image Image1 
         Height          =   1710
         Left            =   0
         Picture         =   "frmBilling.frx":09FA
         Top             =   120
         Width           =   5025
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      LargeChange     =   800
      Left            =   12960
      Max             =   2600
      SmallChange     =   600
      TabIndex        =   91
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   480
      ScaleHeight     =   5985
      ScaleWidth      =   12945
      TabIndex        =   63
      Top             =   3360
      Width           =   12975
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   8535
         Left            =   0
         TabIndex        =   64
         Top             =   -120
         Width           =   12615
         Begin VB.Frame frmSimpanWarranty 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   2880
            TabIndex        =   137
            Top             =   2520
            Visible         =   0   'False
            Width           =   3975
            Begin VB.CommandButton btnBatalWar 
               BackColor       =   &H000000FF&
               Caption         =   "Batal"
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
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   143
               Top             =   1200
               Width           =   1215
            End
            Begin VB.CommandButton btnSimpanWar 
               BackColor       =   &H0000FF00&
               Caption         =   "Simpan"
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
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   142
               Top             =   1200
               Width           =   1215
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   42
               Left            =   1200
               TabIndex        =   139
               Top             =   480
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
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
               Index           =   49
               Left            =   1200
               TabIndex        =   141
               Top             =   840
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
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
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Caption         =   "DATA WARRANTY"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   60
               TabIndex        =   144
               Top             =   45
               Width           =   3855
            End
            Begin VB.Label Label33 
               BackColor       =   &H00C0C0C0&
               Caption         =   "No WCAR"
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
               Left            =   120
               TabIndex        =   140
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Label32 
               BackColor       =   &H00C0C0C0&
               Caption         =   "No Claim"
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
               Left            =   120
               TabIndex        =   138
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.ComboBox text42 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   11160
            Style           =   2  'Dropdown List
            TabIndex        =   134
            Top             =   6120
            Width           =   1215
         End
         Begin TrueOleDBGrid70.TDBGrid jobGrid 
            Height          =   2295
            Left            =   0
            TabIndex        =   125
            Top             =   240
            Width           =   10455
            _ExtentX        =   18441
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
            Columns(6).NumberFormat=   "Standard"
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
            Columns(9).NumberFormat=   "Standard"
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
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14933984
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
            Splits(0)._ColumnProps(12)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(17)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(22)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(27)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(33)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(37)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(41)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(45)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(49)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(52)=   "Column(11).Visible=0"
            Splits(0)._ColumnProps(53)=   "Column(11).Order=12"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
         End
         Begin VB.CommandButton TblSimpan 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Simpan"
            Height          =   375
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   8040
            Width           =   1935
         End
         Begin VB.CommandButton cetakAdd 
            BackColor       =   &H00FF0000&
            Height          =   400
            Left            =   10465
            MaskColor       =   &H008080FF&
            Picture         =   "frmBilling.frx":20A8
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   260
            Width           =   1980
         End
         Begin VB.CommandButton cetakSublet 
            BackColor       =   &H00FFC0C0&
            Height          =   360
            Left            =   10440
            MaskColor       =   &H008080FF&
            Picture         =   "frmBilling.frx":28D7
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   4080
            Width           =   2000
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1290
            Left            =   10460
            TabIndex        =   80
            Top             =   600
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   21
               Left            =   915
               TabIndex        =   6
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Mask            =   "999,99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   22
               Left            =   720
               TabIndex        =   81
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
               TabIndex        =   82
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
               TabIndex        =   85
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
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
               TabIndex        =   84
               Top             =   525
               Width           =   495
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
               TabIndex        =   83
               Top             =   885
               Width           =   510
            End
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
            Left            =   1440
            TabIndex        =   93
            Top             =   7080
            Width           =   11000
         End
         Begin VB.CommandButton cmdSublet 
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   10460
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmBilling.frx":30B7
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   4440
            Width           =   2000
         End
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
            Height          =   390
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   8520
            Width           =   1890
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1290
            Left            =   10460
            TabIndex        =   74
            Top             =   2640
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   23
               Left            =   915
               TabIndex        =   7
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
               TabIndex        =   75
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
               TabIndex        =   76
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
               TabIndex        =   79
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
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
               TabIndex        =   78
               Top             =   525
               Width           =   300
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
               TabIndex        =   77
               Top             =   840
               Width           =   180
            End
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1170
            Left            =   10460
            TabIndex        =   69
            Top             =   4800
            Width           =   2055
            Begin VB.Frame Frame4 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   375
               Left            =   0
               TabIndex        =   131
               Top             =   760
               Width           =   2055
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   26
               Left            =   840
               TabIndex        =   8
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   28
               Left            =   720
               TabIndex        =   1
               Top             =   795
               Visible         =   0   'False
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Mask            =   "99999999"
               PromptChar      =   ">"
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
               Left            =   120
               TabIndex        =   73
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
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
               TabIndex        =   72
               Top             =   525
               Width           =   450
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Material"
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
               Left            =   75
               TabIndex        =   71
               Top             =   840
               Visible         =   0   'False
               Width           =   570
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   0
            TabIndex        =   65
            Top             =   7440
            Width           =   12525
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   32
               Left            =   1920
               TabIndex        =   5
               Top             =   165
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   503
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
               Index           =   33
               Left            =   11475
               TabIndex        =   66
               Top             =   165
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCC"
               PromptChar      =   ">"
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRC"
               Height          =   195
               Left            =   11040
               TabIndex        =   68
               Top             =   210
               Width           =   330
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Suplay Slip/Keterangan"
               Height          =   255
               Left            =   0
               TabIndex        =   67
               Top             =   180
               Width           =   1755
            End
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   30
            Left            =   1440
            TabIndex        =   4
            Top             =   6735
            Width           =   5655
            _ExtentX        =   9975
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
         Begin MSMask.MaskEdBox text 
            Height          =   405
            Index           =   29
            Left            =   10560
            TabIndex        =   86
            Top             =   6600
            Width           =   1900
            _ExtentX        =   3360
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
            Index           =   46
            Left            =   9240
            TabIndex        =   123
            Top             =   2640
            Width           =   1200
            _ExtentX        =   2117
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
            Mask            =   "CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   2055
            Left            =   0
            TabIndex        =   126
            Top             =   2640
            Width           =   10455
            _ExtentX        =   18441
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
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14933984
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
            Splits(0)._ColumnProps(16)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(21)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "PART ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
         Begin TrueOleDBGrid70.TDBGrid subletGrid 
            Height          =   1815
            Left            =   0
            TabIndex        =   127
            Top             =   4800
            Width           =   10455
            _ExtentX        =   18441
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
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14933984
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
            Splits(0)._ColumnProps(12)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(26)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "SUBLET ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
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
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Materai"
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
            Left            =   10560
            TabIndex        =   92
            Top             =   6165
            Width           =   540
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            Height          =   255
            Left            =   405
            TabIndex        =   90
            Top             =   6780
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Terbilang (Rupiah)"
            Height          =   195
            Left            =   60
            TabIndex        =   89
            Top             =   7200
            Width           =   1305
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8880
            TabIndex        =   88
            Top             =   6645
            Width           =   1530
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
            Left            =   10680
            TabIndex        =   87
            Top             =   6360
            Width           =   1920
         End
      End
   End
   Begin VB.CheckBox ChkCtkRate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rate tercetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12120
      TabIndex        =   12
      Top             =   60
      Width           =   1380
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2760
      TabIndex        =   11
      Top             =   45
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   300
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   45
      Width           =   1300
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      Left            =   3840
      TabIndex        =   45
      Top             =   40
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   480
      TabIndex        =   47
      Top             =   2400
      Width           =   13005
      Begin VB.CheckBox chPWC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "PWC"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   146
         Top             =   600
         Width           =   1095
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
         TabIndex        =   57
         Top             =   0
         Width           =   1095
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
         TabIndex        =   56
         Top             =   0
         Width           =   975
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
         TabIndex        =   55
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   8400
         TabIndex        =   48
         Top             =   570
         Width           =   1170
         _ExtentX        =   2064
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   19
         Left            =   6915
         TabIndex        =   49
         Top             =   240
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   35
         Left            =   120
         TabIndex        =   52
         Top             =   285
         Width           =   1410
         _ExtentX        =   2487
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   36
         Left            =   120
         TabIndex        =   54
         Top             =   570
         Width           =   3975
         _ExtentX        =   7011
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
         Index           =   37
         Left            =   5040
         TabIndex        =   58
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   38
         Left            =   6915
         TabIndex        =   60
         Top             =   570
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   31
         Left            =   8400
         TabIndex        =   96
         Top             =   240
         Width           =   1170
         _ExtentX        =   2064
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   43
         Left            =   5040
         TabIndex        =   101
         Top             =   570
         Width           =   1170
         _ExtentX        =   2064
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   47
         Left            =   5520
         TabIndex        =   128
         Top             =   600
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   48
         Left            =   10305
         TabIndex        =   129
         Top             =   240
         Width           =   990
         _ExtentX        =   1746
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
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label JualNRM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCM"
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
         Left            =   11500
         TabIndex        =   145
         Top             =   550
         Width           =   1500
      End
      Begin VB.Label lblServicePoint 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT SERVICE POINT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4080
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label25 
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
         Left            =   9885
         TabIndex        =   130
         Top             =   285
         Width           =   300
      End
      Begin VB.Label LblReturn 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RETURN JOB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11500
         TabIndex        =   103
         Top             =   325
         Visible         =   0   'False
         Width           =   1500
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
         Left            =   4320
         TabIndex        =   102
         Top             =   615
         Width           =   570
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
         Left            =   7935
         TabIndex        =   97
         Top             =   285
         Width           =   345
      End
      Begin VB.Label lblPDS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11500
         TabIndex        =   94
         Top             =   100
         Visible         =   0   'False
         Width           =   1500
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
         Left            =   6435
         TabIndex        =   61
         Top             =   615
         Width           =   375
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
         Left            =   4680
         TabIndex        =   59
         Top             =   285
         Width           =   210
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
         TabIndex        =   53
         Top             =   30
         Width           =   615
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
         Left            =   6420
         TabIndex        =   51
         Top             =   285
         Width           =   390
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
         Left            =   7680
         TabIndex        =   50
         Top             =   615
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   480
      TabIndex        =   13
      Top             =   360
      Width           =   12970
      Begin VB.CommandButton cmdLengkap 
         BackColor       =   &H00FF0000&
         Height          =   420
         Left            =   5640
         MaskColor       =   &H00FF0000&
         Picture         =   "frmBilling.frx":3853
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   24
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   4
         Left            =   930
         TabIndex        =   14
         Top             =   510
         Width           =   3975
         _ExtentX        =   7011
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
         Index           =   5
         Left            =   930
         TabIndex        =   15
         Top             =   800
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
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
         Index           =   7
         Left            =   930
         TabIndex        =   17
         Top             =   1380
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
         Left            =   5670
         TabIndex        =   18
         Top             =   210
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   12
         Left            =   5670
         TabIndex        =   19
         Top             =   510
         Width           =   2295
         _ExtentX        =   4048
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
         Index           =   15
         Left            =   5670
         TabIndex        =   20
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
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
         Index           =   14
         Left            =   5670
         TabIndex        =   21
         Top             =   810
         Width           =   3135
         _ExtentX        =   5530
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
         Index           =   13
         Left            =   8010
         TabIndex        =   22
         Top             =   510
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   23
         Top             =   240
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
         Left            =   9615
         TabIndex        =   25
         Top             =   210
         Width           =   1170
         _ExtentX        =   2064
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   9615
         TabIndex        =   26
         Top             =   510
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   7590
         TabIndex        =   27
         Top             =   195
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   930
         TabIndex        =   28
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
         Left            =   3450
         TabIndex        =   29
         Top             =   1380
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
         Left            =   9615
         TabIndex        =   30
         Top             =   795
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   41
         Left            =   3840
         TabIndex        =   62
         Top             =   240
         Width           =   825
         _ExtentX        =   1455
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   40
         Left            =   2500
         TabIndex        =   95
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
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
         Index           =   34
         Left            =   9615
         TabIndex        =   98
         Top             =   1080
         Width           =   3300
         _ExtentX        =   5821
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
         Index           =   44
         Left            =   9615
         TabIndex        =   110
         Top             =   1380
         Width           =   3300
         _ExtentX        =   5821
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
         Index           =   45
         Left            =   9615
         TabIndex        =   113
         Top             =   1680
         Width           =   3300
         _ExtentX        =   5821
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
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   0
         X2              =   12960
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         X1              =   0
         X2              =   12960
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         X1              =   0
         X2              =   12960
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
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
         Left            =   9150
         TabIndex        =   114
         Top             =   1725
         Width           =   435
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   9120
         TabIndex        =   111
         Top             =   1425
         Width           =   465
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
         Left            =   465
         TabIndex        =   100
         Top             =   1725
         Width           =   315
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
         Left            =   9240
         TabIndex        =   99
         Top             =   1125
         Width           =   345
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   375
         TabIndex        =   43
         Top             =   555
         Width           =   405
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   42
         Top             =   855
         Width           =   495
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   450
         TabIndex        =   41
         Top             =   1125
         Width           =   330
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   405
         TabIndex        =   40
         Top             =   1425
         Width           =   375
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
         Left            =   4875
         TabIndex        =   39
         Top             =   255
         Width           =   645
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
         Left            =   5100
         TabIndex        =   38
         Top             =   525
         Width           =   420
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
         Left            =   5040
         TabIndex        =   37
         Top             =   1125
         Width           =   480
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   36
         Top             =   240
         Width           =   660
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
         Left            =   8920
         TabIndex        =   35
         Top             =   255
         Width           =   645
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
         Left            =   9075
         TabIndex        =   34
         Top             =   540
         Width           =   510
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
         Left            =   6840
         TabIndex        =   33
         Top             =   240
         Width           =   585
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
         Left            =   2865
         TabIndex        =   32
         Top             =   1425
         Width           =   435
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
         Left            =   9045
         TabIndex        =   31
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton TblKeluar 
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
      Height          =   400
      Left            =   11440
      Picture         =   "frmBilling.frx":3F84
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   9390
      Width           =   2010
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   100
      Left            =   12360
      TabIndex        =   132
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
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
      Mask            =   "CCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.CheckBox chCtkDisc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discount Tidak tercetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10080
      TabIndex        =   44
      Top             =   60
      Width           =   2100
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   12360
      TabIndex        =   133
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TESTING, NIS SVC PENYERAGAMAN AKUN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   135
      Top             =   9360
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9855
      Left            =   0
      Picture         =   "frmBilling.frx":4653
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   480
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
      Left            =   3135
      TabIndex        =   46
      Top             =   40
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. WO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   10
      Top             =   45
      Width           =   675
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsTemp As ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsSublet As ADODB.Recordset
Dim rsJobDta As New ADODB.Recordset
Dim rsCustdat As New ADODB.Recordset
Dim rsInvSublet As New ADODB.Recordset
Dim objBilling As New BClsBilling
Dim Report As Object
Dim rsGetDate As New ADODB.Recordset
Dim rsMappingCustomer As ADODB.Recordset

Dim subJob As Double
Dim subBody As Double
Dim subPart As Double
Dim subOli As Double
Dim subSublet As Double

Dim adaAddJob As Boolean
Dim statusSBI As Boolean

Dim subPartGenuine As Double
Dim subPartNGenuine As Double
Dim subOliGenuine As Double
Dim subOliNGenuine As Double

Dim nHeavy As Boolean
Dim nTHS As Boolean
Dim nRC As Boolean
Dim gSpart As Boolean
Dim adaSublet As Boolean
Dim simpan As Boolean
Dim boleh As Boolean
Dim tipeByr As String
Dim gKdPlgDealer As String
Dim woStat As String
Dim nDuoInv As Boolean

Dim vTipeCustomer As String
Dim noKFaktur As String
Dim tglKFaktur As Date
Dim paid As String
Dim totRate As Double
Dim limitMOU As Boolean
Dim vKodeLgn As String
Dim vKeterangan As String

' UNTUK INVOICE KEDUA
Dim gKdPlgDealer2 As String
Dim gKodePlg2 As String
Dim gInvCust2 As String
Dim gInsCust2 As String
Dim gInvAlmt2 As String
Dim gInvKota2 As String
Dim gInvNPWP2 As String
Dim vNamaTabel As String

' untuk faktur pajak
Dim noFaktur1 As String
Dim noFaktur2 As String
Dim support As New clsFormDealer
Dim strSql As String
Dim PSQL As String
Dim nSQLo(4) As String
Dim nRangeSyarat As Double
Dim nRangeSyarat2 As Double
Private txt5 As String

Private isFP_OK As Boolean
Private isNoFP_OK As Boolean
Private isNoFP_New As Double

Private KodeCabangFaktur As String
Private ExecMRS As String
Private ExecMRSDo As Boolean
Private nSisa As Integer
Private nSQLUPDATECounterFP As String
Private nKD As String

Private rsFI As ADODB.Recordset
Private isUpdateCol As Boolean
Private pesanIT As String
Private dDate As Date
Private isLockDiscKacab As Boolean

Private Function isiNoInv() As String
1        On Error GoTo isiNoInv_Error

2         isiNoInv = objHT.FetchDataBengkel_String("SELECT FN_NO_INV FROM T_Profil")
3        Call objHT.UpdateData_Bengkel("UPDATE T_PROFIL SET FN_NO_INV = " & isiNoInv + 1)
         
         'On Error GoTo 0
4        Exit Function

isiNoInv_Error:
5         erTranAll = True
6         ErrMsg.judul = ""
7         ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiNoInv of Form frmBilling")
8         ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiNoInv of Form frmBilling")
End Function


Private Sub doCreatePart()
   On Error GoTo doCreatePart_Error
    If rsPart.State Then rsPart.Close
   ' Set rsPart = Nothing
    With rsPart
        'If .State Then .Close
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NAMAITEM", adVarChar, 200
        .Fields.Append "QTY", adInteger
        .Fields.Append "HARGA", adDouble
        .Fields.Append "OLI", adBoolean
        .Fields.Append "EXTDISC", adDouble
        .Open
    End With

   'On Error GoTo 0
   Exit Sub

doCreatePart_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreatePart of Form frmBilling")
End Sub

Private Sub doHeadPart()
          
1        On Error GoTo doHeadPart_Error

2         With partGrid
3             .Columns(0).Width = 1800
4             .Columns(0).Caption = "KODE PART"
5             .Columns(0).DataField = "KODEITEM"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 2500
8             .Columns(1).Caption = "NAMA PART"
9             .Columns(1).DataField = "NAMAITEM"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 800
12            .Columns(2).Caption = "QTY"
13            .Columns(2).DataField = "QTY"
14            .Splits(0).Columns(2).Locked = True
              
15            .Columns(3).Caption = "HARGA"
16            .Columns(3).DataField = "HRG"
17            .Splits(0).Columns(3).Locked = True
              
18            .Columns(4).Caption = "OLI"
19            .Columns(4).DataField = "OLI"
20            .Splits(0).Columns(4).Locked = True
21        End With

22       'On Error GoTo 0
23       Exit Sub

doHeadPart_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadPart of Form frmBilling")
End Sub

Private Sub doCreateJob()
    
   On Error GoTo doCreateJob_Error
    If rsJob.State Then rsJob.Close
   ' Set rsJob = Nothing
    With rsJob
        'If .State Then .Close
        .Fields.Append "JOB", adVarChar, 50
        .Fields.Append "RATE", adDouble
        .Fields.Append "ACTRT", adDouble
        .Fields.Append "GRP", adVarChar, 6
        .Fields.Append "MECH", adVarChar, 10
        .Fields.Append "ADDJOB", adBoolean
        .Fields.Append "HRG", adDouble
        .Fields.Append "DISC", adDouble
        .Fields.Append "RP", adBoolean
        .Fields.Append "DISCRP", adDouble
        .Fields.Append "STLDISC", adDouble
        .Fields.Append "VALID", adVarChar, 10
        .Fields.Append "FREEC", adBoolean           ' TANDA JIKA DIBEBANKAN KE CUSTOMER
        .Open
    End With

   'On Error GoTo 0
   Exit Sub

doCreateJob_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreateJob of Form frmBilling")
End Sub

Private Sub doHeadJob()
          
1        On Error GoTo doHeadJob_Error

2         With JobGrid
3             .Columns(0).Width = 4300
4             .Columns(0).Caption = "JOB ORDER"
5             .Columns(0).DataField = "JOB"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 600
8             .Columns(1).Caption = "RATE"
9             .Columns(1).DataField = "RATE"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 0
12            .Columns(2).Caption = "ACTRATE"
13            .Columns(2).DataField = "ACTRT"
              '.Columns(2).Visible = False
14            .Splits(0).Columns(2).Locked = True
              
15            .Columns(3).Width = 0
16            .Columns(3).Caption = "GROUP"
17            .Columns(3).DataField = "GRP"
              '.Columns(3).Visible = False
18            .Splits(0).Columns(3).Locked = True
              
19            .Columns(4).Width = 0
20            .Columns(4).Caption = "MEKANIK"
21            .Columns(4).DataField = "MECH"
              '.Columns(4).Visible = False
22            .Splits(0).Columns(4).Locked = True
              
23            .Columns(5).Width = 0
24            .Columns(5).Caption = "Additional Job"
25            .Columns(5).DataField = "ADDJOB"
              '.Columns(5).Visible = False
26            .Splits(0).Columns(5).Locked = True
              
27            .Columns(6).Width = 1200
28            .Columns(6).Caption = "HARGA"
29            .Columns(6).DataField = "HRG"
30            .Splits(0).Columns(6).Locked = True
                     
31            .Columns(7).Width = 500
32            .Columns(7).Caption = "DISC"
33            .Columns(7).DataField = "DISC"
34            If Log.discKacab = 1 And isLockDiscKacab = True Then
35                .Splits(0).Columns(7).Locked = True
36            End If
              
37            .Columns(8).Width = 400
38            .Columns(8).Caption = "RP"
39            .Columns(8).DataField = "RP"
40            .Columns(8).Alignment = dbgCenter
              
41            .Columns(9).Width = 1200
42            .Columns(9).Caption = "DISC RUPIAH"
43            .Columns(9).DataField = "DISCRP"
              
44            .Columns(10).Width = 1200
45            .Columns(10).Caption = "SUB TOTAL"
46            .Columns(10).DataField = "STLDISC"
47            .Splits(0).Columns(10).Locked = True
              
48            .Columns(11).Width = 400
49            .Columns(11).Caption = "C"
50            .Columns(11).DataField = "FREEC"
51            .Columns(11).Alignment = dbgCenter
52            .Splits(0).Columns(11).Locked = False
              '.Columns(11).Visible = False
53            If nDuoInv Then
54                .Columns(11).Visible = True
55            End If
56        End With

         'On Error GoTo 0
57       Exit Sub

doHeadJob_Error:

58        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadJob of Form frmBilling")
End Sub

Private Sub doHeadSublet()
    With subletGrid
        .Columns(0).Width = 1500
        .Columns(0).Caption = "NO NOTA"
        .Columns(0).DataField = "NONOTA"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 1800
        .Columns(1).Caption = "SUPPLIER"
        .Columns(1).DataField = "SUPPLIER"
        .Columns(1).Visible = True
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 4000
        .Columns(2).Caption = "NAMA BARANG"
        .Columns(2).DataField = "NAMABARANG"
        .Columns(2).Visible = True
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1500
        .Columns(3).Caption = "KODE SATUAN"
        .Columns(3).DataField = "KODESATUAN"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 800
        .Columns(4).Caption = "JUMLAH"
        .Columns(4).DataField = "QTY"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 800
        .Columns(5).Caption = "VALID"
        .Columns(5).DataField = "VALID"
        '.Columns(5).Visible = False
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Width = 1500
        .Columns(6).Caption = "HARGA"
        .Columns(6).DataField = "HARGA"
        .Columns(6).NumberFormat = "#,#0"
        .Splits(0).Columns(6).Locked = True
    End With
End Sub

Private Sub hapusForm()
   On Error GoTo hapusForm_Error
   Unload frmOSublet
   
    Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & Trim(text(0)) & "'")
    dDate = "01/01/1900"
    For i = 0 To 48
        Select Case i
            Case 1, 11, 43
                text(i).Mask = "99-99-9999"
            Case 10, 21, 22, 23, 24, 25, 26, 27, 28, 29, 39
                text(i) = 0
            Case 47
                text(i).Mask = "99:99"
            Case Else
                If i <> 42 Then text(i) = ""
        End Select
    Next
    
    VScroll1.value = 0
    opWO(0).value = True
    
    Set rsJob = Nothing
    Set rsPart = Nothing
    Set rsSublet = Nothing
    
    Set JobGrid.DataSource = Nothing
    Set partGrid.DataSource = Nothing
    Set subletGrid.DataSource = Nothing
    
    adaAddJob = False
    cetakSublet.Visible = False
    cetakAdd.Visible = False
    isJualNRM = False
    vTipeCustomer = ""
    woStat = ""
    nDuoInv = False
    noKFaktur = ""
    tglKFaktur = "01/01/1900"
    TblSimpan.Enabled = False
    noFaktur1 = ""
    noFaktur2 = ""
    If rsJobDta.State <> adStateClosed Then rsJobDta.Close
    Dim rsMet As ADODB.Recordset
    Set rsMet = objO.FetchData("Select FN_Nilai From T_Meterai WITH (NOLOCK) order by FN_Nilai")
    
    With Text42
        .Clear
        rsMet.MoveFirst
        Do Until rsMet.EOF
            .AddItem rsMet.Fields(0).value
            rsMet.MoveNext
        Loop
        .text = "0"
    End With
    subBody = 0
    subJob = 0
    subOli = 0
    subPart = 0
    subSublet = 0
    subOliGenuine = 0
    subOliNGenuine = 0
    subPartGenuine = 0
    subPartNGenuine = 0
    chPWC.value = False
    chPWC.Enabled = False
    
   'On Error GoTo 0
   Exit Sub

hapusForm_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmBilling")
End Sub

Private Sub panggilWO(noWO As String)
      ''logwashing = 1

      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset
      Dim rsHlp3 As ADODB.Recordset
      Dim rsHlp4 As ADODB.Recordset
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim fOth As Boolean
      Dim vHitung As Double
      Dim nHarga As Double
          
1        On Error GoTo panggilWO_Error
2         isLockDiscKacab = False
3         gKdPlgDealer = ""
4         gKodePlg = ""
5         gInvCust = ""
6         gInvAlmt = ""
7         gInvKota = ""
8         gInvNPWP = ""
9         nDuoInv = False
10        paid = ""
11        totRate = 0
          
12        JobGrid.Splits(0).Columns(9).Locked = True
          
13        If Log.lockWashing = 1 Then
            ''cek washing apa tidak , kalo washing dicek proses washing sudah finish. Klo belum data tidak ditampilkan
            ''washing belum selesai , wo tidak dapat diproses
14          Set rsTemp = objO.FetchData("SELECT WASHING FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' AND WASHING = '1'")
15          If Not rsTemp.EOF Then
            
16             Set rsTemp = objO.FetchData("SELECT * FROM TempProgressCuci WITH (NOLOCK) WHERE NOORDER = '" & noWO & "' AND FLAG = 'F' ")
17             If rsTemp.EOF Then
                  ''jika belum finish cuci / pengering
18                MsgBox "Unit belum finish cuci / pengering, tidak dapat melanjutkan proses Invoice", vbInformation
19                Exit Sub
20             End If
21          End If
            
22        End If
          
23        Set rsTemp = objBilling.searchOrderDetail(noWO)
24        If Not rsTemp.EOF Then
25            lblServicePoint.Visible = False
26            If rsTemp("job") = "1" Then
27                lblServicePoint.Visible = True
28            End If
29            TblSimpan.Enabled = True
              
              ' TIPE PEMBAYARAN
30            Set rsHlp = objBilling.searchPaymentType(noWO)
31            If Not rsHlp.EOF Then
32                tipeByr = Nul(Trim(rsHlp("PEMBAYARAN")))
33                gVerifikator = Nul(Trim(rsHlp("VERIFIKATOR")))
34            End If
35            Set rsHlp = objBilling.searchDtlOrder(rsTemp("NO_WO"), rsTemp("NO_CHASSIS"))
36            If Not rsHlp.EOF Then
37                nHeavy = IIf(rsHlp("HEAVYUSE") <> 0, True, False)
38            End If
              
              ' TAMPIL DETAIL CUSTOMER
39            text(2) = Nul(rsTemp("FS_KD_PLG"))
40            text(3) = Nul(rsTemp("FS_KD_WIL"))
41            text(4) = Nul(rsTemp("CUSTOMER"))
42            txt5 = Nul(rsTemp("CUST_ADRES"))
43            text(5) = Left(Nul(rsTemp("CUST_ADRES")), 60)
44            text(6) = Nul(rsTemp("KOTA"))
45            text(7) = Nul(rsTemp("TELPON"))
46            text(8) = Nul(rsTemp("GSM"))
47            text(9) = Nul(rsTemp("CDMA"))
48            text(10) = Nul(rsTemp("FS_TH_MOBIL"))
49            text(11).Mask = ""
50            text(11) = Nul(IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DOTAM"), "/", "-"))))
51            text(12) = Nul(rsTemp("MDL"))
52            text(13) = Nul(rsTemp("MODEL"))
53            text(14) = Nul(rsTemp("FMDL"))
54            text(15) = Nul(rsTemp("COLOR"))
55            text(16) = Nul(rsTemp("NO_POLISI"))
56            text(17) = Nul(rsTemp("NO_CHASSIS"))
57            text(18) = Nul(rsTemp("NO_ENGINE"))
58            text(19) = Nul(rsTemp("FS_LOG_ADD"))
59            If Nul(Trim(rsTemp("NOINV"))) <> "" Then
60                text(20) = Nul(Trim(rsTemp("NOINV")))
61            End If
62            text(31) = Nul(rsTemp("NOFAK"))
63            text(33) = Nul(rsTemp("FS_RC"))
64            text(37) = Nul(rsTemp("KM"))
65            text(38) = Nul(rsTemp("MECH"))
66            text(40) = Nul(rsTemp("STATUS"))
67            text(43) = TglDDMMYYYY(rsTemp("DATE"))
68            text(44) = Nul(Trim(rsTemp("FS_NM_DEALER")))
69            nKD = objO.FetchData_String("Select KDPLG from T_Dealer WITH (NOLOCK) Where FS_KD_DEALER = '" & Nul(Trim(rsTemp("FS_KD_DEALER"))) & "' ")
70            text(47) = Trim(rsTemp("TIME"))
71            text(48) = Trim(rsTemp("NOFK"))
72            nTHS = IIf(rsTemp("THS") = 1, True, False)
73            LblReturn.Visible = IIf(rsTemp("FN_RETURN") = "1", True, False)
74            lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
              
              ''2016-05-23 || penambahan ddate dari data wo
75            dDate = Format(IIf(Trim(rsTemp!dDate) = "", "01/01/1900", Trim(rsTemp!dDate)), "mm/dd/yyyy")
              
              ' CEK NPWP ACC
76            text(45) = ""
77                If OnTopSV = 0 Then
78                    Set rsHlp = objO.FetchDataACC("SELECT NPWP, SyaratBayar FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(text(2)) & "'")
                      'UNTUK FEBUARI
79                Else
80                    Set rsHlp = objO.FetchDataACC("SELECT NPWP, TOPGRP as SyaratBayar FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(text(2)) & "'")
81                End If
82                If Not rsHlp.EOF Then
83                    text(45) = Nul(Trim(rsHlp("NPWP")))
84                    text(41) = Nul(Trim(rsHlp("SyaratBayar")))
85                End If
86            End If
87            myChassis = Trim(text(17))
              ' CHECK DEFAULT DISCOUNT
88            text(30) = ""
89            text(35).Visible = False
90            text(36).Visible = False
91            vTipeCustomer = ""
92            statusSBI = False
93            If returnJob(LblReturn) Then
94                If typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
95                    opWO(1).value = True
96                    Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
97                    If Not rsHlp.EOF Then
98                        text(35) = Nul(rsHlp("ID_PROGRAM"))
99                        text(36) = Nul(rsHlp("PROGRAM"))
100                       text(35).Visible = True
101                       text(36).Visible = True
102                   End If
103                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
104                       nDuoInv = True
                          'RTJ WO BERKEMBANG
105                       Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
106                       If Not rsHlp.EOF Then
107                           gKdPlgDealer2 = Nul(Trim(rsHlp("KDPLG")))
108                           gKodePlg2 = Nul(Trim(rsHlp("KDPLG")))
109                           gInvCust2 = Nul(Trim(rsHlp("FS_NM_DEALER")))
110                           gInsCust2 = Nul(Trim(rsHlp("INIS_DEALER")))
111                           gInvAlmt2 = Nul(Trim(rsHlp("FS_ALAMAT")))
112                           gInvKota2 = Nul(Trim(rsHlp("FS_KOTA")))
113                           gInvNPWP2 = Nul(Trim(rsHlp("FS_NPWP")))
114                           If Trim(rsHlp("INTERNAL")) = 1 Then
115                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
116                                   vTipeCustomer = "AFFNONPPN"
117                               Else
118                                   vTipeCustomer = "AFFPPN"
119                               End If
120                           Else
121                               vTipeCustomer = "EXT"
122                           End If
123                       End If
124                       Call giveDiscount("100", text(21), False)
125                       text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
126                   End If
127               Else
                      ' HANYA RTJ
128                   Call giveDiscount("100", text(21), False)
129                   vTipeCustomer = "INT"
130               End If
131           Else
132               If typeJob(Trim(rsTemp("TPSVC"))) = "WARRANTY" Then
133                   opWO(2).value = True
134                   Set rsHlp = objBilling.searchWarrantyDetail(rsTemp("KODETIPE"))
135                   If Not rsHlp.EOF Then
136                       text(35) = Nul(rsHlp("ID_WARRANTY"))
137                       text(36) = Nul(rsHlp("WARRANTY"))
138                       text(35).Visible = True
139                       text(36).Visible = True
140                   End If
141                   Call giveDiscount("0", text(21), True)
142                   text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
                      
143                   Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH (NOLOCK)")
144                   If Not rsHlp.EOF Then
145                       gKdPlgDealer = Nul(Trim(rsHlp("FN_HOKDPLG")))
146                       gKodePlg = Nul(Trim(rsHlp("FN_HOKDPLG")))
147                       gInvCust = Nul(Trim(rsHlp("FN_HO")))
148                       gInvAlmt = Nul(Trim(rsHlp("FN_HOADRES")))
149                       gInvKota = Nul(Trim(rsHlp("FN_HOKOTA")))
150                       gInvNPWP = Nul(Trim(rsHlp("FN_HONPWP")))
151                       text(40) = "AFFPPN"
                          
152                       text(2) = gKodePlg
153                       text(4) = gInvCust
154                       txt5 = gInvAlmt
155                       text(5) = gInvAlmt
156                       text(6) = gInvKota
157                       text(45) = gInvNPWP
158                       text(7) = Nul(Trim(rsHlp("FN_HOTELP")))
159                       text(8) = ""
160                       text(9) = ""
161                   End If
                       
162               ElseIf typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
163                   opWO(1).value = True
                       
164                   Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
165                   If Not rsHlp.EOF Then
166                       text(35) = Nul(rsHlp("ID_PROGRAM"))
167                       text(36) = Nul(rsHlp("PROGRAM"))
168                       text(35).Visible = True
169                       text(36).Visible = True
170                   End If
171                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
172                       nDuoInv = True
                          'SBI WO BERKEMBANG
173                       If rsTemp("FS_RC") = "R001" Then
174                           statusSBI = True
175                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
176                           If Not rsHlp.EOF Then
177                               gKdPlgDealer2 = Trim(Nul(rsHlp("KDPLG")))
178                               gKodePlg2 = Trim(Nul(rsHlp("KDPLG")))
179                               gInvCust2 = Trim(Nul(rsHlp("FS_NM_DEALER")))
180                               gInsCust2 = Trim(Nul(rsHlp("INIS_DEALER")))
181                               gInvAlmt2 = Trim(Nul(rsHlp("FS_ALAMAT")))
182                               gInvKota2 = Trim(Nul(rsHlp("FS_KOTA")))
183                               gInvNPWP2 = Trim(Nul(rsHlp("FS_NPWP")))
                                  
184                               If Trim(rsHlp("INTERNAL")) = 1 Then
185                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
186                                       vTipeCustomer = "AFFNONPPN"
187                                   Else
188                                       vTipeCustomer = "AFFPPN"
189                                   End If
190                               Else
191                                   vTipeCustomer = "EXT"
192                               End If
193                           End If
194                           Call giveDiscount("100", text(21), False)
195                           text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
196                       Else
197                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
198                           If Not rsHlp.EOF Then
199                               gKdPlgDealer2 = Nul(Trim(rsHlp("KDPLG")))
200                               gKodePlg2 = Nul(Trim(rsHlp("KDPLG")))
201                               gInvCust2 = Nul(Trim(rsHlp("FS_NM_DEALER")))
202                               gInsCust2 = Nul(Trim(rsHlp("INIS_DEALER")))
203                               gInvAlmt2 = Nul(Trim(rsHlp("FS_ALAMAT")))
204                               gInvKota2 = Nul(Trim(rsHlp("FS_KOTA")))
205                               gInvNPWP2 = Nul(Trim(rsHlp("FS_NPWP")))
                                  
206                               If Trim(rsHlp("INTERNAL")) = 1 Then
207                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
208                                       vTipeCustomer = "AFFNONPPN"
209                                   Else
210                                       vTipeCustomer = "AFFPPN"
211                                   End If
212                               Else
213                                   vTipeCustomer = "EXT"
214                               End If
215                           End If
216                           Call giveDiscount("100", text(21), False)
217                           text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
218                       End If
219                   Else
220                       Call giveDiscount(objBilling.takeDiscountProgram(rsTemp("KODETIPE")), text(21), False)
221                   End If
222               ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
223                   If lblPDS.Visible Then
224                       Call giveDiscount("0", text(21), True)
                          
225                       Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
226                       If Not rsHlp.EOF Then
227                           gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
228                           If Trim(rsHlp("INTERNAL")) = 1 Then
229                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
230                                   vTipeCustomer = "AFFNONPPN"
231                               Else
232                                   vTipeCustomer = "AFFPPN"
233                               End If
234                           Else
235                               vTipeCustomer = "EXT"
236                           End If
237                       End If
238                   ElseIf rsTemp("FS_RC") = "R001" Then
                          ' HANYA SBI
239                       paid = "T"
240                       woStat = "SBI"
241                       Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
242                       If Not rsHlp.EOF Then
243                           gKdPlgDealer = Nul(Trim(rsHlp("KDPLG")))
244                           gKodePlg = Nul(Trim(rsHlp("KDPLG")))
245                           gInvCust = Nul(Trim(rsHlp("FS_NM_DEALER")))
246                           gInsCust2 = Nul(Trim(rsHlp("INIS_DEALER")))
247                           gInvAlmt = Nul(Trim(rsHlp("FS_ALAMAT")))
248                           gInvKota = Nul(Trim(rsHlp("FS_KOTA")))
249                           gInvNPWP = Nul(Trim(rsHlp("FS_NPWP")))
250                           If Trim(rsHlp("INTERNAL")) = 1 Then
251                               If Trim(gNpwp) = Trim(gInvNPWP) Then
252                                   text(40) = "AFFNONPPN"
253                               Else
254                                   text(40) = "AFFPPN"
255                               End If
256                           Else
257                               text(40) = "EXT"
258                           End If
259                       End If
260                       Call giveDiscount("0", text(21), False)
261                       text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
262                   ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
263                       Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
264                       If Not rsHlp.EOF Then
265                           text(34) = Nul(Trim(rsHlp("NO_MOU")))
266                           cmdLengkap.Visible = False
267                           If Trim(text(34)) <> "" Then
268                               cmdLengkap.Visible = True
                                  ' TAMPIL KELENGKAPAN
269                               Set rsHlp2 = objO.FetchData("SELECT SPKKet, LainKet FROM JOBLENGKAP WITH (NOLOCK) WHERE NO_WO = '" & Trim(noWO) & "'")
270                               If Not rsHlp2.EOF Then
271                                   If Trim(rsHlp2("SPKKet")) <> "" Then
272                                       opLengkap(0) = Checked
273                                       text(115) = Trim(rsHlp2("SPKKet"))
274                                   End If
275                                   If Trim(rsHlp2("LainKet")) <> "" Then
276                                       opLengkap(1) = Checked
277                                       text(115) = Trim(rsHlp2("LainKet"))
278                                   End If
279                               End If
280                           End If
281                       End If
282                       Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")), text(21), False)
283                   Else
284                       Call giveDiscount("0", text(21), True)
285                   End If
286               End If
287           End If
              
              ' TAMPIL JOB ORDER
              Dim nilaiRate As Double
288           text(22) = 0
289           subJob = 0
290           For i = 1 To 7
291               hlp = "REP" & i
292               hlp2 = "RTE" & i
293               hlp3 = "KLPKERJA" & i
294               hlp4 = "ACTRATE" & i
295               hlp5 = "MEKANIK" & i
                  
296               If rsTemp(hlp) <> "" Then
297                       fOth = False
298                       Set rsHlp = objO.FetchData("SELECT MDL FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & rsTemp("MDL") & "'")
299                       If rsHlp.EOF Then
300                           fOth = True
301                       End If
                          
302                       If returnJob(LblReturn) Then
303                           If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
304                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
305                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
306                               End If
307                           Else
308                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
309                           End If
310                       Else
311                           If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
312                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
313                           ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
314                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
315                                   If rsTemp("FS_RC") = "R001" Then
316                                       If Trim(rsTemp(hlp3)) = "SBI" Then
317                                           Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
318                                           If Not rsHlp2.EOF Then
319                                               nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
320                                               If nStatusDealer = 1 Then
321                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
322                                               Else
323                                                   If nStatusDealer = 2 Then
324                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
325                                                   ElseIf nStatusDealer = 3 Then
326                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
327                                                   End If
328                                               End If
329                                           End If
330                                       Else
331                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
332                                       End If
333                                   Else
334                                       Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
335                                       If Not rsHlp2.EOF Then
336                                           nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
337                                           If nStatusDealer = 1 Then
338                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
339                                           Else
340                                               If nStatusDealer = 2 Then
341                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
342                                               ElseIf nStatusDealer = 3 Then
343                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
344                                               End If
345                                           End If
346                                       End If
347                                   End If
348                               Else
349                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
350                               End If
351                           ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
352                               If lblPDS.Visible Then
353                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
354                               ElseIf rsTemp("FS_RC") = "R001" Then
355                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
356                                   If Not rsHlp2.EOF Then
357                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
358                                       If nStatusDealer = 1 Then
359                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
360                                       Else
361                                           If nStatusDealer = 2 Then
362                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
363                                           ElseIf nStatusDealer = 3 Then
364                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
365                                           End If
366                                       End If
367                                   End If
368                               ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
369                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
370                               Else
371                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
372                               End If
373                           End If
374                       End If

375                       If isOnCarClass = False Then
376                           MsgBox "Kelas Kendaraan Tidak Terdaftar, Invoice di Cancel !", vbCritical + vbOKOnly, "Car Class"
377                           Call hapusForm
378                           Exit Sub
379                       End If
                                      
380                       nilaiRate = CDbl(rsTemp(hlp2))
381                       If Trim(rsTemp(hlp3)) <> "SBE" And Trim(rsTemp(hlp3)) <> "SBI" Then
382                           Me.Tag = ""
383                           gVerifikator = ""
384                           If Log.validasiKabeng = 1 Then
385                               If MsgBox("Pekerjaan : " & Trim(rsTemp(hlp)) & vbCr & "Adjustment Rate? ", vbDefaultButton2 + vbYesNo, "PENYESUAIAN RATE") = vbYes Then
386                                   gBantuan = ""
387                                   frmValidasi.Tag = "ADJ"
388                                   frmValidasi.Show vbModal
389                                   Me.Tag = gBantuan
390                               End If
391                               If Me.Tag = "valid" Then
392                                   nHarga = 0
393                               End If
394                           End If
395                       End If
                          ''----------------------------------------
                          '' 2017-06-07 request disc
                          ''----------------------------------------
396                       If Log.discKacab = 1 Then
397                           Set rsHlp = objO.FetchData("SELECT discJasaValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(0)) & "' and no_chassis = '" & Trim(text(17)) & "'")
398                           If Not rsHlp.EOF Then
399                               isLockDiscKacab = True
400                               If nDuoInv = False Then
401                                 Call giveDiscount(CInt(Nul(rsHlp!discJasaValid)), text(21), False)
402                               End If
403                           End If
404                       End If
405                       text(22) = text(22) + Val(CDbl(nHarga) * nilaiRate) * ((100 - text(21)) / 100)
                          
406                       With rsJob
407                       If .State = adStateClosed Then doCreateJob
408                          .AddNew
409                          !job = Trim(rsTemp(hlp))
410                          !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
411                          totRate = totRate + nilaiRate
412                          !ACTRT = CDbl(rsTemp(hlp4))
413                          !grp = Trim(rsTemp(hlp3))
414                          !Mech = Trim(rsTemp(hlp5))
415                          !hrg = Val(CDbl(nHarga) * nilaiRate)
416                          !addJob = False
417                          !disc = CDbl(text(21))
418                          !rp = False
419                          !DISCRP = CDbl(!hrg) * (CDbl(text(21)) / 100)
420                          !STLDISC = CDbl(!hrg) * ((100 - !disc) / 100)
421                          !valid = gVerifikator
422                          !FREEC = False
423                          .Update
424                       End With
425                       subJob = subJob + Val(CDbl(nHarga) * nilaiRate) 'IIf(LblReturn.Visible, 0, (CDbl(nHarga) * CDbl(nilaiRate)))
426                       If isJualNRM = True Then
427                           Me.JualNRM.Caption = nJualNRM
428                       Else
429                           Me.JualNRM.Caption = ""
430                       End If
431               End If
432           Next
              
              ' add JOB
433           Set rsHlp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
434           If Not rsHlp2.EOF Then
435               Do While Not rsHlp2.EOF
436                   If rsHlp2("REP") <> "" Then
437                       If returnJob(LblReturn) Then
438                           If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
439                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
440                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
441                               End If
442                           Else
443                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
444                           End If
445                       Else
446                           If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
447                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
448                           ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
449                               If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
450                                   If rsTemp("FS_RC") = "R001" Then
451                                       If rsHlp2("KLPKERJA") = "SBI" Then
452                                           Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
453                                           If Not rsHlp3.EOF Then
454                                               nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
455                                               If nStatusDealer = 1 Then
456                                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
457                                               Else
458                                                   If nStatusDealer = 2 Then
459                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
460                                                   ElseIf nStatusDealer = 3 Then
461                                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
462                                                   End If
463                                               End If
464                                           End If
465                                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
466                                       Else
467                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
468                                       End If
469                                   Else
470                                       Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
471                                       If Not rsHlp3.EOF Then
472                                           nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
473                                           If nStatusDealer = 1 Then
474                                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0 '
475                                           Else
476                                               If nStatusDealer = 2 Then
477                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
478                                               ElseIf nStatusDealer = 3 Then
479                                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
480                                               End If
481                                           End If
482                                       End If
483                                   End If
484                               Else
485                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
486                               End If
487                           ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
488                               If lblPDS.Visible Then
489                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
490                               ElseIf rsTemp("FS_RC") = "R001" Then
                                       
491                                   Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
492                                   If Not rsHlp3.EOF Then
493                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
494                                       If nStatusDealer = 1 Then
495                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
496                                       Else
497                                           If nStatusDealer = 2 Then
498                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
499                                           ElseIf nStatusDealer = 3 Then
500                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
501                                           End If
502                                       End If
503                                   End If
504                               ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
505                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
506                               Else
507                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
508                               End If
509                           End If
510                       End If
                          
511                       If isOnCarClass = False Then
512                           MsgBox "Kelas Kendaraan Tidak Terdaftar, Invoice di Cancel !", vbCritical + vbOKOnly, "Car Class"
513                           Call hapusForm
514                           Exit Sub
515                       End If
                                      
                          ' ADJUSTMENT RATE
516                       nilaiRate = CDbl(rsHlp2("RTE"))
517                       If Trim(rsTemp(hlp3)) <> "SBE" And Trim(rsTemp(hlp3)) <> "SBI" Then
518                           Me.Tag = ""
519                           gVerifikator = ""
520                           If Log.validasiKabeng = 1 Then
521                               If MsgBox("Pekerjaan : " & Trim(rsHlp2("REP")) & vbCr & "Adjustment Rate? ", vbDefaultButton2 + vbYesNo, "PENYESUAIAN RATE") = vbYes Then
522                                   gBantuan = ""
523                                   frmValidasi.Tag = "ADJ"
524                                   frmValidasi.Show vbModal
525                                   Me.Tag = gBantuan
526                               End If
527                               If Me.Tag = "valid" Then
                                      'nilaiRate = 0
528                                   nHarga = 0
529                               End If
530                           End If
531                       End If
                          ''----------------------------------------
                          '' 2017-06-07 request disc
                          ''----------------------------------------
532                       If Log.discKacab = 1 Then
533                           Set rsHlp = objO.FetchData("SELECT discJasaValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(0)) & "' and no_chassis = '" & Trim(text(17)) & "'")
534                           If Not rsHlp.EOF Then
535                               isLockDiscKacab = True
536                               Call giveDiscount(CInt(Nul(rsHlp!discJasaValid)), text(21), False)
537                           End If
538                       End If
539                       text(22) = text(22) + Val(CDbl(nHarga) * CDbl(nilaiRate)) * ((100 - text(21)) / 100)
                              
540                       With rsJob
541                       If .State = adStateClosed Then doCreateJob
542                           .AddNew
543                           !job = Trim(rsHlp2("REP"))
544                           !Rate = nilaiRate 'CDbl(rsHlp2("RTE"))
545                           totRate = totRate + nilaiRate
546                           !ACTRT = CDbl(rsHlp2("ACTRTE"))
547                           !grp = Trim(rsHlp2("KLPKERJA"))
548                           !Mech = Trim(rsHlp2("MEKANIK"))
549                           !hrg = Val(CDbl(nHarga) * CDbl(nilaiRate))
550                           !addJob = True
551                           !disc = CDbl(text(21))
552                           !rp = False
553                           !DISCRP = CDbl(!hrg) * (CDbl(text(21)) / 100)
554                           !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
555                           !valid = gVerifikator
556                           !FREEC = False
557                           .Update
558                           adaAddJob = True
559                       End With
560                       subJob = subJob + Val(CDbl(nHarga) * CDbl(nilaiRate)) 'IIf(LblReturn.Visible, 0, (CDbl(nHarga) * CDbl(nilaiRate)))
561                   End If
562                   rsHlp2.MoveNext
563               Loop
564           End If
              
565           If Log.discKacab = 0 Then
                  'PEMBERIAN DISKON DARI VALIDASI BILLING
566               Set rsHlp = objO.FetchData("Select * from ValidBilling WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "' And Dellog = ''")
567               If Not rsHlp.EOF Then
568                   text(22) = 0
569                   Do While Not rsHlp.EOF
570                       rsJob.Find "JOB ='" & Trim(rsHlp("REP")) & "'", , adSearchForward, 1
571                       If Not rsJob.EOF Then
572                           rsJob("DISC") = CDbl(rsHlp("DISC"))
573                           rsJob("DISCRP") = Round(CDbl(rsHlp("HARGA")) * (CDbl(rsHlp("DISC")) / 100), 0)
574                           rsJob("STLDISC") = Round(CDbl(rsHlp("HARGA")) - (rsJob("DISCRP")), 0)
575                           text(22) = text(22) + rsJob("STLDISC")
576                       End If
577                       rsHlp.MoveNext
578                   Loop
579               End If
580           End If
581           Set JobGrid.DataSource = rsJob
582           Call doHeadJob
583           JobGrid.Columns(9).Locked = True
              
              ' TAMPIL PART ORDER
584           noKFaktur = ""
585           tglKFaktur = "01/01/1900"
586           gSpart = True
587           text(24) = 0
588           text(25) = 0
589           subPartGenuine = 0
590           subPartNGenuine = 0
591           subOliGenuine = 0
592           subOliNGenuine = 0
              Dim HargaDisc As Double
593           If rsPart.State Then rsPart.Close
              
              ' // AMBIL PART YG BAWA SENDIRI //
594           Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, TIPE, GENUINE, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli WHERE no_WO = '" & noWO & "' " _
                                      & "AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y' ") 'AND BAWA <> 'Y'
595           If Not rsHlp.EOF Then
596               rsHlp.MoveFirst
597               Do While Not rsHlp.EOF
598                   With rsPart
599                   If .State = adStateClosed Then doCreatePart
600                       gSpart = True
                          
601                       If rsHlp("BAWA") = "Y" Then
602                           .AddNew
603                           !KodeItem = "BAWA SENDIRI"
604                           !NAMAITEM = rsHlp("PARTBAWA")
605                           !qty = rsHlp("QTY")
606                           !Harga = 0
607                           !oli = IIf(rsHlp("TIPE") <> "P", True, False)
608                           .Update
609                       Else
610                           If Not gFPart Then
                                  '///// TIDAK CONNECT ACCOUNTING SPAREPART ////
611                               .AddNew
612                               !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
613                               !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
614                               !oli = IIf(rsHlp("TIPE") <> "P", True, False)
615                               !qty = rsHlp("QTY")
                                  
616                                   Set rsHlp2 = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & rsHlp("KODE_ITEM") & "'")
617                                   If Not rsHlp2.EOF Then
618                                       HargaDisc = CDbl(rsHlp2("HARGAJUALA"))
619                                       !Harga = CDbl(HargaDisc)
      '                                    If rsHlp("OLI") <> "" Then
      '                                        If rsHlp("TIPE") = "P" Then
      '                                            If rsHlp("GENUINE") = 0 Then
      '                                                ' PART GENUINE
      '                                                subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                            ElseIf rsHlp("GENUINE") = 1 Then
      '                                                ' PART NON GENUINE
      '                                                subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                            End If
      '                                            Text(24) = Text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                        ElseIf rsHlp("TIPE") = "O" Then
      '                                            If rsHlp("GENUINE") = 0 Then
      '                                                ' OLI GENUINE
      '                                                subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                            ElseIf rsHlp("GENUINE") = 1 Then
      '                                                ' OLI NON GENUINE
      '                                                subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                            End If
      '                                            Text(25) = Text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                        End If
      '                                    Else
      '                                        ' PART GENUINE
      '                                        subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                        Text(24) = Text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
      '                                    End If

620                                       If rsHlp2("SCC") <> "" Then
621                                           If Left(rsHlp2("SCC"), 3) = "PAR" Then
622                                               If rsHlp2("TGenuine") = False Then
                                                      ' PART NON GENUINE
623                                                   subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
624                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' PART  GENUINE
625                                                   subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
626                                               End If
627                                               text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
628                                           ElseIf Left(rsHlp2("SCC"), 3) = "ACC" Then
629                                               If rsHlp2("TGenuine") = False Then
                                                      ' ACCS GENUINE
630                                                   subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
631                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' ACCS NON GENUINE
632                                                   subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
633                                               End If
634                                               text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
635                                           ElseIf Left(rsHlp2("SCC"), 3) = "OLI" Then
636                                               If rsHlp2("TGenuine") = False Then
                                                      ' OLI GENUINE
637                                                   subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
638                                               ElseIf rsHlp2("TGenuine") = True Then
                                                      ' OLI NON GENUINE
639                                                   subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
640                                               End If
641                                               text(25) = text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
642                                           End If
643                                       Else
                                              ' PART GENUINE
644                                           subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
645                                           text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
646                                       End If
647                                   End If
                                  
648                               .Update
649                           End If
650                       End If
651                   End With
652                   rsHlp.MoveNext
653               Loop
654           End If
              
655           If gFPart Then
656               If Not newTipeFakturPart(Trim(text(0))) Then
                      '/// KONEKSI KE SPAREPART
657                   If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
658                       With rsPart
                               
659                           Set rsHlp2 = objO.FetchDataACCX("SELECT A.KODEJURNAL, A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " & _
                                                              "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' AND ISNULL(NOFJDULU,'') = '' " & _
                                                              "AND A.NOFAKTUR NOT IN (SELECT DISTINCT ISNULL(NOFJDULU,'') AS NOFJDULU FROM SIHDR A INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "')", SQLConACC1)
660                           If Not rsHlp2.EOF Then
661                               rsHlp2.MoveFirst
662                               Do While Not rsHlp2.EOF
663                                   text(46) = Trim(rsHlp2("KODEJURNAL"))
664                                   noKFaktur = Trim(rsHlp2("NOFAKTUR"))
665                                   tglKFaktur = Format(rsHlp2("TglFaktur"), "DD/MM/YYYY")
666                                   If .State = adStateClosed Then doCreatePart
667                                   .AddNew
668                                   !KodeItem = Trim(rsHlp2("KODEITEM"))
669                                   !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
670                                   vHitung = (Koma2Titik(rsHlp2("Qty")) * (Koma2Titik(rsHlp2("HargaJual")) - (Koma2Titik(rsHlp2("ItemDisc")) / Abs(Koma2Titik(rsHlp2("Qty"))))))
671                                   vHitung = vHitung - (Koma2Titik(rsHlp2("ExtDisPersen")) * vHitung)
672                                   If Trim(UCase(rsHlp2("TypePpn"))) = "E" Then
673                                       vHitung = Round(vHitung + vHitung * 0.1, 2)
674                                   End If
675                                   If gJurnalPart Then
676                                       vHitung = cekIncludePPn(vHitung)
677                                   End If
678                                   !qty = CDbl(rsHlp2("QTY"))
679                                   !Harga = CDbl(vHitung)
680                                   !EXTDISC = CDbl(rsHlp2("ItemDisPsn")) * 100
                  
681                                   Set rsHlp3 = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp2("KODEITEM"), 1) & "'")
682                                   Set rsHlp4 = objO.FetchDataACCX("SELECT A.hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp2("KODEITEM") & "'", SQLConACC1)
      '                                If Not rsHlp3.EOF Then
      '                                    If rsHlp3("TIPE") = "O" Then
      '                                        !oli = True
      '                                        If rsHlp3("GENUINE") = 1 Then
      '                                            subOliGenuine = subOliGenuine + CDbl(vHitung)
      '                                        ElseIf rsHlp3("GENUINE") = 0 Then
      '                                            subOliNGenuine = subOliNGenuine + CDbl(vHitung)
      '                                        End If
      '                                       Text(25) = Text(25) + CDbl(vHitung)
      '                                    ElseIf rsHlp3("TIPE") = "P" Then
      '                                        !oli = False
      '                                        If rsHlp3("GENUINE") = 1 Then
      '                                            subPartGenuine = subPartGenuine + CDbl(vHitung)
      '                                        ElseIf rsHlp3("GENUINE") = 0 Then
      '                                            subPartNGenuine = subPartNGenuine + CDbl(vHitung)
      '                                        End If
      '                                       Text(24) = Text(24) + CDbl(vHitung)
      '                                    End If
      '                                Else
      '                                    subPartGenuine = subPartGenuine + CDbl(vHitung)
      '                                    Text(24) = Text(24) + CDbl(vHitung)
      '                                End If
                                      
                                      'update rudi 7-8-2014
                                      
                                      'If Not rsHlp3.EOF And Not rsHlp4.EOF Then
683                                   If Not rsHlp4.EOF Then
684                                       If Left(rsHlp4("SCC"), 3) = "OLI" Then
685                                           !oli = True
686                                           If rsHlp4("TGenuine") = True Then
687                                               subOliGenuine = subOliGenuine + CDbl(vHitung)
688                                           ElseIf rsHlp4("TGenuine") = False Then
689                                               subOliNGenuine = subOliNGenuine + CDbl(vHitung)
690                                           End If
691                                          text(25) = text(25) + CDbl(vHitung)
                                          
692                                       Else 'If Left(rsHlp4("SCC"), 3) = "PAR" Then
693                                           !oli = False
694                                           If rsHlp4("TGenuine") = True Then
695                                               subPartGenuine = subPartGenuine + CDbl(vHitung)
696                                           ElseIf rsHlp4("TGenuine") = False Then
697                                               subPartNGenuine = subPartNGenuine + CDbl(vHitung)
698                                           End If
699                                          text(24) = text(24) + CDbl(vHitung)
      '                                    ElseIf Left(rsHlp4("SCC"), 3) = "ACC" Then
      '                                        !oli = False
      '                                        If rsHlp4("TGenuine") = True Then
      '                                            subPartGenuine = subPartGenuine + CDbl(vHitung)
      '                                        ElseIf rsHlp4("TGenuine") = False Then
      '                                            subPartNGenuine = subPartNGenuine + CDbl(vHitung)
      '                                        End If
      '                                       Text(24) = Text(24) + CDbl(vHitung)
700                                       End If
701                                   Else
702                                       subPartGenuine = subPartGenuine + CDbl(vHitung)
703                                       text(24) = text(24) + CDbl(vHitung)
704                                   End If
                                      
705                                   .Update
706                                   rsHlp2.MoveNext
707                               Loop
708                           Else
709                               strSql = "Select Z.* FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                          "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                          "(INMST.HARGAJUALA) as HARGAJUALA, " & _
                                          "INMST.NORACK,0 as DisItem, 0 as ExtDis FROM INMST " & _
                                          "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                          "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                          "Where Z.Norack = '" & Trim(text(0)) & "'"
710                               Set rsHlp3 = objO.FetchDataACCX(strSql, SQLConACC1)
                                  
711                               If Not rsHlp3.EOF Then
712                                   MsgBox "Invoice Sparepart Belum ditutup !" & vbCr & "Hubungi Bagian Sparepart.", vbInformation, "INVOICE SPAREPART"
713                                   Call objO.InsertData("INSERT INTO NOTEPART (NOSO, STATUS, TIPE) VALUES ('" & Trim(text(0)) & "', 'TUTUP', 'O')")
714                                   TblSimpan.Enabled = False
715                                   gSpart = False
716                               Else
717                                   Set rsHlp = objO.FetchData("SELECT isnull(SUM(QTYORDER),0) AS nORDER, isnull(SUM(QTYSUPPLY),0) AS SUPPLY FROM NOTEPART WITH (NOLOCK) WHERE NOSO = '" & Trim(text(0)) & "' HAVING(ISNULL(SUM(QTYORDER),0)) > 0")
718                                   If Not rsHlp.EOF Then
719                                       If CDbl(rsHlp("nORDER")) <> CDbl(rsHlp("SUPPLY")) Then
720                                           If CDbl(rsHlp("SUPPLY")) = 0 Then
721                                               MsgBox "Kebutuhan Sparepart belum tersupply !!!!"
722                                               gSpart = False
723                                           Else
724                                               If CDbl(rsHlp("nORDER")) > CDbl(rsHlp("SUPPLY")) Then
725                                                   MsgBox "Supply part tidak lengkap !!"
726                                                   gSpart = False
727                                               ElseIf CDbl(rsHlp("nORDER")) < CDbl(rsHlp("SUPPLYa")) Then
728                                                   MsgBox "Pencatatan part yang dipesan tidak sesuai dengan pengambilan Sparepart !!"
729                                                   gSpart = False
730                                               End If
731                                           End If
732                                       End If
733                                   Else
734                                       gSpart = False
735                                   End If
736                               End If
737                           End If
738                       End With
739                   End If
                      
740                   If Not gSpart Then
741                       If Get_ConnectionNisSP2(NisSPDb2, NisSPUSER, NisSPPass, NisSPKomp) Then
742                           With rsPart
743                               Set rsHlp2 = objO.FetchDataACCX("SELECT A.KODEJURNAL, A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " & _
                                                                   "FROM SIHDR A INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' AND ISNULL(NOFJDULU,'') = '' " & _
                                                                   "AND A.NOFAKTUR NOT IN (SELECT DISTINCT ISNULL(NOFJDULU,'') AS NOFJDULU FROM SIHDR A INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "')", SQLConACC2)
744                               If Not rsHlp2.EOF Then
745                                   rsHlp2.MoveFirst
746                                   text(46) = Trim(rsHlp2("KODEJURNAL"))
747                                   noKFaktur = Trim(rsHlp2("NOFAKTUR"))
748                                   tglKFaktur = Format(rsHlp2("TglFaktur"), "DD/MM/YYYY")
749                                   gSpart = True
750                                   Do While Not rsHlp2.EOF
751                                       If .State = adStateClosed Then doCreatePart
752                                       .AddNew
753                                       !KodeItem = Trim(rsHlp2("KODEITEM"))
754                                       !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
755                                       vHitung = (Koma2Titik(rsHlp2("Qty")) * (Koma2Titik(rsHlp2("HargaJual")) - (Koma2Titik(rsHlp2("ItemDisc")) / Abs(Koma2Titik(rsHlp2("Qty"))))))
756                                       vHitung = vHitung - (Koma2Titik(rsHlp2("ExtDisPersen")) * vHitung)
757                                       If Trim(UCase(rsHlp2("TypePpn"))) = "E" Then
758                                           vHitung = Round(vHitung + vHitung * 0.1, 2)
759                                       End If
760                                       If gJurnalPart Then
761                                           vHitung = cekIncludePPn(vHitung)
762                                       End If
763                                       !qty = CDbl(rsHlp2("QTY"))
764                                       !Harga = CDbl(vHitung)
765                                       !EXTDISC = CDbl(rsHlp2("ItemDisPsn")) * 100
                  
766                                       Set rsHlp3 = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp2("KODEITEM"), 1) & "'")
767                                       Set rsHlp4 = objO.FetchDataACCX("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp2("KODEITEM") & "'", SQLConACC1)
                                                      
                                          'update rudi 7-8-2014
                                          
                                          'If Not rsHlp3.EOF And Not rsHlp4.EOF Then
768                                       If Not rsHlp4.EOF Then
769                                           If Left(rsHlp4("SCC"), 3) = "OLI" Then
770                                               !oli = True
771                                               If rsHlp4("TGenuine") = True Then
772                                                   subOliGenuine = subOliGenuine + CDbl(vHitung)
773                                               ElseIf rsHlp4("TGenuine") = False Then
774                                                   subOliNGenuine = subOliNGenuine + CDbl(vHitung)
775                                               End If
776                                              text(25) = text(25) + CDbl(vHitung)
777                                           ElseIf Left(rsHlp4("SCC"), 3) = "PAR" Then
778                                               !oli = False
779                                               If rsHlp4("TGenuine") = True Then
780                                                   subPartGenuine = subPartGenuine + CDbl(vHitung)
781                                               ElseIf rsHlp4("TGenuine") = False Then
782                                                   subPartNGenuine = subPartNGenuine + CDbl(vHitung)
783                                               End If
784                                              text(24) = text(24) + CDbl(vHitung)
785                                           ElseIf Left(rsHlp4("SCC"), 3) = "ACC" Then
786                                               !oli = False
787                                               If rsHlp4("TGenuine") = True Then
788                                                   subPartGenuine = subPartGenuine + CDbl(vHitung)
789                                               ElseIf rsHlp4("TGenuine") = False Then
790                                                   subPartNGenuine = subPartNGenuine + CDbl(vHitung)
791                                               End If
792                                              text(24) = text(24) + CDbl(vHitung)
793                                           End If
794                                       Else
795                                           subPartGenuine = subPartGenuine + CDbl(vHitung)
796                                           text(24) = text(24) + CDbl(vHitung)
797                                       End If
798                                       .Update
799                                       rsHlp2.MoveNext
800                                   Loop
801                               Else
802                                   temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbDefaultButton1 + vbYesNo, "KONFIRMASI PART")
803                                   If temp = vbYes Then
804                                       gSpart = False
805                                   Else
806                                       gSpart = True
807                                   End If
808                               End If
809                           End With
810                       Else
811                           temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbDefaultButton1 + vbYesNo, "KONFIRMASI PART")
812                           If temp = vbYes Then
813                               gSpart = False
814                           Else
815                               gSpart = True
816                           End If
817                       End If
818                   End If
819               Else
820                   If partOkInvoice(Trim(text(0))) Then
821                       If FakturDb(Trim(text(0))) <> "" Then
822                           If Get_ConnectionNisSP(FakturDb(Trim(text(0))), NisSPUSER, NisSPPass, NisSPKomp) Then
823                               With rsPart
                                       
824                                   Set rsHlp2 = objO.FetchDataACCX("SELECT A.KODEJURNAL, A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " & _
                                                                      "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' AND ISNULL(NOFJDULU,'') = '' " & _
                                                                      "AND A.NOFAKTUR NOT IN (SELECT DISTINCT ISNULL(NOFJDULU,'') AS NOFJDULU FROM SIHDR A INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "')", SQLConACC1)
825                                   If Not rsHlp2.EOF Then
826                                       rsHlp2.MoveFirst
827                                       Do While Not rsHlp2.EOF
828                                           text(46) = Trim(rsHlp2("KODEJURNAL"))
829                                           noKFaktur = Trim(rsHlp2("NOFAKTUR"))
830                                           tglKFaktur = Format(rsHlp2("TglFaktur"), "DD/MM/YYYY")
831                                           If .State = adStateClosed Then doCreatePart
832                                           .AddNew
833                                           !KodeItem = Trim(rsHlp2("KODEITEM"))
834                                           !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
835                                           vHitung = (Koma2Titik(rsHlp2("Qty")) * (Koma2Titik(rsHlp2("HargaJual")) - (Koma2Titik(rsHlp2("ItemDisc")) / Abs(Koma2Titik(rsHlp2("Qty"))))))
836                                           vHitung = vHitung - (Koma2Titik(rsHlp2("ExtDisPersen")) * vHitung)
837                                           If Trim(UCase(rsHlp2("TypePpn"))) = "E" Then
838                                               vHitung = Round(vHitung + vHitung * 0.1, 2)
839                                           End If
840                                           If gJurnalPart Then
841                                               vHitung = cekIncludePPn(vHitung)
842                                           End If
843                                           !qty = CDbl(rsHlp2("QTY"))
844                                           !Harga = CDbl(vHitung)
845                                           !EXTDISC = CDbl(rsHlp2("ItemDisPsn")) * 100
                          
846                                           Set rsHlp3 = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp2("KODEITEM"), 1) & "'")
847                                           Set rsHlp4 = objO.FetchDataACCX("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp2("KODEITEM") & "'", SQLConACC1)
                                               
                                              'update rudi 7-8-2014
                                              
                                              'If Not rsHlp3.EOF And Not rsHlp4.EOF Then
848                                           If Not rsHlp4.EOF Then
849                                               If Left(rsHlp4("SCC"), 3) = "OLI" Then
850                                                   !oli = True
851                                                   If rsHlp4("TGenuine") = 1 Or rsHlp4("TGenuine") = True Then
852                                                       subOliGenuine = subOliGenuine + CDbl(vHitung)
853                                                   ElseIf rsHlp4("TGenuine") = 0 Or rsHlp4("TGenuine") = False Then
854                                                       subOliNGenuine = subOliNGenuine + CDbl(vHitung)
855                                                   End If
856                                                  text(25) = text(25) + CDbl(vHitung)
857                                               Else 'If Left(rsHlp4("SCC"), 3) = "PAR" Then
858                                                   !oli = False
859                                                   If rsHlp4("TGenuine") = 1 Or rsHlp4("TGenuine") = True Then
860                                                       subPartGenuine = subPartGenuine + CDbl(vHitung)
861                                                   ElseIf rsHlp4("TGenuine") = 0 Or rsHlp4("TGenuine") = False Then
862                                                       subPartNGenuine = subPartNGenuine + CDbl(vHitung)
863                                                   End If
864                                                  text(24) = text(24) + CDbl(vHitung)
      '                                            ElseIf Left(rsHlp4("SCC"), 3) = "ACC" Then
      '                                                !oli = False
      '                                                If rsHlp4("TGenuine") = True Then
      '                                                    subPartGenuine = subPartGenuine + CDbl(vHitung)
      '                                                ElseIf rsHlp4("TGenuine") = False Then
      '                                                    subPartNGenuine = subPartNGenuine + CDbl(vHitung)
      '                                                End If
      '                                               Text(24) = Text(24) + CDbl(vHitung)
865                                               End If
866                                           Else
867                                               subPartGenuine = subPartGenuine + CDbl(vHitung)
868                                               text(24) = text(24) + CDbl(vHitung)
869                                           End If
870                                           .Update
871                                           rsHlp2.MoveNext
872                                       Loop
873                                   Else
874                                       strSql = "Select Z.* FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                                  "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                                  "(INMST.HARGAJUALA) as HARGAJUALA, " & _
                                                  "INMST.NORACK,0 as DisItem, 0 as ExtDis FROM INMST " & _
                                                  "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                                  "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                                  "Where Z.Norack = '" & Trim(text(0)) & "'"
875                                       Set rsHlp3 = objO.FetchDataACCX(strSql, SQLConACC1)
                                          
876                                       If Not rsHlp3.EOF Then
877                                           MsgBox "Invoice Sparepart Belum ditutup !" & vbCr & "Hubungi Bagian Sparepart.", vbInformation, "INVOICE SPAREPART"
878                                           Call objO.InsertData("INSERT INTO NOTEPART (NOSO, STATUS, TIPE) VALUES ('" & Trim(text(0)) & "', 'TUTUP', 'O')")
879                                           TblSimpan.Enabled = False
880                                           gSpart = False
881                                       Else
882                                           Set rsHlp = objO.FetchData("SELECT isnull(SUM(QTYORDER),0) AS nORDER, isnull(SUM(QTYSUPPLY),0) AS SUPPLY FROM NOTEPART WITH (NOLOCK) WHERE NOSO = '" & Trim(text(0)) & "' HAVING(ISNULL(SUM(QTYORDER),0)) > 0")
883                                           If Not rsHlp.EOF Then
884                                               If CDbl(rsHlp("nORDER")) <> CDbl(rsHlp("SUPPLY")) Then
885                                                   If CDbl(rsHlp("SUPPLY")) = 0 Then
886                                                       MsgBox "Kebutuhan Sparepart belum tersupply !!!!"
887                                                       gSpart = False
888                                                   Else
889                                                       If CDbl(rsHlp("nORDER")) > CDbl(rsHlp("SUPPLY")) Then
890                                                           MsgBox "Supply part tidak lengkap !!"
891                                                           gSpart = False
892                                                       ElseIf CDbl(rsHlp("nORDER")) < CDbl(rsHlp("SUPPLY")) Then
893                                                           MsgBox "Pencatatan part yang dipesan tidak sesuai dengan pengambilan Sparepart !!"
894                                                           gSpart = False
895                                                       End If
896                                                   End If
897                                               End If
898                                           Else
899                                               temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbDefaultButton1 + vbYesNo, "KONFIRMASI PART")
900                                               If temp = vbYes Then
901                                                   gSpart = False
902                                               End If
903                                           End If
904                                       End If
905                                   End If
906                               End With
907                           End If
908                       Else
909                           If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
910                               With rsPart
911                                   Set rsHlp2 = objO.FetchDataACCX("SELECT A.KODEJURNAL, A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " & _
                                                                      "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "' AND ISNULL(NOFJDULU,'') = '' " & _
                                                                      "AND A.NOFAKTUR NOT IN (SELECT DISTINCT ISNULL(NOFJDULU,'') AS NOFJDULU FROM SIHDR A INNER JOIN SISAL B ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(text(0)) & "' and KodeLgn = '" & text(2) & "')", SQLConACC1)
912                                   If Not rsHlp2.EOF Then
913                                       rsHlp2.MoveFirst
914                                       Do While Not rsHlp2.EOF
915                                           text(46) = Trim(rsHlp2("KODEJURNAL"))
916                                           noKFaktur = Trim(rsHlp2("NOFAKTUR"))
917                                           tglKFaktur = Format(rsHlp2("TglFaktur"), "DD/MM/YYYY")
918                                           If .State = adStateClosed Then doCreatePart
919                                           .AddNew
920                                           !KodeItem = Trim(rsHlp2("KODEITEM"))
921                                           !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
922                                           vHitung = (Koma2Titik(rsHlp2("Qty")) * (Koma2Titik(rsHlp2("HargaJual")) - (Koma2Titik(rsHlp2("ItemDisc")) / Abs(Koma2Titik(rsHlp2("Qty"))))))
923                                           vHitung = vHitung - (Koma2Titik(rsHlp2("ExtDisPersen")) * vHitung)
924                                           If Trim(UCase(rsHlp2("TypePpn"))) = "E" Then
925                                               vHitung = Round(vHitung + vHitung * 0.1, 2)
926                                           End If
927                                           If gJurnalPart Then
928                                               vHitung = cekIncludePPn(vHitung)
929                                           End If
930                                           !qty = CDbl(rsHlp2("QTY"))
931                                           !Harga = CDbl(vHitung)
932                                           !EXTDISC = CDbl(rsHlp2("ItemDisPsn")) * 100
                          
933                                           Set rsHlp3 = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp2("KODEITEM"), 1) & "'")
934                                           Set rsHlp4 = objO.FetchDataACCX("SELECT A.hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp2("KODEITEM") & "'", SQLConACC1)
                                              
                                              'update rudi 7-8-2014
                                              
                                              'If Not rsHlp3.EOF And Not rsHlp4.EOF Then
935                                           If Not rsHlp4.EOF Then
                                              
936                                               If Left(rsHlp4("SCC"), 3) = "OLI" Then
937                                                   !oli = True
938                                                   If rsHlp4("TGenuine") = True Then
939                                                       subOliGenuine = subOliGenuine + CDbl(vHitung)
940                                                   ElseIf rsHlp4("TGenuine") = False Then
941                                                       subOliNGenuine = subOliNGenuine + CDbl(vHitung)
942                                                   End If
943                                                  text(25) = text(25) + CDbl(vHitung)
944                                               Else 'If Left(rsHlp4("SCC"), 3) = "PAR" Then
945                                                   !oli = False
946                                                   If rsHlp4("TGenuine") = True Then
947                                                       subPartGenuine = subPartGenuine + CDbl(vHitung)
948                                                   ElseIf rsHlp4("TGenuine") = False Then
949                                                       subPartNGenuine = subPartNGenuine + CDbl(vHitung)
950                                                   End If
951                                                  text(24) = text(24) + CDbl(vHitung)
      '                                            ElseIf Left(rsHlp4("SCC"), 3) = "ACC" Then
      '                                                !oli = False
      '                                                If rsHlp4("TGenuine") = True Then
      '                                                    subPartGenuine = subPartGenuine + CDbl(vHitung)
      '                                                ElseIf rsHlp4("TGenuine") = False Then
      '                                                    subPartNGenuine = subPartNGenuine + CDbl(vHitung)
      '                                                End If
      '                                               Text(24) = Text(24) + CDbl(vHitung)
952                                               End If
953                                           Else
954                                               subPartGenuine = subPartGenuine + CDbl(vHitung)
955                                               text(24) = text(24) + CDbl(vHitung)
956                                           End If
957                                           .Update
958                                           rsHlp2.MoveNext
959                                       Loop
960                                   Else
961                                       strSql = "Select Z.* FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                                  "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                                  "(INMST.HARGAJUALA) as HARGAJUALA, " & _
                                                  "INMST.NORACK,0 as DisItem, 0 as ExtDis FROM INMST " & _
                                                  "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                                  "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                                  "Where Z.Norack = '" & Trim(text(0)) & "'"
962                                       Set rsHlp3 = objO.FetchDataACCX(strSql, SQLConACC1)
                                          
963                                       If Not rsHlp3.EOF Then
964                                           MsgBox "Invoice Sparepart Belum ditutup !" & vbCr & "Hubungi Bagian Sparepart.", vbInformation, "INVOICE SPAREPART"
965                                           Call objO.InsertData("INSERT INTO NOTEPART (NOSO, STATUS, TIPE) VALUES ('" & Trim(text(0)) & "', 'TUTUP', 'O')")
966                                           TblSimpan.Enabled = False
967                                           gSpart = False
968                                       Else
969                                           Set rsHlp = objO.FetchData("SELECT isnull(SUM(QTYORDER),0) AS nORDER, isnull(SUM(QTYSUPPLY),0) AS SUPPLY FROM NOTEPART WITH (NOLOCK) WHERE NOSO = '" & Trim(text(0)) & "' HAVING(ISNULL(SUM(QTYORDER),0)) > 0")
970                                           If Not rsHlp.EOF Then
971                                               If CDbl(rsHlp("nORDER")) <> CDbl(rsHlp("SUPPLY")) Then
972                                                   If CDbl(rsHlp("SUPPLY")) = 0 Then
973                                                       MsgBox "Kebutuhan Sparepart belum tersupply !!!!"
974                                                       gSpart = False
975                                                   Else
976                                                       If CDbl(rsHlp("nORDER")) > CDbl(rsHlp("SUPPLY")) Then
977                                                           MsgBox "Supply part tidak lengkap !!"
978                                                           gSpart = False
979                                                       ElseIf CDbl(rsHlp("nORDER")) < CDbl(rsHlp("SUPPLY")) Then
980                                                           MsgBox "Pencatatan part yang dipesan tidak sesuai dengan pengambilan Sparepart !!"
981                                                           gSpart = False
982                                                       End If
983                                                   End If
984                                               End If
985                                           Else
986                                               temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbDefaultButton1 + vbYesNo, "KONFIRMASI PART")
987                                               If temp = vbYes Then
                                                      'MsgBox "Dbase Part tidak ketemu !"
988                                                   gSpart = False
989                                               End If
990                                           End If
991                                       End If
992                                   End If
993                               End With
994                           End If
995                       End If
996                   Else
997                       If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
998                           strSql = "Select Z.* FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                      "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN  where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                      "(INMST.HARGAJUALA) as HARGAJUALA, " & _
                                      "INMST.NORACK,0 as DisItem, 0 as ExtDis FROM INMST " & _
                                      "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                      "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                      "Where Z.Norack = '" & Trim(text(0)) & "'"
999                           Set rsHlp3 = objO.FetchDataACCX(strSql, SQLConACC1)
                              
1000                          If Not rsHlp3.EOF Then
1001                              MsgBox "Supply Slip sudah diretur ! " & vbCr & "Harap Invoice Ulang ....", vbCritical, "Supply Slip"
1002                              Call objO.InsertData("INSERT INTO NOTEPART (NOSO, STATUS, TIPE) VALUES ('" & Trim(text(0)) & "', 'TUTUP', 'O')")
1003                              TblSimpan.Enabled = False
1004                              gSpart = False
1005                          Else
1006                              temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbDefaultButton1 + vbYesNo, "KONFIRMASI PART")
1007                              If temp = vbYes Then
1008                                  gSpart = False
1009                              End If
1010                          End If
1011                          End If
1012                  End If
1013              End If
1014          End If
1015          text(24) = Round(text(24), 2)
1016          text(25) = Round(text(25), 2)
      '        subPartGenuine = Round(subPartGenuine, 0)
      '        subPartNGenuine = Round(subPartNGenuine, 0)
      '        subOliGenuine = Round(subOliGenuine, 0)
      '        subOliNGenuine = Round(subOliNGenuine, 0)
              
1017          subPart = CDbl(text(24))
1018          subOli = CDbl(text(25))
1019          Set partGrid.DataSource = rsPart
1020          Call doHeadPart
              
              ' TAMPIL SUBLET ORDER
1021          adaSublet = False
1022          cmdSublet.Visible = False
1023          text(27) = 0
1024          Set rsSublet = objO.FetchData("SELECT NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, PRICE as HARGA FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
1025          If Not rsSublet.EOF Then
1026              Do While Not rsSublet.EOF
1027                  If rsSublet("VALID") = 0 Then
1028                      adaSublet = True
1029                      MsgBox "Sublet belum di validasi"
1030                      If rsSublet.State Then rsSublet.Close
1031                      Exit Sub
1032                  End If
1033                  text(27) = text(27) + (CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY")))
1034                  rsSublet.MoveNext
1035              Loop
                  ''2016-07-22 || tanpa tempjobsublet
                  ''Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
                       & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & noWO & "' AND no_chassis = '" & rstemp("NO_CHASSIS") & "') and no_wo = '" & noWO & "' AND no_chassis = '" & rstemp("NO_CHASSIS") & "' " _
                       & "ORDER BY nonota, kodeitem")
                       
                  
1036              subSublet = CDbl(text(27))
1037              adaSublet = True
1038              cmdSublet.Visible = True
1039              Set subletGrid.DataSource = rsSublet
1040              Call doHeadSublet
1041          End If

             ' text(29) = Format(Round(CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)), 0), "#0,#")
1042          text(29) = CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28))
1043          text(29) = Format(objO.FetchData_String("SELECT ROUND(" & Replace(text(29), ",", ".") & ",0)"), "#0,#")
1044          txtTotalRp = Angka2Huruf(text(29), 1)
              
1045          Call panggilWarranty(Trim(text(0)))
         'On Error GoTo 0
1046     Exit Sub

panggilWO_Error:

1047      MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWO of Form frmBilling")

End Sub

Private Sub panggilWarranty(noWO As String)
1         On Error GoTo panggilWarranty_error
          
          Dim sPWC As Integer
          
2         sPWC = objO.FetchData_Integer("SELECT PWC FROM DIAGAPPROVAL WHERE NO_WO = '" & Trim(noWO) & "' AND ALASAN_TDKSETUJU = ''")
3         chPWC.value = sPWC
          
4         Exit Sub
          
panggilWarranty_error:
5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmBilling")
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
1        On Error GoTo returnJob_Error

2         returnJob = False
3         If labelReturn.Visible Then
4             returnJob = True
5         End If

6        'On Error GoTo 0
7        Exit Function

returnJob_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmBilling")
End Function

Private Function typeJob(statusJob As String) As String
1        On Error GoTo typeJob_Error

2         If statusJob = "W" Then
3             typeJob = "WARRANTY"
4         ElseIf statusJob = "P" Then
5             typeJob = "PROGRAM"
6         Else
7             typeJob = "REGULAR"
8         End If

9        'On Error GoTo 0
10       Exit Function

typeJob_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure typeJob of Form frmBilling")
End Function

Private Function giveDiscount(discountValue As Double, discountText As Object, statusText As Boolean)
1        On Error GoTo giveDiscount_Error

2         discountText.Mask = ""
3         discountText = Koma2Titik(CDbl(discountValue))
4         discountText.Enabled = statusText
5         JobGrid.Splits(0).Columns(7).Locked = Not statusText
6         JobGrid.Splits(0).Columns(8).Locked = Not statusText
7         JobGrid.Splits(0).Columns(9).Locked = Not statusText
8        'On Error GoTo 0
9        Exit Function

giveDiscount_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure giveDiscount of Form frmBilling")
End Function

Private Function RangeTempo(kodeLanggan As String, ind As Integer)
1     On Error GoTo vb_error
          Dim nSyaratAs As String
2         nSyaratAs = objBilling.FetchJatuhTempo(Trim(kodeLanggan))
3         Select Case UCase(Left(nSyaratAs, 1))
              Case "D"
4                 Select Case ind
                  Case 0
5                     nRangeSyarat = Val(Mid(nSyaratAs, 2, 5))
6                 Case 1
7                     nRangeSyarat2 = Val(Mid(nSyaratAs, 2, 5))
8                 End Select
9             Case "W"
10                Select Case ind
                  Case 0
11                    nRangeSyarat = Val(Mid(nSyaratAs, 2, 5)) * 7
12                Case 1
13                    nRangeSyarat2 = Val(Mid(nSyaratAs, 2, 5)) * 7
14                End Select
15            Case "M"
16                Select Case ind
                  Case 0
17                    nRangeSyarat = Val(Mid(nSyaratAs, 2, 5)) * 30
18                Case 1
19                    nRangeSyarat2 = Val(Mid(nSyaratAs, 2, 5)) * 30
20                End Select
21        End Select
22        Exit Function
vb_error:
23        MsgBox "Line number : " & Erl & vbCrLf & "Deskripsi : " & err.Description & vbCrLf & " Fungsi RangeTempo - frmBilling", vbInformation
          
End Function


Private Sub btnBatalWar_Click()
    text(49) = ""
    text(42) = ""
    frmSimpanWarranty.Visible = False
End Sub

Private Sub btnCancel_Click()
    frmLengkap.Visible = False
End Sub

Private Sub btnSimpanWar_Click()
1         On Error GoTo vb_error
2         If Trim(text(42)) = "" Then
3             MsgBox "No Claim harus diisi!", vbInformation
4             Exit Sub
5         Else
6             Set rsTemp = objO.FetchData("SELECT * FROM ClaimWarranty WHERE status = 'OK' AND NOCLAIM = '" & Trim(text(42)) & "'")
7             If Not rsTemp.EOF Then
8                 MsgBox "No Claim sudah pernah diisikan!", vbInformation
9                 Exit Sub
10            End If
11        End If
12        If Trim(text(49)) = "" Then
13            MsgBox "No WCAR harus diisi!", vbInformation
14            Exit Sub
15        End If
          
          '------------------------------------------------
          '- 2015-04-10 | RUDI | PENYIMPANAN DATA         -
          '------------------------------------------------
          
16            Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 0)
17            Call gTglServerBilling
              
18            nRangeSyarat = 0
19            nRangeSyarat2 = 0
              ' KODE PELANGGAN INVOICE PERTAMA
20            Call RangeTempo(vKodeLgn, 0)
              '' KODE PELANGGAN INVOICE KEDUA
21            If Trim(gKdPlgDealer2) <> "" Then Call RangeTempo(gKdPlgDealer2, 1)
              
22            erTranAll = False
              ' BACA JOBDTA
23            Set rsJobDta = objO.FetchData("SELECT TOP 1 * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS =  '" & Trim(text(17)) & "'")
24            Set rsCustdat = objO.FetchData("SELECT TOP 1 * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(17)) & "'")
25            Set rsInvSublet = objO.FetchData("SELECT PKP, NAMABARANG, NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, PRICE*QTY AS HARGAJUAL, NONJURNAL FROM JOBSUBLET a WITH (NOLOCK) INNER JOIN T_SUPPLIER B WITH (NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(Trim(text(0))) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
              '"SELECT PKP, NAMABARANG, NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, PRICE*QTY AS HARGAJUAL, NONJURNAL FROM JOBSUBLET a INNER JOIN T_SUPPLIER B ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(Trim(text(0))) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''"
                      
26            Call objHT.OpenTransactionAll
              '5/11/2012
27            Call objHO.AddNew_LogHO(Trim(text(0)), "BATCH", "BATCH", "-")
28            ExecMRSDo = False
29            ExecMRS = ""
30            If erTranAll = False Then Call simpanInvoice(Trim(text(0)))
31            If nDuoInv And erTranAll = False Then
32                Call simpanInvoice2(Trim(text(0)))
33            End If
              
              
              'TANGGAL 21-01-2015 RUDI
34            If erTranAll = False Then Call objHT.InsertData_Bengkel("UPDATE CHASSINGFORM SET TANGGALINVOICE = GETDATE() WHERE NOWO = '" & Trim(text(0)) & "'")
              
35            If erTranAll = False Then Call simpanHistory(Trim(text(0)))
36            If Trim(text(34)) <> "" And tipeByr = "T" And erTranAll = False Then
37                Call simpanTrnMOU(Trim(text(0)))
38            End If
              
              'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
39            If OnNewFP = 0 Then
                  'Set rsFaktur = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
40                Set rsFaktur = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
41            Else
                  'Set rsFaktur = objO.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
42                Set rsFaktur = objHT.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
43            End If
44            KodeCabangFaktur = Trim(rsFaktur("KODEDEALER"))
                  
45            If gIERC Then
                  'If CDbl(text(29)) <> 0 Then
46                    If erTranAll = False Then Call simpanFaktur(nDuoInv)
                  'End If
      '            If nDuoInv Then
      '                Call simpanFaktur2
      '            End If
47                If nDuoInv Then '
                      '' 2015-10-16 || jika npwp dealer = dealer unit maka
48                    If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) >= 0 Then
49                        If Trim(gNpwp) <> Trim(gInvNPWP2) Then
50                            If erTranAll = False Then Call simpanFaktur2
51                        End If
52                    Else
53                        If erTranAll = False Then Call simpanFaktur2
54                    End If
55                End If
56            End If
              
                      
              
              ' SIMPAN ACC
              
              'PERUBAHAN PENYERAGAMAN AKUN
              'TANGGAL 04-09-2014
              
              Dim hasilSimpanACC As Boolean
57            If erTranAll = True Then
58                hasilSimpanACC = False
59            End If
60            If erTranAll = False Then
61                If Log.isAkun = 1 Then
62                    hasilSimpanACC = simpanACC2(Trim(text(0)))
63                Else
64                    hasilSimpanACC = simpanACCNonAkunPenyeragaman(Trim(text(0)))
65                End If
66            End If
67            frmSimpanWarranty.Visible = False
             
              
68            If hasilSimpanACC Then
69                If ExecMRSDo = True Then
70                    Call objO.UpdateData(ExecMRS)
71                End If
                  
                  Dim rsQue As ADODB.Recordset
72                Set rsQue = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
73                If Not rsQue.EOF Then
74                    If Trim(rsQue("THS")) = "0" And Trim(rsQue("PDS")) = "0" And Trim(rsQue("TPSVC")) <> "W" Then
                          ' PANGGIL CUSTOMER
75                        Call finishService(Trim(text(16)), Trim(text(14)), Trim(namaModel(text(14))), namaPembawa(Trim(text(0))), Format(isNowBilling, "MM/DD/YYYY"), Format(isNowBilling, "HH:NN"), 1)
76                    End If
77                End If


                  ' SIMPAN KE HO
                  'Call simpanHO(Trim(text(0)))
                      
                  '------ 2014-12-18 RUDI KRISTANTO ---------------
                  'lakukan pengiriman ke Ho untuk hasil FI next job
                  
78                Set rsFI = objO.FetchData("SELECT * FROM INSPECTION WHERE NOWO = '" & Trim(text(0)) & "'")
79                If rsFI.RecordCount > 0 Then
                      'If Log.LogHOConn = 1 Then
      ''                If Get_ConnectionProfil_Bengkel99 Then
      ''
      ''                    'Call objO.InsertDataHO("INSERT INTO INSPECTION VALUES ('" & Trim(rsFI("NOWO")) & "','" & Format(Trim(rsFI("TGLENTRY")), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')")
      ''                    ''2015-09-15 | GANTI SIMPAN INSPECTION ''
      ''                    Call objO.InsertDataHO("INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Format(Trim(rsFI("TGLENTRY")), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')")
      ''                Else
      ''                    strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
      ''                    strSql = Replace(strSql, "'", "~")
      ''                    Dim strLog As String
      ''
      ''                    strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
      ''                   "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
      ''
      ''                    Call objO.InsertData(strLog)
      ''                End If
                      Dim strLog As String
80                    If Log.LogHOConn = 1 Then
81                        If ConHO = False Then
82                            strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
83                            strSql = Replace(strSql, "'", "~")
84                            strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                       "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
85                            Call objO.InsertData(strLog)
86                        Else
87                            If Get_ConnectionProfil_Bengkel99 = False Then
88                                strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
89                                strSql = Replace(strSql, "'", "~")
90                                strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                           "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
91                                Call objO.InsertData(strLog)
92                            Else
93                                Call objO.InsertDataHO("INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Format(Trim(rsFI("TGLENTRY")), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')")
94                            End If
95                        End If
96                    Else
97                        strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
98                        strSql = Replace(strSql, "'", "~")
99                        strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                   "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
100                       Call objO.InsertData(strLog)
101                   End If
102               End If
                      
                  '---------------------------------------------------------------
                  '- 2015-04-09 | RUDI | PENYIMPANAN DATA KE CLAIM WARRANTY      -
                  '---------------------------------------------------------------
                  
103               If opWO(2).value = True Then
104                   strSql = "INSERT INTO [ClaimWarranty] ([No_wo],[Noinv],[NoClaim],[NoWac],[Status],[Transfer],[TglTransfer],[KodeBengkel],[NamaCabang], [pwc])"
105                   strSql = strSql & " VALUES "
106                   strSql = strSql & " ('" & Trim(text(0)) & "','" & Trim(text(20)) & "','" & Trim(text(42)) & "','" & Trim(text(49)) & "','OK','0','" & Format(gTanggal, "MM/DD/YYYY") & "','" & Log.KodeDealer & "','" & Log.LogNamaNas & "','" & Trim(IIf(chPWC.value, 1, 0)) & "')"
                      'Call objHT.InsertData_Bengkel(strSql)
107                   objO.InsertData (strSql)
                      'lakukan proses transfer HO
108                   If Log.LogHOConn = 1 Then
109                       If ConHO = False Then
110                           sSQL = Replace(strSql, "'", "~")
111                           strSql = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                              "Values ( 'CLAIM','" & Trim(text(0)) & "', 'Insert', 0,'01/01/1900', '" & sSQL & "')"
                              'Call objHT.InsertData_Bengkel(strSql)
112                           Call objO.InsertData(strSql)
113                       Else
114                           If Get_ConnectionProfil_Bengkel99 = False Then
115                               sSQL = Replace(strSql, "'", "~")
116                               strSql = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                  "Values ( 'CLAIM','" & Trim(text(0)) & "', 'Insert', 0,'01/01/1900', '" & sSQL & "')"
                                  'Call objHT.InsertData_Bengkel(strSql)
117                               Call objO.InsertData(strSql)
118                           Else
                                  'Call objHT.InsertData_Ho(strSql)
119                               Call objO.InsertDataHO(strSql)
                                  'Call objHT.UpdateData_Bengkel("UPDATE [CLAIMWARRANTY] SET TRANSFER = '1' WHERE NO_WO = '" & Trim(Text(0)) & "' AND NOINV = '" & Trim(Text(20)) & "' AND NOCLAIM = '" & Trim(Text(42)) & "'")
120                               Call objO.InsertData("UPDATE [CLAIMWARRANTY] SET TRANSFER = '1' WHERE NO_WO = '" & Trim(text(0)) & "' AND NOINV = '" & Trim(text(20)) & "' AND NOCLAIM = '" & Trim(text(42)) & "'")
                                  
121                           End If
122                       End If
123                   Else
124                       sSQL = Replace(strSql, "'", "~")
125                       strSql = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                          "Values ( 'CLAIM','" & Trim(text(0)) & "', 'Insert', 0,'01/01/1900', '" & sSQL & "')"
                          'Call objHT.InsertData_Bengkel(strSql)
126                       Call objO.InsertData(strSql)
127                   End If
128               End If
                      
                      Dim nNoFak As String
129                   nNoFak = objO.FetchData_String("Select NoFAK FROM History WITH (NOLOCK) Where NO_WO = '" & Trim(text(0)) & "'")
              
130                       If Log.LogHOConn = 1 Then
131                           If ConHO = False Then
132                               Call Saving_Log_HO("Insert", sSQLi, "1")
133                               If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
134                               Call Saving_Log_HO("Insert", nSQLo(0), "i0")
135                               If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
136                               If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
137                               If nSQLo(3) <> "" Then
138                                   If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
139                               End If
                                  'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
140                               Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")

                                  'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
141                               Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
                                  
142                           Else
143                               If Get_ConnectionProfil_Bengkel99 = False Then
144                                   Call Saving_Log_HO("Insert", sSQLi, "1")
145                                   If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
146                                   Call Saving_Log_HO("Insert", nSQLo(0), "i0")
147                                   If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
148                                   If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
149                                   If nSQLo(3) <> "" Then
150                                       If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
151                                   End If
                                      'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
152                                   Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
                                      
                                      'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
153                                   Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
                                      
154                               Else
                                      
155                                   errHO = False
                                      
156                                   Call objHO.OpenTransactionLog_NRM
157                                   Call objHO.AddNew_Billing(sSQLi)
158                                   If errHO = False Then
159                                       If nDuoInv Then Call objHO.AddNew_Billing(sSQLi2)
160                                   End If
161                                   If errHO = False Then
162                                       Call objHO.AddNew_Part(nSQLo(0))
163                                   End If
164                                   If errHO = False Then
165                                       If nDuoInv Then Call objHO.AddNew_Part(nSQLo(1))
166                                   End If
167                                   If errHO = False Then
168                                       If nSQLo(2) <> "" Then Call objHO.AddNew_Part(nSQLo(2))
169                                   End If
                                      
170                                   If errHO = False Then
171                                       If nSQLo(3) <> "" Then
172                                           If nDuoInv Then Call objHO.AddNew_Part(nSQLo(3))
173                                       End If
174                                   End If
                                      
175                                   If errHO = False Then
                                          Dim rsPr As ADODB.Recordset
176                                       Set rsPr = objO.FetchData("Select * FROm T_Part WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "'")
177                                       If rsPr.RecordCount <> 0 Then
178                                           rsPr.MoveFirst
179                                           Do Until rsPr.EOF
180                                               Call objHO.AddNew_Part("Insert Into T_PARTnew Values ('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "'," & Val(rsPr.Fields(3).value) & "," & RepRp(rsPr.Fields(4).value) & "," & IIf(rsPr.Fields(5).value = False, 0, 1) & ",'" & Trim(rsPr.Fields(6).value) & "'," & Val(rsPr.Fields(7).value) & ",'','01/01/3000','" & Log.LogBengkel & "')")
181                                               rsPr.MoveNext
182                                           Loop
183                                       End If
184                                   End If
185                                   If errHO = False Then
186                                       Set rsPr = objO.FetchData("Select * FROm Hist_Sublet WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "'")
187                                       If rsPr.RecordCount <> 0 Then
188                                           rsPr.MoveFirst
189                                           Do Until rsPr.EOF
190                                               Call objHO.AddNew_Part("Insert Into Hist_SUBLETnew Values ('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "','" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "'," & Val(rsPr.Fields(8).value) & "," & RepRp(rsPr.Fields(9).value) & ",'" & Trim(rsPr.Fields(10).value) & "','" & Trim(rsPr.Fields(11).value) & "','" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & Log.LogBengkel & "')")
191                                               rsPr.MoveNext
192                                           Loop
193                                       End If
194                                   End If
                                      
195                                   If errHO = False Then
196                                       Set rsPr = objO.FetchData("Select * FROm History WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "'")
197                                       If rsPr.RecordCount <> 0 Then
198                                           rsPr.MoveFirst
199                                           Do Until rsPr.EOF
200                                               Call objHO.AddNew_Part("Insert Into Historynew Values " & _
                                                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "','" & TglNul2(rsPr.Fields(9).value) & "'," & RepRp(rsPr.Fields(10).value) & "," & RepRp(rsPr.Fields(11).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(12).value) & "," & RepRp(rsPr.Fields(13).value) & "," & RepRp(rsPr.Fields(14).value) & "," & RepRp(rsPr.Fields(15).value) & "," & RepRp(rsPr.Fields(16).value) & "," & RepRp(rsPr.Fields(17).value) & " , " & _
                                                  " " & RepRp(rsPr.Fields(18).value) & "," & RepRp(rsPr.Fields(19).value) & ",'" & Trim(rsPr.Fields(20).value) & "','" & Trim(rsPr.Fields(21).value) & "','" & Trim(rsPr.Fields(22).value) & "','" & Trim(rsPr.Fields(23).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(24).value) & "','" & Trim(rsPr.Fields(25).value) & "','" & Trim(rsPr.Fields(26).value) & "','" & Trim(rsPr.Fields(27).value) & "','" & Trim(rsPr.Fields(28).value) & "','" & Trim(rsPr.Fields(29).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(30).value) & "','" & Trim(rsPr.Fields(31).value) & "','" & Trim(rsPr.Fields(32).value) & "','" & Trim(rsPr.Fields(33).value) & "','" & Trim(rsPr.Fields(34).value) & "','" & Trim(rsPr.Fields(35).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(36).value) & "','" & Trim(rsPr.Fields(37).value) & "','" & Trim(rsPr.Fields(38).value) & "','" & Trim(rsPr.Fields(39).value) & "','" & Trim(rsPr.Fields(40).value) & "','" & Trim(rsPr.Fields(41).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(42).value) & "','" & Trim(rsPr.Fields(43).value) & "','" & Trim(rsPr.Fields(44).value) & "','" & Trim(rsPr.Fields(45).value) & "','" & Trim(rsPr.Fields(46).value) & "','" & Trim(rsPr.Fields(47).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(48).value) & "'," & Val(rsPr.Fields(49).value) & "," & Val(rsPr.Fields(50).value) & ",'" & Trim(rsPr.Fields(51).value) & "'," & Val(rsPr.Fields(52).value) & ",'" & Trim(rsPr.Fields(53).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(54).value) & "','" & Trim(rsPr.Fields(55).value) & "','" & Trim(rsPr.Fields(56).value) & "','" & Trim(rsPr.Fields(57).value) & "','" & Trim(rsPr.Fields(58).value) & "','" & Trim(rsPr.Fields(59).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(60).value) & "'," & Val(rsPr.Fields(61).value) & ",'" & Trim(rsPr.Fields(62).value) & "'," & Val(rsPr.Fields(63).value) & ",'" & Trim(rsPr.Fields(64).value) & "','" & Format(rsPr.Fields(65).value, "MM/DD/YYYY") & "', " & _
                                                  "'" & Trim(rsPr.Fields(66).value) & "','" & Trim(rsPr.Fields(67).value) & "','" & Format(rsPr.Fields(68).value, "MM/DD/YYYY") & "','" & Log.LogBengkel & "')")
                                                  
201                                               rsPr.MoveNext
202                                           Loop
203                                       End If
204                                   End If
205                                   If errHO = False Then
206                                       Call objHO.Remove_WO("Delete FROM JOBDTA  Where No_Wo = '" & Trim(text(0)) & "'")
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
207                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
208                                   End If
209                                   If errHO = False Then
210                                       Set rsPr = objO.FetchData("Select * from JOBDTA WITH (NOLOCK) Where No_Wo = '" & Trim(text(0)) & "'")
211                                       Call objHO.AddNew_Part("Insert Into JOBDTA Values " & _
                                                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "','" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "','" & Trim(rsPr.Fields(9).value) & "','" & Trim(rsPr.Fields(10).value) & "','" & Trim(rsPr.Fields(11).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & Trim(rsPr.Fields(15).value) & "','" & Trim(rsPr.Fields(16).value) & "'," & RepRp(rsPr.Fields(17).value) & "," & RepRp(rsPr.Fields(18).value) & "," & RepRp(rsPr.Fields(19).value) & "," & RepRp(rsPr.Fields(20).value) & ",'" & Trim(rsPr.Fields(21).value) & "','" & Trim(rsPr.Fields(22).value) & "','" & Trim(rsPr.Fields(23).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(24).value) & "','" & TglNul2(rsPr.Fields(25).value) & "','" & Trim(rsPr.Fields(26).value) & "','" & Trim(rsPr.Fields(27).value) & "','" & Trim(rsPr.Fields(28).value) & "','" & Trim(rsPr.Fields(29).value) & "','" & Trim(rsPr.Fields(30).value) & "','" & Trim(rsPr.Fields(31).value) & "'," & RepRp(rsPr.Fields(32).value) & "," & RepRp(rsPr.Fields(33).value) & "," & RepRp(rsPr.Fields(34).value) & "," & RepRp(rsPr.Fields(35).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(36).value) & "," & RepRp(rsPr.Fields(37).value) & "," & RepRp(rsPr.Fields(38).value) & "," & RepRp(rsPr.Fields(39).value) & ",'" & Trim(rsPr.Fields(40).value) & "','" & Trim(rsPr.Fields(41).value) & "','" & Trim(rsPr.Fields(42).value) & "','" & Trim(rsPr.Fields(43).value) & "','" & Trim(rsPr.Fields(44).value) & "','" & Trim(rsPr.Fields(45).value) & "','" & Trim(rsPr.Fields(46).value) & "','" & Trim(rsPr.Fields(47).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(48).value) & "','" & Trim(rsPr.Fields(49).value) & "','" & Trim(rsPr.Fields(50).value) & "','" & Trim(rsPr.Fields(51).value) & "','" & Trim(rsPr.Fields(52).value) & "','" & Trim(rsPr.Fields(53).value) & "','" & Trim(rsPr.Fields(54).value) & "','" & Trim(rsPr.Fields(55).value) & "','" & Trim(rsPr.Fields(56).value) & "'," & RepRp(rsPr.Fields(57).value) & "," & RepRp(rsPr.Fields(58).value) & "," & RepRp(rsPr.Fields(59).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(60).value) & "," & RepRp(rsPr.Fields(61).value) & "," & RepRp(rsPr.Fields(62).value) & "," & RepRp(rsPr.Fields(63).value) & "," & RepRp(rsPr.Fields(64).value) & ",'" & Trim(rsPr.Fields(65).value) & "','" & Trim(rsPr.Fields(66).value) & "','" & Trim(rsPr.Fields(67).value) & "','" & Trim(rsPr.Fields(68).value) & "','" & Trim(rsPr.Fields(69).value) & "','" & Trim(rsPr.Fields(70).value) & "','" & Trim(rsPr.Fields(71).value) & "', " & _
                                                  "'" & TglNul2(rsPr.Fields(72).value) & "','" & Trim(rsPr.Fields(73).value) & "','" & Trim(rsPr.Fields(74).value) & "','" & Trim(rsPr.Fields(75).value) & "','" & Trim(rsPr.Fields(76).value) & "','" & Trim(rsPr.Fields(77).value) & "','" & Trim(rsPr.Fields(78).value) & "','" & Trim(rsPr.Fields(79).value) & "','" & Trim(rsPr.Fields(80).value) & "','" & Trim(rsPr.Fields(81).value) & "','" & Trim(rsPr.Fields(82).value) & "','" & Trim(rsPr.Fields(83).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(84).value) & "','" & Trim(rsPr.Fields(85).value) & "','" & Trim(rsPr.Fields(86).value) & "','" & Trim(rsPr.Fields(87).value) & "','" & Trim(rsPr.Fields(88).value) & "','" & Trim(rsPr.Fields(89).value) & "','" & Trim(rsPr.Fields(90).value) & "','" & Trim(rsPr.Fields(91).value) & "','" & Trim(rsPr.Fields(92).value) & "','" & Trim(rsPr.Fields(93).value) & "','" & Trim(rsPr.Fields(94).value) & "','" & Trim(rsPr.Fields(95).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(96).value) & "','" & Trim(rsPr.Fields(97).value) & "','" & Trim(rsPr.Fields(98).value) & "','" & Trim(rsPr.Fields(99).value) & "','" & Trim(rsPr.Fields(100).value) & "','" & Trim(rsPr.Fields(101).value) & "'," & RepRp(rsPr.Fields(102).value) & "," & RepRp(rsPr.Fields(103).value) & "," & RepRp(rsPr.Fields(104).value) & "," & RepRp(rsPr.Fields(105).value) & "," & RepRp(rsPr.Fields(106).value) & "," & RepRp(rsPr.Fields(107).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(108).value) & "," & RepRp(rsPr.Fields(109).value) & "," & RepRp(rsPr.Fields(110).value) & "," & RepRp(rsPr.Fields(111).value) & "," & RepRp(rsPr.Fields(112).value) & "," & RepRp(rsPr.Fields(113).value) & "," & RepRp(rsPr.Fields(114).value) & "," & RepRp(rsPr.Fields(115).value) & "," & RepRp(rsPr.Fields(116).value) & "," & RepRp(rsPr.Fields(117).value) & ",'" & Trim(rsPr.Fields(118).value) & "','" & Trim(rsPr.Fields(119).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(120).value) & "','" & Trim(rsPr.Fields(121).value) & "','" & Trim(rsPr.Fields(122).value) & "','" & Trim(rsPr.Fields(123).value) & "','" & Trim(rsPr.Fields(124).value) & "','" & Trim(rsPr.Fields(125).value) & "','" & Trim(rsPr.Fields(126).value) & "','" & Trim(rsPr.Fields(127).value) & "','" & Trim(rsPr.Fields(128).value) & "'," & Val(rsPr.Fields(129).value) & ",'" & Trim(rsPr.Fields(130).value) & "','" & Trim(rsPr.Fields(131).value) & "', " & _
                                                  "'" & TglNul2(rsPr.Fields(132).value) & "','" & Trim(rsPr.Fields(133).value) & "','" & Trim(rsPr.Fields(134).value) & "','" & Trim(rsPr.Fields(135).value) & "','" & Trim(rsPr.Fields(136).value) & "','" & Trim(rsPr.Fields(137).value) & "'," & RepRp(rsPr.Fields(138).value) & ",'" & Trim(rsPr.Fields(139).value) & "','" & Trim(rsPr.Fields(140).value) & "','" & Trim(rsPr.Fields(141).value) & "','" & Trim(rsPr.Fields(142).value) & "','" & Trim(rsPr.Fields(143).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(144).value) & "','" & Trim(rsPr.Fields(145).value) & "','" & TglNul2(rsPr.Fields(146).value) & "','" & Trim(rsPr.Fields(147).value) & "','" & Trim(rsPr.Fields(148).value) & "','" & TglNul2(rsPr.Fields(149).value) & "','" & Trim(rsPr.Fields(150).value) & "')")
212                                   End If
213                                   If errHO = False Then
214                                       Set rsPr = objO.FetchData("Select * from Hist_JobAdd WITH (NOLOCK) Where No_Wo = '" & Trim(text(0)) & "'")
215                                       If rsPr.RecordCount <> 0 Then
216                                           rsPr.MoveFirst
217                                           Do Until rsPr.EOF
218                                               Call objHO.AddNew_Part("Insert Into Hist_JobAddnew Values " & _
                                                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & TglNul2(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "'," & RepRp(rsPr.Fields(5).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(6).value) & ",'" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "'," & RepRp(rsPr.Fields(9).value) & "," & RepRp(rsPr.Fields(10).value) & ",'" & Trim(rsPr.Fields(11).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & TglNul2(rsPr.Fields(15).value) & "','" & Trim(rsPr.Fields(16).value) & "','" & Log.LogBengkel & "')")
219                                               rsPr.MoveNext
220                                           Loop
221                                       End If
222                                       Call objHO.CloseTransactionLog_NRM
                                          
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
223                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
224                                   End If
225                                   If errHO = True Then
226                                       Call Saving_Log_HO("Insert", sSQLi, "1")
227                                       If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
228                                       Call Saving_Log_HO("Insert", nSQLo(0), "i0")
229                                       If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
230                                       If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
231                                       If nSQLo(3) <> "" Then
232                                           If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
233                                       End If
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
234                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
235                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
236                                   End If
237                               End If
238                           End If
239                       Else
240                           Call Saving_Log_HO("Insert", sSQLi, "1")
241                           If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
242                           Call Saving_Log_HO("Insert", nSQLo(0), "i0")
243                           If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
244                           If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
245                           If nSQLo(3) <> "" Then
246                               If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
247                           End If
                              'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
248                           Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
                              'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
249                           Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
250                       End If
              
251               If gUkuranInvoice = False Then
252                   Call cetakInv1
253                   If nDuoInv Then
254                       Call cetakInv2
255                   End If
256               Else
257                   Call CetakFormKecil
      '                If nDuoInv Then
      '                    Call CetakFormKecil2
      '                End If
258               End If
                  'TANGGAL 19-01-2015 - RUDI
                  'KIRIM SMS BILLING
259               If nPesan <> "" Then
260                  nPesan = Replace(Replace(nPesan, "#NM", Trim(namaTelepon)), "#RP", text(29))
261               End If
262               If nTuj = True And Log.smsBilling = 1 Then
263                   Set rsTemp = objO.FetchData("SELECT * FROM ANTRICUCI WHERE NOORDER = '" & Trim(text(0)) & "' AND FLAG <> 'CANCEL'")
                     ' Dim strSql As String
264                   If Not rsTemp.EOF Then
                          'jika sudah ada data, lakukan pengecekan sudah finish cuci
265                       Set rsTemp = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & Trim(text(0)) & "' AND FLAG = 'F'")
266                       If rsTemp.RecordCount > 0 Then
                              'jika sudah finish cuci
267                           strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                                  " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','1')"
268                           Call objO.InsertData(strSql)
269                           If nPesan <> "" Then
270                              strSql = "INSERT INTO " & Log.LogSMS & ".[dbo].[OUTBOX]( [Pesan], [Tujuan], [Modul], [Tanggal], [UserLog]) " & _
                                    " VALUES ('" & nPesan & "','" & noTelepon & "','NOTIF',getdate(),'" & gSA & "')"
271                              Call objO.InsertData(strSql)
272                           End If
273                       Else
                              'jika belum finish cuci
274                           strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                                  " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','0')"
275                           Call objO.InsertData(strSql)
276                       End If
277                   Else
                          'jika tidak ada data / tidak melakukan cuci insert di smsnotif dan sms
278                       strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                                  " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','1')"
279                       Call objO.InsertData(strSql)
280                       If nPesan <> "" Then
281                           strSql = "INSERT INTO " & Log.LogSMS & ".[dbo].[OUTBOX]( [Pesan], [Tujuan], [Modul], [Tanggal], [UserLog]) " & _
                                 " VALUES ('" & nPesan & "','" & noTelepon & "','NOTIF',getdate(),'" & gSA & "')"
282                           Call objO.InsertData(strSql)
283                       End If
284                   End If
285               Else
                      'jika notelp tidak valid
286                   strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                          " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','0')"
287                   Call objO.InsertData(strSql)
288                  '' If nPesan <> "" Then
289                     strSql = "INSERT INTO [" & Log.LogSMS & "].[dbo].[InvalidNumb]([Pesan], [Tujuan], [Modul], [Tanggal], [UserLog]) " & _
                              " VALUES ('" & nPesan & "','" & noTelepon & "','NOTIF',GETDATE(),'" & gSA & "')"
290                     Call objO.InsertData(strSql)
291                  '' End If
292               End If
                  
                  
                  '--------------------------
293               If gPreviewCetak = True Then
294                   text(100).SetFocus
295               End If
296               If rsJobDta.State <> adStateClosed Then rsJobDta.Close
297               If rsCustdat.State <> adStateClosed Then rsCustdat.Close
298               If rsInvSublet.State <> adStateClosed Then rsInvSublet.Close
299           Else
300               Exit Sub
301           End If
              
              'Call cekTampungan
302           TblSimpan.Enabled = False
303           If cmdSublet.Visible Then
304               cetakSublet.Visible = True
305           End If
306           cmdSublet.Visible = False
          
307       Exit Sub
vb_error:
308       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure btnSimpanWar of Form frmBilling")

End Sub

Private Sub JobGrid_LostFocus()
    'If opWO(1).value = True And JobGrid.Columns(11) Then
        '' jika program dan
    'Else
'        Call JobGrid_AfterColUpdate(7)
    'End If
    If isUpdateCol = False Then
        Call JobGrid_AfterColUpdate(7)
    End If
End Sub

Private Sub tblSimpan_Click()

      '   On Error GoTo tblSimpan_Click_Error

1     Call CexRangeFP
2     If isFP_OK = False Then
3         MsgBox "Anda tidak punya Nomor Faktur" & Chr(13) & "Hubungi Administrasi (ADH) - FP" & Chr(13) & "Invoice tidak Bisa Dibuka", vbCritical + vbOKOnly
4         Call hapusForm
5         Unload Me
6     End If
7     If isNoFP_OK = False Then
8          MsgBox "Sisa Nomor Faktur : " & isNoFP_New & Chr(13) & "Segera Hubungi Administrasi (ADH) - FP", vbInformation + vbOKOnly
9     End If
10    If isNoFP_New < 2 Then
11        MsgBox "Anda tidak punya Nomor Faktur" & Chr(13) & "Hubungi Administrasi (ADH) - FP" & Chr(13) & "Invoice tidak Bisa Dibuka", vbCritical + vbOKOnly
12        Call hapusForm
13        Unload Me
14    End If
          
      Dim boleh As Boolean
      Dim pesan As String
          
15        nSQLo(0) = ""
16        nSQLo(1) = ""
17        nSQLo(2) = ""
18        nSQLo(3) = ""
          
19        On Error GoTo vbError
20        boleh = True
21        pesan = ""
22        If text(0) = "" Then
23            pesan = pesan + "Isi no WO ! " & vbCr
24            boleh = False
25        End If
26        If limitMOU Then
27            pesan = pesan + "Tagihan MOU sudah mencapai batas. Tidak bisa invoice. " & vbCr & "Harap lunasi invoice sebelumnya. " & vbCr
28            boleh = False
29        End If
30        If adaSublet Then
31            If rsSublet.State <> adStateClosed Then
32                If subletGrid.Columns(0) = "" Or subletGrid.Columns(5) = 0 Then
33                    boleh = False
34                    pesan = pesan + "Lengkapi Data Sublet. " & vbCr
35                End If
36            Else
37                boleh = False
38                pesan = pesan + "Validasi Sublet. " & vbCr
39            End If
40        End If
41        If Not gSpart Then
42            boleh = False
43            pesan = pesan + "Invoice Sparepart belum ditutup. " & vbCr
44        End If
45        If Trim(text(13)) = "" Then
46            boleh = False
47            pesan = pesan + "Kelas Model Kendaraan belum lengkap" & vbCr
48        End If
49        If Trim(text(44)) = "" Then
50            If MsgBox("Dealer Penjual Kosong, Tetap melanjutkan ?", vbYesNo, "DEALER PENJUAL") = vbNo Then
51                boleh = False
52                pesan = pesan + "Dealer Penjual. " & vbCr
53            End If
54        End If
55        If Trim(text(41)) = "" Then
56            boleh = False
57            pesan = pesan + "Syarat Bayar. " & vbCr
58        End If
59        If CDbl(totRate) <= 0 Then
60            If subSublet = 0 Then
61                If CDbl(text(29)) - CDbl(Text42.text) = 0 Then
62                    boleh = False
63                    pesan = pesan + "Total Rate Pekerjaan NOL. Tidak boleh dibuat Invoice. " & vbCr
64                End If
65            End If
66        End If
          
67        If CDbl(Right(NisDb, 2)) <> CDbl(Month(isnow)) Then
68            boleh = False
69            pesan = pesan + "Beda Fiskal Accounting tidak dapat dibuat invoice. " & vbCr & "Pindah Fiskal Accounting pada bulan bersangkutan. " & vbCr
70        End If
          
71        If rsJob.State <> adStateClosed Then
72            If Trim(text(38)) = "" Then
73                boleh = False
74                pesan = pesan + "Mekanik tidak boleh Kosong. " & vbCr
75            End If
76        End If
          
          
          Dim rsSearch As ADODB.Recordset
77        Set rsSearch = objO.FetchData("SELECT NOINV FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
78        If Not rsSearch.EOF Then
79            If Trim(rsSearch("NOINV")) <> "" Then
80                boleh = False
81                pesan = pesan + "Invoice sudah tercetak !!!"
82            End If
83        End If
          
          ''' KODE PELANGGAN UNTUK INVOICE KHUSUS
84        vKodeLgn = ""
85        If opWO(2) Then
              ' WARRANTY
86            vKeterangan = " WAR." & text(4) & " " & text(20) & " [ WO :" & Trim(text(0)) & " - " & text(2) & " ]"
87            vKodeLgn = Trim(gKdPlgHo)
88        ElseIf woStat = "SBI" Then
              ' SBI
89            vKeterangan = " SBI." & text(4) & " " & text(20) & " [ WO :" & Trim(text(0)) & " - " & text(2) & " ]"
90            vKodeLgn = Trim(gKdPlgDealer)
91        Else
92            vKeterangan = text(4) & " " & text(20) & " [ WO :" & Trim(text(0)) & " ] "
93            vKodeLgn = text(2)
94        End If
95        If Trim(vKodeLgn) = "" Then
96            pesan = pesan + "Kode Pelanggan kosong !"
97            boleh = False
98        Else
99            Set rsSearch = objO.FetchDataACC("select KodeLgn from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(vKodeLgn) & "'")
100               If rsSearch.BOF Then
101                   pesan = pesan + "Kode Pelanggan tidak ada ! "
102                   boleh = False
103           End If
104       End If
          
          '20/12/2012  Cust Non Axtif
          Dim rsNon As ADODB.Recordset
105       Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & Trim(vKodeLgn) & "' and nonaktifflag = 0")
106       If rsNon.RecordCount = 0 Then
107           boleh = False
108           pesan = "Kode Langgan : " & Trim(vKodeLgn) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Rubah Kode Langgan"
109       End If
          
          ''' KODE PELANGGAN UNTUK DOUBLE INVOICE
110       If nDuoInv Then
111           If gKdPlgDealer2 = "" And subJob <> 0 Then
112               boleh = False
113               pesan = pesan + "Akan tercetak dua Invoice, Kode Pelanggan untuk Dealer belum ada." & vbCr & "Hubungi Admin untuk menambah kode pelanggan dealer penjual" & vbCr
114           End If
115           If Trim(gKdPlgDealer2) <> "" Then
116               Set rsSearch = objO.FetchDataACC("select KodeLgn from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(gKdPlgDealer2) & "'")
117                   If rsSearch.BOF Then
118                       pesan = pesan + "Kode Pelanggan Dealer penjual tidak ada ! "
119                       boleh = False
120               End If
                  '20/12/2012  Cust Non Axtif
121               Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & Trim(vKodeLgn) & "' and nonaktifflag = 0")
122               If rsNon.RecordCount = 0 Then
123                   boleh = False
124                   pesan = "Kode Langgan : " & Trim(vKodeLgn) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Rubah Kode Langgan"
125               End If
126           End If
127       End If

128        If gFiskal <> Format(Month(isnow), "00") Then
129           pesan = "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
130           boleh = False
131       End If
          
          'pengambilan notelp
          Dim noTelepon As String
          Dim namaTelepon As String
          
          Dim nTuj As Boolean
          Dim nPesan As String
132       nTuj = False
133       If Log.formatSmsBilling <> "" Then
134           nPesan = Log.formatSmsBilling
135       Else
               ''2016-03-29 || 2. Jika pada field Bengkelxx..t_profil..smsBilling kosong, maka ketika proses sms di skip. Data tetap dicatat di tabel Bengkelxx..SMSNotif.
               
              ''nPesan = "Yth. Bp/Ibu #NM. Kendaraan Anda telah selesai diservis, estimasi biaya Rp. #RP untuk konfirmasi silahkan hubungi SA kami. - " & Log.LogNamaNas
136       End If
          
137       Set rsTemp = objO.FetchData("SELECT ANCUST,TELPANCUST FROM JOBDTA WHERE NO_WO = '" & Trim(text(0)) & "'")
138       If rsTemp.RecordCount > 0 Then
139           rsTemp.MoveFirst
140           noTelepon = Trim(rsTemp("telpancust"))
141           namaTelepon = Trim(rsTemp("ANCUST"))
              
142           If noTelepon <> "" Then
143               If InStr(noTelepon, "/") <> 0 Then '-------------------------------------------- Jika "081/082"
144                   If Trim(Left(noTelepon, 2)) = "08" Or Trim(Left(noTelepon, 2)) = "02" Then '---- Jika "081"="08" Or "081"="02"
145                       noTelepon = Trim(Left(noTelepon, InStr(noTelepon, "/") - 1)) '------------------ "081"
146                   Else '------------------------------------------------------------------
147                       If Trim(Mid(noTelepon, InStr(noTelepon, "/") + 1, 2)) = "08" Or Trim(Mid(noTelepon, InStr(noTelepon, "/") + 1, 2)) = "02" Then
148                           noTelepon = Trim(Mid(noTelepon, InStr(noTelepon, "/") + 1)) '--------------- "082"
149                       Else
150                           nTuj = True
151                       End If
152                   End If
153               Else '-------------------- Jika <> "081/082"
154                   If InStr(noTelepon, " 0") <> 0 Then '--------------------------------------- Jika "ABC 081"
155                       If Trim(Mid(noTelepon, InStr(noTelepon, " 0") + 1, 2)) = "08" Or Trim(Mid(noTelepon, InStr(noTelepon, " 0") + 1, 2)) = "02" Then
156                           NoTuj = Trim(Mid(noTelepon, InStr(noTelepon, " 0") + 1)) '--------------------- "081"
157                       Else
158                           nTuj = True
159                       End If
160                   Else
161                       If InStr(noTelepon, ":0") <> 0 Then '----------------------------------- Jika "ABC:081"
162                           If Trim(Mid(noTelepon, InStr(noTelepon, ":0") + 1, 2)) = "08" Or Trim(Mid(noTelepon, InStr(noTelepon, ":0") + 1, 2)) = "02" Then
163                               noTelepon = Trim(Mid(noTelepon, InStr(noTelepon, ":0") + 1)) '----------------- "081"
164                           Else
165                               nTuj = True
166                           End If
167                       Else
168                           If InStr(noTelepon, "-0") <> 0 Then '------------------------------------ Jika "ABC-081"
169                               If Trim(Mid(noTelepon, InStr(noTelepon, "-0") + 1, 2)) = "08" Or Trim(Mid(noTelepon, InStr(noTelepon, "-0") + 1, 2)) = "02" Then
170                                   NoTuj = Trim(Mid(noTelepon, InStr(noTelepon, "-0") + 1)) '----------------- "081"
171                               Else
172                                   nTuj = True
173                               End If
174                           Else
175                               If InStr(noTelepon, "-") <> 0 Then
176                                   noTelepon = Replace(noTelepon, "-", "")
177                               Else
178                                   If InStr(noTelepon, "+") = 1 Then '---------------------------------------- Jika "+62"
179                                       If Trim(Mid(noTelepon, InStr(noTelepon, "+") + 1, 2)) = "62" Then
180                                           noTelepon = Replace(noTelepon, "+62", "0") '-------------------
181                                       Else
182                                           nTuj = True
183                                       End If
184                                   Else
185                                       nTuj = True
186                                   End If
187                               End If
188                           End If
189                       End If
190                   End If
191               End If
192           End If
193       End If
          
194       If Mid(noTelepon, 1, 2) <> "08" And Mid(noTelepon, 1, 2) <> "02" Or Len(Trim(noTelepon)) < 8 Or Trim(noTelepon) = "" Then
              'pesan = "No Telepon user panggilan tidak valid !"
              'boleh = False
              'invalid number
195           nTuj = False
196       End If
          
          
          
197       If boleh Then
            '  nKodeSi = vKodeLgn 'Trim(text(2)) 'vKodeLgn
198         isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(vKodeLgn)  'cex dari ARMST apakah AFL / not AFL
199         If isExAflAR = True Then  'AFL
200             isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(vKodeLgn) 'cex dari SIJURNALNEW apakah AFL terdaftar
201             If isExAflSI = False Then
                    
202                 ErrMsg.judul = ""
203                 ErrMsg.message = "Data Customer '" & vKodeLgn & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting"
204                 ErrMsg.Show
205                 Exit Sub
206             End If
207         Else
                Dim isCust As String
208             isCust = Trim(text(40))
209             If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
210                 ErrMsg.judul = ""
211                 ErrMsg.message = "Data Customer '" & vKodeLgn & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting"
212                 ErrMsg.Show
213                 Exit Sub
214             End If
215             If isCust = "INT" Then
216                 isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(vKodeLgn) 'cex dari SIJURNALNEW apakah INT terdaftar
217                 If isExIntSI = False Then
218                     ErrMsg.judul = ""
219                     ErrMsg.message = "Data Customer '" & vKodeLgn & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting"
220                     ErrMsg.Show
221                     Exit Sub
222                 End If
223             End If
224         End If
225           If opWO(2).value = True Then
226               text(42) = ""
227               text(49) = ""
228               frmSimpanWarranty.Visible = True
229               Exit Sub
230           End If
231           Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 0)
232           Call gTglServerBilling
              
233           nRangeSyarat = 0
234           nRangeSyarat2 = 0
              ' KODE PELANGGAN INVOICE PERTAMA
235           Call RangeTempo(vKodeLgn, 0)
              '' KODE PELANGGAN INVOICE KEDUA
236           If Trim(gKdPlgDealer2) <> "" Then Call RangeTempo(gKdPlgDealer2, 1)
              
237           erTranAll = False
              ' BACA JOBDTA
238           Set rsJobDta = objO.FetchData("SELECT TOP 1 * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS =  '" & Trim(text(17)) & "'")
239           Set rsCustdat = objO.FetchData("SELECT TOP 1 * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(17)) & "'")
240           Set rsInvSublet = objO.FetchData("SELECT PKP, NAMABARANG, NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, PRICE*QTY AS HARGAJUAL, NONJURNAL FROM JOBSUBLET a WITH (NOLOCK) INNER JOIN T_SUPPLIER B WITH (NOLOCK) ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(Trim(text(0))) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
              '"SELECT PKP, NAMABARANG, NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, PRICE*QTY AS HARGAJUAL, NONJURNAL FROM JOBSUBLET a INNER JOIN T_SUPPLIER B ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(Trim(text(0))) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''"
                      
241           Call objHT.OpenTransactionAll
              '5/11/2012
242           Call objHO.AddNew_LogHO(Trim(text(0)), "BATCH", "BATCH", "-")
243           ExecMRSDo = False
244           ExecMRS = ""
              
245           If erTranAll = False Then Call simpanInvoice(Trim(text(0)))
              
246           If nDuoInv And erTranAll = False Then
247               Call simpanInvoice2(Trim(text(0)))
248           End If
              
              
              'TANGGAL 21-01-2015 RUDI
249           If erTranAll = False Then Call objHT.InsertData_Bengkel("UPDATE CHASSINGFORM SET TANGGALINVOICE = GETDATE() WHERE NOWO = '" & Trim(text(0)) & "'")
              
250           If erTranAll = False Then Call simpanHistory(Trim(text(0)))
251           If Trim(text(34)) <> "" And tipeByr = "T" And erTranAll = False Then
252               Call simpanTrnMOU(Trim(text(0)))
253           End If
              
              'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
254           If OnNewFP = 0 Then
                  'Set rsFaktur = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
255               Set rsFaktur = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
256           Else
                  'Set rsFaktur = objO.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
257               Set rsFaktur = objHT.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
258           End If
259           KodeCabangFaktur = Trim(rsFaktur("KODEDEALER"))
                  
260           If gIERC Then
                  'If CDbl(text(29)) <> 0 Then
                  
261               If erTranAll = False Then Call simpanFaktur(nDuoInv)
                  
                  'End If
      '            If nDuoInv Then
      '                Call simpanFaktur2
      '            End If
262               If nDuoInv Then '
                      '' 2015-10-16 || jika npwp dealer = dealer unit maka
263                   If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) >= 0 Then
264                       If Trim(gNpwp) <> Trim(gInvNPWP2) Then
265                           If erTranAll = False Then Call simpanFaktur2
266                       End If
267                   Else
268                       If erTranAll = False Then Call simpanFaktur2
269                   End If
270               End If
271           End If
              
              
              ' SIMPAN ACC
              
              'PERUBAHAN PENYERAGAMAN AKUN
              'TANGGAL 04-09-2014
              
              Dim hasilSimpanACC As Boolean
272           If erTranAll = True Then hasilSimpanACC = False
              
              ''2017-10-26 || menambahkan user jika pada nasmoco app belum ada
273           If Log.loyalityCustomer = 1 Then
                  Dim emailCustomer As String
                  
274               emailCustomer = objHT.FetchDataBengkel_String("SELECT USEEMAIL FROM CUSTDAT WHERE NO_CHASSIS = '" & Trim(text(17)) & "'")
275               If emailCustomer <> "" Then
276                   Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM RPLNASAPP..MASTERUSER WHERE ALAMATEMAIL = '" & emailCustomer & "' ")
277                   If rsTemp.EOF Then
                          ''jika data tidak ada akan mencari kodelgn tersebut sudah ada
                          
278                       Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM RPLNASAPP..MASTERUSER WHERE isnull(idUser,'') = '" & Trim(text(2)) & "@apps.nasmoco" & "'")
279                       If rsTemp.EOF Then
                              ''proses menambahkan unit dan new customer pada nasmoco app
                              
280                           sSQL = "SELECT a.FS_KD_PLG,'' as NOKTP, a.NPWP ,a.CUSTOMER,'01/01/1900' as TglLahir,a.USEAGAMA as agama,"
281                           sSQL = sSQL & " a.USEJOBDSC as pekerjaan,'' as propinsi,a.KOTA,a.KECAMATAN,'' as kelurahan,a.KODEPOS,"
282                           sSQL = sSQL & " '' as rt,'' as rw,a.GSM ,a.USEFS_NM_HOBI as hobi,a.CUST_ADRES,a.USEEMAIL,"
283                           sSQL = sSQL & " NO_CHASSIS,FS_KD_DEALER,FS_NM_DEALER,ddate,NO_POLISI,MDL,MODEL,FMDL,no_Engine,FS_TH_MOBIL,COLOR,'01/01/1900' as tglawalpolis,'01/01/1900' as tglakhirpolis,'' as namapolis"
284                           sSQL = sSQL & " FROM custdat a WITH (NOLOCK)"
285                           sSQL = sSQL & " WHERE a.no_chassis = '" & Trim(text(17)) & "'"
286                           Set rsTemp = objHT.FetchDataBengkel(sSQL)
                              
287                           If Not rsTemp.EOF Then
                                  Dim nama1 As String
                                  Dim nama2 As String
                                  Dim nama3 As String
                                  Dim namaCust() As String
                                  
288                               namaCust() = Split(Trim(rsTemp!CUSTOMER), " ")
289                               For i = 0 To UBound(namaCust)
290                                   If i = 0 Then
291                                       nama1 = UCase(namaCust(i))
292                                   ElseIf i = 1 Then
293                                       nama2 = UCase(namaCust(i))
294                                   ElseIf i >= 2 Then
295                                       nama3 = nama3 & " " & UCase(namaCust(i))
296                                   End If
297                               Next
                                  
298                               sSQL = "INSERT INTO [RplNasApp]..[MasterUser]([AlamatEmail],[Nama],[Password],[Telp],[KodeAktivasi],[Device],[Verifikasi],[TglEntry],[TglUpdVerivikasi],[VerifikasiEmail],[TglVerifikasiEmail],[idUser])"
299                               sSQL = sSQL & " VALUES ('" & Trim(text(2)) & "@apps.nasmoco" & "','" & Trim(rsTemp!CUSTOMER) & "','" & Int(Rnd * 99999999) & "','','','','1',getdate(),getdate(),'0',getdate(),'" & Trim(text(2)) & "@apps.nasmoco" & "')"
300                               Call objHT.InsertData_Bengkel(sSQL)
                                  
301                               sSQL = "INSERT INTO [RplNasApp]..NewCustomer(topVhc,BatasVhc,topGrp,BatasGrp,topBp,BatasBp,topPart,BatasPart,nonPPn,fpgab,typeCustomer,KodeDealer,IDCust,noktp , nonpwp ,title , namacust1 , namacust2 ,namacust3 , tgllahir ,agama ,pekerjaan , propinsi ,kota ,kecamatan , kelurahan ,kodepos , rt ,rw ,idtelp ,telp,hp1,hobi, alamat,email,userentry,tglentry)"
302                               sSQL = sSQL & " VALUES (0,0,0,0,0,0,0,0,0,0,'Perorangan','" & gKdCbg & "','" & Nul(rsTemp!FS_KD_PLG) & "','" & Nul(rsTemp!noktp) & "','" & Nul(rsTemp!npwp) & "','','" & Nul(nama1) & "','" & Nul(nama2) & "','" & Nul(nama3) & "',"
303                               sSQL = sSQL & " '" & Nul(rsTemp!tgllahir) & "','" & Nul(rsTemp!Agama) & "','" & Nul(rsTemp!Pekerjaan) & "','" & Nul(rsTemp!Propinsi) & "','" & Nul(rsTemp!Kota) & "','" & Nul(rsTemp!kecamatan) & "','" & Nul(rsTemp!kelurahan) & "',"
304                               sSQL = sSQL & " '" & Nul(rsTemp!KodePos) & "','" & Nul(rsTemp!rt) & "','" & Nul(rsTemp!rw) & "','','" & Nul(rsTemp!GSM) & "','" & Nul(rsTemp!GSM) & "','" & Nul(rsTemp!hobi) & "','" & Nul(rsTemp!cust_adres) & "','" & Trim(text(2)) & "@apps.nasmoco" & "','" & gSA & "',getdate())"
305                               Call objHT.InsertData_Bengkel(sSQL)
                                  
306                               sSQL = "INSERT INTO [RplNasApp]..UnitCustomer(ddate,fmdl,noEngine,IDCustService,NoRangka,KdDlrPenjual,nmDlrPenjual,noPolisi,mdl,thmobil,userentry,tglentry)"
307                               sSQL = sSQL & " VALUES ('" & Nul(rsTemp!dDate) & "','" & Nul(rsTemp!Fmdl) & "','" & Nul(rsTemp!No_Engine) & "','" & Nul(rsTemp!FS_KD_PLG) & "','" & Nul(rsTemp!No_Chassis) & "','" & Nul(rsTemp!FS_KD_DEALER) & "','" & Nul(rsTemp!FS_NM_DEALER) & "','" & Nul(rsTemp!NO_POLISI) & "','" & Nul(rsTemp!MDL) & "','" & Nul(rsTemp!FS_TH_MOBIL) & "','" & gSA & "',getdate())"
308                               Call objHT.InsertData_Bengkel(sSQL)
                                    
309                           End If
                              
310                       End If
311                   End If
312               End If
313           End If
              
314           If erTranAll = False Then
315               If Log.isAkun = 1 Then
316                   hasilSimpanACC = simpanACC2(Trim(text(0)))
317               Else
318                   hasilSimpanACC = simpanACCNonAkunPenyeragaman(Trim(text(0)))
319               End If
320           End If
321           If hasilSimpanACC Then
322               If ExecMRSDo = True Then
323                   Call objO.UpdateData(ExecMRS)
324               End If
                  Dim bolehSms As Boolean
                  
325               bolehSms = False
      '            On Error GoTo Ho_Error
      '            Call objHO.OpenTransactionLog_NRM
                  
                  Dim rsQue As ADODB.Recordset
326               Set rsQue = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
327               If Not rsQue.EOF Then
328                   If Trim(rsQue("THS")) = "0" And Trim(rsQue("PDS")) = "0" And Trim(rsQue("TPSVC")) <> "W" Then
                          ' PANGGIL CUSTOMER
329                       bolehSms = True
330                       Call finishService(Trim(text(16)), Trim(text(14)), Trim(namaModel(text(14))), namaPembawa(Trim(text(0))), Format(isNowBilling, "MM/DD/YYYY"), Format(isNowBilling, "HH:NN"), 1)
331                   End If
332               End If
                  
                  '------ 2014-12-18 RUDI KRISTANTO ---------------
                  'lakukan pengiriman ke Ho untuk hasil FI next job
                  
333               Set rsFI = objO.FetchData("SELECT * FROM INSPECTION WHERE NOWO = '" & Trim(text(0)) & "'")
334               If rsFI.RecordCount > 0 Then
      ''                If Get_ConnectionProfil_Bengkel99 Then
      ''                    '' 2015-09-15 || ganti inspection
      ''                    Call objHO.AddNew_WO("INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Format(Trim(rsFI("TGLENTRY")), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')")
      ''                Else
      ''                    strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Format(Trim(rsFI("TGLENTRY")), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
      ''                    strSql = Replace(strSql, "'", "~")
      ''                     Dim strLog As String
      ''                    strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
      ''                   "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
      ''                    Call objO.InsertData(strLog)
      ''                End If

                      Dim strLog As String
335                   If Log.LogHOConn = 1 Then
336                       If ConHO = False Then
337                           strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
338                           strSql = Replace(strSql, "'", "~")
339                           strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                       "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
340                           Call objO.InsertData(strLog)
341                       Else
342                           If Get_ConnectionProfil_Bengkel99 = False Then
343                               strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
344                               strSql = Replace(strSql, "'", "~")
345                               strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                           "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
346                               Call objO.InsertData(strLog)
347                           Else
348                               Call objO.InsertDataHO("INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Format(Trim(rsFI("TGLENTRY")), "YYYY-MM-DD HH:MM:SS") & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')")
349                           End If
350                       End If
351                   Else
352                       strSql = "INSERT INTO INSPECTION ([NoWo],[TglEntry],[UserEntry],[HasilFI],[HasilTestDrive],[SNJ1],[SNJ2],[SNJ3],[SNJ4],[SNJ5],[SNJ6],[SNJ7],[DEALER]) VALUES ('" & Trim(rsFI("NOWO")) & "','" & Trim(rsFI("TGLENTRY")) & "','" & Trim(rsFI("USERENTRY")) & "','" & Trim(rsFI("HASILFI")) & "','" & Trim(rsFI("HASILTESTDRIVE")) & "','" & Trim(rsFI("SNJ1")) & "','" & Trim(rsFI("SNJ2")) & "','" & Trim(rsFI("SNJ3")) & "','" & Trim(rsFI("SNJ4")) & "','" & Trim(rsFI("SNJ5")) & "','" & Trim(rsFI("SNJ6")) & "','" & Trim(rsFI("SNJ7")) & "','" & Log.KodeDealer & "')"
353                       strSql = Replace(strSql, "'", "~")
354                       strLog = "Insert Into LogHO (Tipe,Kode, Status, LogHO, Transfer, Query) " & _
                                   "Values ( 'INS','" & Trim(rsFI("NOWO")) & "', 'INSERT', 0,'01/01/1900', '" & strSql & "')"
355                       Call objO.InsertData(strLog)
356                   End If
357               End If

                      ' SIMPAN KE HO
                      'Call simpanHO(Trim(text(0)))
                      Dim nNoFak As String
358                   nNoFak = objO.FetchData_String("Select NoFAK FROM History WITH (NOLOCK) Where NO_WO = '" & Trim(text(0)) & "'")
              
359                       If Log.LogHOConn = 1 Then
360                           If ConHO = False Then
361                               Call Saving_Log_HO("Insert", sSQLi, "1")
362                               If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
363                               Call Saving_Log_HO("Insert", nSQLo(0), "i0")
364                               If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
365                               If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
366                               If nSQLo(3) <> "" Then
367                                   If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
368                               End If
                                  'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
369                               Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")

                                  'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
370                               Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
                                  
371                           Else
372                               If Get_ConnectionProfil_Bengkel99 = False Then
373                                   Call Saving_Log_HO("Insert", sSQLi, "1")
374                                   If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
375                                   Call Saving_Log_HO("Insert", nSQLo(0), "i0")
376                                   If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
377                                   If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
378                                   If nSQLo(3) <> "" Then
379                                       If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
380                                   End If
                                      'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
381                                   Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
                                      
                                      'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
382                                   Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
                                      
383                               Else
                                      
384                                   errHO = False
                                      
385                                   Call objHO.OpenTransactionLog_NRM
386                                   Call objHO.AddNew_Billing(sSQLi)
387                                   If errHO = False Then
388                                       If nDuoInv Then Call objHO.AddNew_Billing(sSQLi2)
389                                   End If
390                                   If errHO = False Then
391                                       Call objHO.AddNew_Part(nSQLo(0))
392                                   End If
393                                   If errHO = False Then
394                                       If nDuoInv Then Call objHO.AddNew_Part(nSQLo(1))
395                                   End If
396                                   If errHO = False Then
397                                       If nSQLo(2) <> "" Then Call objHO.AddNew_Part(nSQLo(2))
398                                   End If
399                                   If errHO = False Then
400                                       If nSQLo(3) <> "" Then
401                                           If nDuoInv Then Call objHO.AddNew_Part(nSQLo(3))
402                                       End If
403                                   End If
                                      Dim rsPr As ADODB.Recordset
404                                   If errHO = False Then
                                          
405                                       Set rsPr = objO.FetchData("Select * FROm T_Part WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "'")
406                                       If rsPr.RecordCount <> 0 Then
407                                           rsPr.MoveFirst
408                                           Do Until rsPr.EOF
409                                               Call objHO.AddNew_Part("Insert Into T_PARTnew Values ('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "'," & Val(rsPr.Fields(3).value) & "," & RepRp(rsPr.Fields(4).value) & "," & IIf(rsPr.Fields(5).value = False, 0, 1) & ",'" & Trim(rsPr.Fields(6).value) & "'," & Val(rsPr.Fields(7).value) & ",'','01/01/3000','" & Log.LogBengkel & "')")
410                                               rsPr.MoveNext
411                                           Loop
412                                       End If
413                                   End If
414                                   If errHO = False Then
415                                       Set rsPr = objO.FetchData("Select * FROm Hist_Sublet WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "'")
416                                       If rsPr.RecordCount <> 0 Then
417                                           rsPr.MoveFirst
418                                           Do Until rsPr.EOF
419                                               Call objHO.AddNew_Part("Insert Into Hist_SUBLETnew Values ('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "','" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "'," & Val(rsPr.Fields(8).value) & "," & RepRp(rsPr.Fields(9).value) & ",'" & Trim(rsPr.Fields(10).value) & "','" & Trim(rsPr.Fields(11).value) & "','" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & Log.LogBengkel & "')")
420                                               rsPr.MoveNext
421                                           Loop
422                                       End If
423                                   End If
424                                   If errHO = False Then
425                                       Set rsPr = objO.FetchData("Select * FROm History WITH (NOLOCK) Where No_WO = '" & Trim(text(0)) & "'")
426                                       If rsPr.RecordCount <> 0 Then
427                                           rsPr.MoveFirst
428                                           Do Until rsPr.EOF
429                                               Call objHO.AddNew_Part("Insert Into Historynew Values " & _
                                                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "','" & TglNul2(rsPr.Fields(9).value) & "'," & RepRp(rsPr.Fields(10).value) & "," & RepRp(rsPr.Fields(11).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(12).value) & "," & RepRp(rsPr.Fields(13).value) & "," & RepRp(rsPr.Fields(14).value) & "," & RepRp(rsPr.Fields(15).value) & "," & RepRp(rsPr.Fields(16).value) & "," & RepRp(rsPr.Fields(17).value) & " , " & _
                                                  " " & RepRp(rsPr.Fields(18).value) & "," & RepRp(rsPr.Fields(19).value) & ",'" & Trim(rsPr.Fields(20).value) & "','" & Trim(rsPr.Fields(21).value) & "','" & Trim(rsPr.Fields(22).value) & "','" & Trim(rsPr.Fields(23).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(24).value) & "','" & Trim(rsPr.Fields(25).value) & "','" & Trim(rsPr.Fields(26).value) & "','" & Trim(rsPr.Fields(27).value) & "','" & Trim(rsPr.Fields(28).value) & "','" & Trim(rsPr.Fields(29).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(30).value) & "','" & Trim(rsPr.Fields(31).value) & "','" & Trim(rsPr.Fields(32).value) & "','" & Trim(rsPr.Fields(33).value) & "','" & Trim(rsPr.Fields(34).value) & "','" & Trim(rsPr.Fields(35).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(36).value) & "','" & Trim(rsPr.Fields(37).value) & "','" & Trim(rsPr.Fields(38).value) & "','" & Trim(rsPr.Fields(39).value) & "','" & Trim(rsPr.Fields(40).value) & "','" & Trim(rsPr.Fields(41).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(42).value) & "','" & Trim(rsPr.Fields(43).value) & "','" & Trim(rsPr.Fields(44).value) & "','" & Trim(rsPr.Fields(45).value) & "','" & Trim(rsPr.Fields(46).value) & "','" & Trim(rsPr.Fields(47).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(48).value) & "'," & Val(rsPr.Fields(49).value) & "," & Val(rsPr.Fields(50).value) & ",'" & Trim(rsPr.Fields(51).value) & "'," & Val(rsPr.Fields(52).value) & ",'" & Trim(rsPr.Fields(53).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(54).value) & "','" & Trim(rsPr.Fields(55).value) & "','" & Trim(rsPr.Fields(56).value) & "','" & Trim(rsPr.Fields(57).value) & "','" & Trim(rsPr.Fields(58).value) & "','" & Trim(rsPr.Fields(59).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(60).value) & "'," & Val(rsPr.Fields(61).value) & ",'" & Trim(rsPr.Fields(62).value) & "'," & Val(rsPr.Fields(63).value) & ",'" & Trim(rsPr.Fields(64).value) & "','" & Format(rsPr.Fields(65).value, "MM/DD/YYYY") & "', " & _
                                                  "'" & Trim(rsPr.Fields(66).value) & "','" & Trim(rsPr.Fields(67).value) & "','" & Format(rsPr.Fields(68).value, "MM/DD/YYYY") & "','" & Log.LogBengkel & "')")
                                                  
430                                               rsPr.MoveNext
431                                           Loop
432                                       End If
433                                   End If
434                                   If errHO = False Then
435                                       Call objHO.Remove_WO("Delete FROM JOBDTA  Where No_Wo = '" & Trim(text(0)) & "'")
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
436                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
437                                   End If
438                                   If errHO = False Then
439                                       Set rsPr = objO.FetchData("Select * from JOBDTA WITH (NOLOCK) Where No_Wo = '" & Trim(text(0)) & "'")
440                                       Call objHO.AddNew_Part("Insert Into JOBDTA Values " & _
                                                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & Trim(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "','" & Trim(rsPr.Fields(5).value) & "','" & Trim(rsPr.Fields(6).value) & "','" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "','" & Trim(rsPr.Fields(9).value) & "','" & Trim(rsPr.Fields(10).value) & "','" & Trim(rsPr.Fields(11).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & Trim(rsPr.Fields(15).value) & "','" & Trim(rsPr.Fields(16).value) & "'," & RepRp(rsPr.Fields(17).value) & "," & RepRp(rsPr.Fields(18).value) & "," & RepRp(rsPr.Fields(19).value) & "," & RepRp(rsPr.Fields(20).value) & ",'" & Trim(rsPr.Fields(21).value) & "','" & Trim(rsPr.Fields(22).value) & "','" & Trim(rsPr.Fields(23).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(24).value) & "','" & TglNul2(rsPr.Fields(25).value) & "','" & Trim(rsPr.Fields(26).value) & "','" & Trim(rsPr.Fields(27).value) & "','" & Trim(rsPr.Fields(28).value) & "','" & Trim(rsPr.Fields(29).value) & "','" & Trim(rsPr.Fields(30).value) & "','" & Trim(rsPr.Fields(31).value) & "'," & RepRp(rsPr.Fields(32).value) & "," & RepRp(rsPr.Fields(33).value) & "," & RepRp(rsPr.Fields(34).value) & "," & RepRp(rsPr.Fields(35).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(36).value) & "," & RepRp(rsPr.Fields(37).value) & "," & RepRp(rsPr.Fields(38).value) & "," & RepRp(rsPr.Fields(39).value) & ",'" & Trim(rsPr.Fields(40).value) & "','" & Trim(rsPr.Fields(41).value) & "','" & Trim(rsPr.Fields(42).value) & "','" & Trim(rsPr.Fields(43).value) & "','" & Trim(rsPr.Fields(44).value) & "','" & Trim(rsPr.Fields(45).value) & "','" & Trim(rsPr.Fields(46).value) & "','" & Trim(rsPr.Fields(47).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(48).value) & "','" & Trim(rsPr.Fields(49).value) & "','" & Trim(rsPr.Fields(50).value) & "','" & Trim(rsPr.Fields(51).value) & "','" & Trim(rsPr.Fields(52).value) & "','" & Trim(rsPr.Fields(53).value) & "','" & Trim(rsPr.Fields(54).value) & "','" & Trim(rsPr.Fields(55).value) & "','" & Trim(rsPr.Fields(56).value) & "'," & RepRp(rsPr.Fields(57).value) & "," & RepRp(rsPr.Fields(58).value) & "," & RepRp(rsPr.Fields(59).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(60).value) & "," & RepRp(rsPr.Fields(61).value) & "," & RepRp(rsPr.Fields(62).value) & "," & RepRp(rsPr.Fields(63).value) & "," & RepRp(rsPr.Fields(64).value) & ",'" & Trim(rsPr.Fields(65).value) & "','" & Trim(rsPr.Fields(66).value) & "','" & Trim(rsPr.Fields(67).value) & "','" & Trim(rsPr.Fields(68).value) & "','" & Trim(rsPr.Fields(69).value) & "','" & Trim(rsPr.Fields(70).value) & "','" & Trim(rsPr.Fields(71).value) & "', " & _
                                                  "'" & TglNul2(rsPr.Fields(72).value) & "','" & Trim(rsPr.Fields(73).value) & "','" & Trim(rsPr.Fields(74).value) & "','" & Trim(rsPr.Fields(75).value) & "','" & Trim(rsPr.Fields(76).value) & "','" & Trim(rsPr.Fields(77).value) & "','" & Trim(rsPr.Fields(78).value) & "','" & Trim(rsPr.Fields(79).value) & "','" & Trim(rsPr.Fields(80).value) & "','" & Trim(rsPr.Fields(81).value) & "','" & Trim(rsPr.Fields(82).value) & "','" & Trim(rsPr.Fields(83).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(84).value) & "','" & Trim(rsPr.Fields(85).value) & "','" & Trim(rsPr.Fields(86).value) & "','" & Trim(rsPr.Fields(87).value) & "','" & Trim(rsPr.Fields(88).value) & "','" & Trim(rsPr.Fields(89).value) & "','" & Trim(rsPr.Fields(90).value) & "','" & Trim(rsPr.Fields(91).value) & "','" & Trim(rsPr.Fields(92).value) & "','" & Trim(rsPr.Fields(93).value) & "','" & Trim(rsPr.Fields(94).value) & "','" & Trim(rsPr.Fields(95).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(96).value) & "','" & Trim(rsPr.Fields(97).value) & "','" & Trim(rsPr.Fields(98).value) & "','" & Trim(rsPr.Fields(99).value) & "','" & Trim(rsPr.Fields(100).value) & "','" & Trim(rsPr.Fields(101).value) & "'," & RepRp(rsPr.Fields(102).value) & "," & RepRp(rsPr.Fields(103).value) & "," & RepRp(rsPr.Fields(104).value) & "," & RepRp(rsPr.Fields(105).value) & "," & RepRp(rsPr.Fields(106).value) & "," & RepRp(rsPr.Fields(107).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(108).value) & "," & RepRp(rsPr.Fields(109).value) & "," & RepRp(rsPr.Fields(110).value) & "," & RepRp(rsPr.Fields(111).value) & "," & RepRp(rsPr.Fields(112).value) & "," & RepRp(rsPr.Fields(113).value) & "," & RepRp(rsPr.Fields(114).value) & "," & RepRp(rsPr.Fields(115).value) & "," & RepRp(rsPr.Fields(116).value) & "," & RepRp(rsPr.Fields(117).value) & ",'" & Trim(rsPr.Fields(118).value) & "','" & Trim(rsPr.Fields(119).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(120).value) & "','" & Trim(rsPr.Fields(121).value) & "','" & Trim(rsPr.Fields(122).value) & "','" & Trim(rsPr.Fields(123).value) & "','" & Trim(rsPr.Fields(124).value) & "','" & Trim(rsPr.Fields(125).value) & "','" & Trim(rsPr.Fields(126).value) & "','" & Trim(rsPr.Fields(127).value) & "','" & Trim(rsPr.Fields(128).value) & "'," & Val(rsPr.Fields(129).value) & ",'" & Trim(rsPr.Fields(130).value) & "','" & Trim(rsPr.Fields(131).value) & "', " & _
                                                  "'" & TglNul2(rsPr.Fields(132).value) & "','" & Trim(rsPr.Fields(133).value) & "','" & Trim(rsPr.Fields(134).value) & "','" & Trim(rsPr.Fields(135).value) & "','" & Trim(rsPr.Fields(136).value) & "','" & Trim(rsPr.Fields(137).value) & "'," & RepRp(rsPr.Fields(138).value) & ",'" & Trim(rsPr.Fields(139).value) & "','" & Trim(rsPr.Fields(140).value) & "','" & Trim(rsPr.Fields(141).value) & "','" & Trim(rsPr.Fields(142).value) & "','" & Trim(rsPr.Fields(143).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(144).value) & "','" & Trim(rsPr.Fields(145).value) & "','" & TglNul2(rsPr.Fields(146).value) & "','" & Trim(rsPr.Fields(147).value) & "','" & Trim(rsPr.Fields(148).value) & "','" & TglNul2(rsPr.Fields(149).value) & "','" & Trim(rsPr.Fields(150).value) & "')")
441                                   End If
442                                   If errHO = False Then
443                                       Set rsPr = objO.FetchData("Select * from Hist_JobAdd WITH (NOLOCK) Where No_Wo = '" & Trim(text(0)) & "'")
444                                       If rsPr.RecordCount <> 0 Then
445                                           rsPr.MoveFirst
446                                           Do Until rsPr.EOF
447                                               Call objHO.AddNew_Part("Insert Into Hist_JobAddnew Values " & _
                                                  "('" & Trim(rsPr.Fields(0).value) & "','" & Trim(rsPr.Fields(1).value) & "','" & TglNul2(rsPr.Fields(2).value) & "','" & Trim(rsPr.Fields(3).value) & "','" & Trim(rsPr.Fields(4).value) & "'," & RepRp(rsPr.Fields(5).value) & ", " & _
                                                  " " & RepRp(rsPr.Fields(6).value) & ",'" & Trim(rsPr.Fields(7).value) & "','" & Trim(rsPr.Fields(8).value) & "'," & RepRp(rsPr.Fields(9).value) & "," & RepRp(rsPr.Fields(10).value) & ",'" & Trim(rsPr.Fields(11).value) & "', " & _
                                                  "'" & Trim(rsPr.Fields(12).value) & "','" & Trim(rsPr.Fields(13).value) & "','" & Trim(rsPr.Fields(14).value) & "','" & TglNul2(rsPr.Fields(15).value) & "','" & Trim(rsPr.Fields(16).value) & "','" & Log.LogBengkel & "')")
448                                               rsPr.MoveNext
449                                           Loop
450                                       End If
451                                       Call objHO.CloseTransactionLog_NRM
                                          
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
452                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
453                                   End If
454                                   If errHO = True Then
455                                       Call Saving_Log_HO("Insert", sSQLi, "1")
456                                       If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
457                                       Call Saving_Log_HO("Insert", nSQLo(0), "i0")
458                                       If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
459                                       If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
460                                       If nSQLo(3) <> "" Then
461                                           If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
462                                       End If
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
463                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
                                          'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
464                                       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
465                                   End If
466                               End If
467                           End If
468                       Else
469                           Call Saving_Log_HO("Insert", sSQLi, "1")
470                           If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
471                           Call Saving_Log_HO("Insert", nSQLo(0), "i0")
472                           If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
473                           If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
474                           If nSQLo(3) <> "" Then
475                               If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
476                           End If
                              'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
477                           Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
                              'Call objO.DeleteData("Delete FROM LOGHO Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
478                           Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
479                       End If
              
480               If gUkuranInvoice = False Then
481                   Call cetakInv1
482                   If nDuoInv Then
483                       Call cetakInv2
484                   End If
                      
485               Else
486                   Call CetakFormKecil
      '                If nDuoInv Then
      '                    Call CetakFormKecil2
      '                End If
487               End If
                  'TANGGAL 19-01-2015 - RUDI
                  'KIRIM SMS BILLING
                  
488               If nPesan <> "" Then
489                  nPesan = Replace(Replace(nPesan, "#NM", Trim(namaTelepon)), "#RP", text(29))
490               End If
                  
491               If nTuj = True And Log.smsBilling = 1 And bolehSms = True Then
492                   Set rsTemp = objO.FetchData("SELECT * FROM ANTRICUCI WHERE NOORDER = '" & Trim(text(0)) & "' AND FLAG <> 'CANCEL'")
                     ' Dim strSql As String
493                   If Not rsTemp.EOF Then
                          'jika sudah ada data, lakukan pengecekan sudah finish cuci
494                       Set rsTemp = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & Trim(text(0)) & "' AND FLAG = 'F'")
495                       If rsTemp.RecordCount > 0 Then
                              'jika sudah finish cuci
496                           strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                                  " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','1')"
497                           Call objO.InsertData(strSql)
498                           If nPesan <> "" Then
499                              strSql = "INSERT INTO " & Log.LogSMS & ".[dbo].[OUTBOX]( [Pesan], [Tujuan], [Modul], [Tanggal], [UserLog]) " & _
                                    " VALUES ('" & nPesan & "','" & noTelepon & "','NOTIF',getdate(),'" & gSA & "')"
500                              Call objO.InsertData(strSql)
501                           End If
502                       Else
                              'jika belum finish cuci
503                           strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                                  " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','0')"
504                           Call objO.InsertData(strSql)
505                       End If
506                   Else
                          'jika tidak ada data / tidak melakukan cuci insert di smsnotif dan sms
507                       strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                                  " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','1')"
508                       Call objO.InsertData(strSql)
509                       If nPesan <> "" Then
510                           strSql = "INSERT INTO " & Log.LogSMS & ".[dbo].[OUTBOX]( [Pesan], [Tujuan], [Modul], [Tanggal], [UserLog]) " & _
                                 " VALUES ('" & nPesan & "','" & noTelepon & "','NOTIF',getdate(),'" & gSA & "')"
511                           Call objO.InsertData(strSql)
512                       End If
513                   End If
514               Else
                      'jika notelp tidak valid
515                   strSql = "INSERT INTO [SMSNotif]([NoWO], [NoTujuan], [NamaCust], [TipeUnit], [TextSMS], [UserLog], [TglNotif], [TipeSMS], [StatusProses]) " & _
                          " VALUES ('" & Trim(text(0)) & "','" & noTelepon & "','" & namaTelepon & "','" & Trim(text(14)) & "','" & nPesan & "','" & gSA & "',getdate(),'FINISHSERVICE','0')"
516                   Call objO.InsertData(strSql)
                      ''If nPesan <> "" Then
517                     strSql = "INSERT INTO [" & Log.LogSMS & "].[dbo].[InvalidNumb]([Pesan], [Tujuan], [Modul], [Tanggal], [UserLog]) " & _
                              " VALUES ('" & nPesan & "','" & noTelepon & "','NOTIF',GETDATE(),'" & gSA & "')"
518                     Call objO.InsertData(strSql)
                      ''End If
519               End If
                  
                  
                  '--------------------------
520               If gPreviewCetak = True Then
521                   text(100).SetFocus
522               End If
523               If rsJobDta.State <> adStateClosed Then rsJobDta.Close
524               If rsCustdat.State <> adStateClosed Then rsCustdat.Close
525               If rsInvSublet.State <> adStateClosed Then rsInvSublet.Close
526           Else
527               Exit Sub
528           End If
              
              'Call cekTampungan
529           TblSimpan.Enabled = False
530           If cmdSublet.Visible Then
531               cetakSublet.Visible = True
532           End If
533           cmdSublet.Visible = False
534           simpan = True
535       Else
536           MsgBox pesan, vbCritical, "VALIDASI"
537           Exit Sub
538       End If
539       Exit Sub
          
Ho_Error:
540       Call objHO.CloseTransactionLog_NRM
          
          '' 2015-09-08 | simpan data kedalam log ''
541       Call Saving_Log_HO("Insert", sSQLi, "1")
542       If nDuoInv Then Call Saving_Log_HO("Insert", sSQLi2, "2")
543       Call Saving_Log_HO("Insert", nSQLo(0), "i0")
544       If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(1), "i1")
545       If nSQLo(2) <> "" Then Call Saving_Log_HO("Insert", nSQLo(2), "i2")
546       If nSQLo(3) <> "" Then
547           If nDuoInv Then Call Saving_Log_HO("Insert", nSQLo(3), "i3")
548       End If
549       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'WO' AND kode = '" & Trim(text(0)) & "'")
550       Call objO.UpdateData("update LOGHO Set LogHo = 1 , Transfer = getdate() Where Tipe = 'BATCH' AND kode = '" & Trim(text(0)) & "'")
          ''------------------------------------------''
          
551       Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
          'MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmBilling")
          
552      Exit Sub
          
          
vbError:
          'MsgBox "Data belum berhasil disimpan." & vbCr & "Ulangi penyimpanan .."
553       Call objHT.CloseTransactionAll_Err
554       Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
555       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmBilling")
          
         'On Error GoTo 0
556      Exit Sub

tblSimpan_Click_Error:

557       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmBilling")
End Sub

Private Sub Saving_Log_HO(nStatus As String, isSQl As String, n2 As String)
    If Left(n2, 1) <> "i" Then
        Call objHO.AddNew_LogHO(Trim(text(0)), nStatus, "INV" & n2, Replace(Replace(isSQl, "')", "','" & Log.LogBengkel & "')"), ") VALUES (", ",CABANG) VALUES ("))
    Else
        Call objHO.AddNew_LogHO(Trim(text(0)), nStatus, "INV" & n2, isSQl)
    End If
End Sub

Private Sub simpanInvoice(noWO As String)
      Dim rsHlp As ADODB.Recordset
      Dim nInvNew As Double
1        On Error GoTo simpanInvoice_Error

2         ClearsSql
3         If text(20) = "" Then
4             nInvNew = isiNoInv
5             text(20) = Format(nInvNew, "00000000")
6             If erTranAll = True Then
7                 text(20) = ""
8                 Exit Sub
9             End If
              
10        End If
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
11        gsSql1(0) = "NOINV"
12        gsSql2(0) = text(20) & IIf(nDuoInv, "-A", "")
13        gsSql1(1) = "CUSTOMER"
14        gsSql2(1) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, gInvCust, text(4))
15        gsSql1(2) = "CUST_ADRES"
16        gsSql2(2) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, gInvAlmt, txt5)
17        gsSql1(3) = "KOTA"
18        gsSql2(3) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, gInvKota, text(6))
19        gsSql1(4) = "TELPON"
20        gsSql2(4) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, "", text(7))
21        gsSql1(5) = "NO_POLISI"
22        gsSql2(5) = text(16)
23        gsSql1(6) = "NO_ENGINE"
24        gsSql2(6) = text(18)
25        gsSql1(7) = "NO_CHASSIS"
26        gsSql2(7) = text(17)
27        gsSql1(8) = "MDL"
28        gsSql2(8) = text(12)
29        gsSql1(9) = "MODEL"
30        gsSql2(9) = text(13)
31        gsSql1(10) = "FMDL"
32        gsSql2(10) = text(14)
33        gsSql1(11) = "COLOR"
34        gsSql2(11) = text(15)
35        gsSql1(12) = "FS_KD_DEALER"
36        gsSql2(12) = Trim(gKdCbg)
37        gsSql1(13) = "FS_NM_DEALER"
38        gsSql2(13) = Trim(gNmCbg)
39        gsSql1(14) = "PPN"
40        gsSql2(14) = 10
          
41        gsSql1(15) = "DISC"
42        gsSql2(15) = setKoma(CDbl(text(21)))
43        gsSql1(16) = "DISCT"
44        gsSql2(16) = CDbl(text(23))
          
45        gsSql1(17) = "FRONT"
46        gsSql2(17) = Trim(text(19))
47        gsSql1(18) = "MECH"
48        gsSql2(18) = text(38)
49        gsSql1(19) = "KM"
50        gsSql2(19) = text(37)
          
51        gsSql1(20) = "PART_G"
52        gsSql2(20) = RepRp(CDbl(subPartGenuine * (1 - (CDbl(text(23)) / 100)))) 'CDbl(Text(24))     'subPart    '
53        gsSql1(21) = "PART_NG"
54        gsSql2(21) = RepRp(CDbl(subPartNGenuine * (1 - (CDbl(text(23)) / 100))))
          
55        gsSql1(22) = "MTRL"
56        gsSql2(22) = RepRp(CDbl(CDbl(text(28))))
57        gsSql1(23) = "OIL_G"
58        gsSql2(23) = RepRp(CDbl(subOliGenuine * (1 - (CDbl(text(23)) / 100)))) 'CDbl(Text(25))     'subOli     '
59        gsSql1(24) = "OIL_NG"
60        gsSql2(24) = RepRp(CDbl(subOliNGenuine * (1 - (CDbl(text(23)) / 100))))
          
61        gsSql1(25) = "MET"
62        gsSql2(25) = IIf(nDuoInv, 0, CDbl(Text42.text))
63        gsSql1(26) = "SBLT"
64        gsSql2(26) = RepRp(CDbl(text(27)))   'subSublet  '

          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
65        gsSql1(27) = "LABOR"
66        gsSql2(27) = IIf(lblPDS.Visible, RepRp(CDbl(subJob)), RepRp(CDbl(text(22))))
67        gsSql1(28) = "BODY"
68        gsSql2(28) = RepRp(CDbl(text(39)))     'subBody    '
69        gsSql1(29) = "TOTAL"
          'gsSql2(29) = RepRp(Round(CDbl((subPartGenuine * (1 - (CDbl(text(23)) / 100))) + (subPartNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(text(28)) + (subOliGenuine * (1 - (CDbl(text(23)) / 100))) + (subOliNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(text42.text) + CDbl(text(27)) + IIf(lblPDS.Visible, subJob, CDbl(text(22))) + CDbl(text(39))), 0))
70        gsSql2(29) = RepRp(SpyRound(CDbl((subPartGenuine * (1 - (CDbl(text(23)) / 100))) + (subPartNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(text(28)) + (subOliGenuine * (1 - (CDbl(text(23)) / 100))) + (subOliNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(IIf(nDuoInv, 0, CDbl(Text42.text))) + CDbl(text(27)) + IIf(lblPDS.Visible, subJob, CDbl(text(22))) + CDbl(text(39))), 0))
      '    Set rsHlp = objO.FetchData("SELECT *, GETDATE() AS tglServ FROM jobdta where no_wo = '" & text(0) & "' AND no_chassis = '" & text(17) & "'")
      '    If Not rsHlp.EOF Then
71            gsSql1(30) = "[DATE]"
72            gsSql2(30) = TglYYYYMMDD(rsJobDta("DATE"))
73            gsSql1(31) = "TIME"
74            gsSql2(31) = rsJobDta("TIME")
75            gsSql1(32) = "TIME_M"
76            gsSql2(32) = rsJobDta("TIME_M")
77            gsSql1(33) = "TIME_S"
78            gsSql2(33) = rsJobDta("TIME_S")
79            gsSql1(34) = "TIME_A"
80            gsSql2(34) = rsJobDta("TIME_A")
81            gsSql1(35) = "PAID"
82            gsSql2(35) = IIf(paid <> "", paid, rsJobDta("PAID"))
              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
83            gsSql1(36) = "STATUS"
              'gsSql2(36) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, text(40), rsJobdta("STATUS"))
84            gsSql2(36) = IIf((Trim(text(33)) = "R001" And Not nDuoInv) Or (opWO(2).value = True), text(40), rsJobDta("STATUS"))
85            gsSql1(37) = "NOFAK"
86            gsSql2(37) = rsJobDta("NOFAK")
87            gsSql1(38) = "REMARK1"
88            gsSql2(38) = rsJobDta("REMARK1")
89            gsSql1(39) = "SUGGEST"
90            gsSql2(39) = Left(rsJobDta("SUGGEST"), 300)
91            gsSql1(40) = "FN_RETURN"
92            gsSql2(40) = rsJobDta("FN_RETURN")
93            gsSql1(41) = "FS_RC"
94            gsSql2(41) = Nul(rsJobDta("FS_RC"))
95            nRC = IIf(Trim(gsSql2(41)) <> "", True, False)
96            gsSql1(42) = "RPR_COND"
97            gsSql2(42) = "S" 'Nul(rsJobdta("RPR_COND"))
              
              Dim hlp As String
98            For i = 1 To 5
99                hlp = "REQ" & i
100               gsSql1(42 + i) = hlp
101               gsSql2(42 + i) = rsJobDta(hlp)
102           Next
      '    End If
          
103       gsSql1(48) = "SLIP"
104       gsSql2(48) = text(32)
105       gsSql1(49) = "INVDATE"
106       gsSql2(49) = "getdate()"
107       gsSql1(50) = "NO_WO"
108       gsSql2(50) = noWO
109       gsSql1(51) = "DELIV"
110       gsSql2(51) = "OKE"
111       gsSql1(52) = "FS_TH_MOBIL"
112       gsSql2(52) = text(10)
113       gsSql1(53) = "FS_CATATAN"
114       gsSql2(53) = text(30)
115       gsSql1(54) = "FS_LOG_ADD"
116       gsSql2(54) = gSA
          
          '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
117       gsSql1(55) = "FS_KD_PLG"
118       gsSql2(55) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, gKodePlg, text(2))
119       gsSql1(56) = "FS_KD_WIL"
120       gsSql2(56) = IIf(Trim(text(33)) = "R001" And Not nDuoInv, Left(gKodePlg, 3), text(3))
          
121       gsSql1(57) = "THS"
122       gsSql2(57) = IIf(nTHS, 1, 0)
123       gsSql1(58) = "PDS"
124       gsSql2(58) = IIf(lblPDS.Visible, 1, 0)
125       gsSql1(59) = "MOU"
126       gsSql2(59) = Trim(text(34))
127       gsSql1(60) = "TPSVC"
128       gsSql2(60) = IIf(opWO(0).value, "", IIf(opWO(1).value, "P", "W"))
129       gsSql1(61) = "KODETIPE"
130       gsSql2(61) = Trim(text(35))
131       gsSql1(62) = "INVTIME"
132       gsSql2(62) = Format(isNowBilling, "HH:NN")
133       gsSql1(63) = "NOFAKTUR"
134       gsSql2(63) = Trim(noKFaktur)
135       gsSql1(64) = "FD_TGL_PART"
136       gsSql2(64) = Trim(Format(tglKFaktur, "MM/DD/YYYY"))
137       gsSql1(65) = "FS_KREDIT"
138       gsSql2(65) = 1
        
139       gsSql1(66) = "JOB"
140       gsSql2(66) = IIf(lblServicePoint.Visible, "1", "0")
          
          ''2016-05-23 | penambahan ddate pada t_invoice
141       gsSql1(67) = "DDATE"
142       gsSql2(67) = Trim(Format(dDate, "mm/dd/yyyy"))
              
          'Set rsHlp = objO.FetchData("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE no_wo = '" & Text(0) & "' AND no_chassis = '" & Text(17) & "'")
          'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
143       Set rsHlp = objHT.FetchDataBengkel("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE no_wo = '" & text(0) & "' AND no_chassis = '" & text(17) & "'")
144       If Not rsHlp.EOF Then
145           For i = 1 To 5
146               hlp = "CMPL" & i
147               gsSql1(67 + i) = hlp
148               gsSql2(67 + i) = Trim(rsHlp(hlp))
149           Next
150       End If
          
151       If rsJob.State <> adStateClosed Then
152           If Not rsJob.BOF Then
153               i = 1
154               K = 0
155               rsJob.MoveFirst
156               Do While Not rsJob.EOF
157                   If rsJob("addJob") Then
158                       K = K + 1
159                   End If
160                   rsJob.MoveNext
161               Loop
162               rsJob.MoveFirst
163               Do While Not rsJob.EOF
164                   If Not rsJob("ADDJOB") Then
165                       gsSql1(72 + i) = "REP" & i
166                       gsSql2(72 + i) = Trim(rsJob("JOB"))
167                       gsSql1(72 + (1 * (rsJob.RecordCount - K)) + i) = "RTE" & i
168                       gsSql2(72 + (1 * (rsJob.RecordCount - K)) + i) = Replace(CDbl(rsJob("RATE")), ",", ".")
169                       gsSql1(72 + (2 * (rsJob.RecordCount - K)) + i) = "ACTRATE" & i
170                       gsSql2(72 + (2 * (rsJob.RecordCount - K)) + i) = Replace(CDbl(rsJob("ACTRT")), ",", ".")
171                       gsSql1(72 + (3 * (rsJob.RecordCount - K)) + i) = "KLPKERJA" & i
172                       gsSql2(72 + (3 * (rsJob.RecordCount - K)) + i) = Trim(rsJob("GRP"))
173                       gsSql1(72 + (4 * (rsJob.RecordCount - K)) + i) = "MEKANIK" & i
174                       gsSql2(72 + (4 * (rsJob.RecordCount - K)) + i) = Trim(rsJob("MECH"))
175                       gsSql1(72 + (5 * (rsJob.RecordCount - K)) + i) = "OFF" & i
176                       gsSql2(72 + (5 * (rsJob.RecordCount - K)) + i) = RepRp(CDbl(rsJob("HRG")))  'CDbl(rsJob("HRG"))
177                       gsSql1(72 + (6 * (rsJob.RecordCount - K)) + i) = "DISC" & i
178                       gsSql2(72 + (6 * (rsJob.RecordCount - K)) + i) = Replace(CDbl(rsJob("DISC")), ",", ".")
179                       gsSql1(72 + (7 * (rsJob.RecordCount - K)) + i) = "VALID" & i
180                       gsSql2(72 + (7 * (rsJob.RecordCount - K)) + i) = Trim(rsJob("VALID"))
181                       gsSql1(72 + (8 * (rsJob.RecordCount - K)) + i) = "FREEC" & i
182                       gsSql2(72 + (8 * (rsJob.RecordCount - K)) + i) = IIf(Trim(rsJob("FREEC")), 1, 0)
183                       i = i + 1
184                   End If
185                   rsJob.MoveNext
186               Loop
187           End If
188       End If
          
          'Set rsTemp = objO.FetchData("SELECT * FROM t_INVOICE WITH (NOLOCK) WHERE noinv = '" & Text(20) & "' AND no_wo = '" & noWO & "' AND no_chassis = '" & Text(17) & "'")
          'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
189       Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM t_INVOICE WITH (NOLOCK) WHERE noinv = '" & text(20) & "' AND no_wo = '" & noWO & "' AND no_chassis = '" & text(17) & "'")
190       If Not rsTemp.EOF Then
191           sSQL = GensSql2("T_INVOICE", True) + " WHERE no_chassis = '" & text(17) & "' AND no_WO = '" & noWO & "' AND noInv = '" & text(20) & "' AND FS_LOG_DEL = ''"
192       Else
193           sSQL = GensSql2("T_INVOICE", False)
194       End If
195       Call objHT.UpdateData_Bengkel(sSQL)
196       sSQLi = Replace(sSQL, "T_INVOICE", "T_INVOICENew")
          
          
          ' UPDATE NO URUT INVOICE
197       If nInvNew = 0 Then
198           nInvNew = isiNoInv
199       End If
      '198       Call objHT.UpdateData_Bengkel("UPDATE T_PROFIL SET FN_NO_INV = " & nInvNew + 1)
          
          
          'Set rsGetDate = objO.FetchData("SELECT Convert(VarChar(12), getdate(), 101) as Tanggal,Convert(VarChar(5), getdate(), 108) as Jam")
          'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
200       Set rsGetDate = objHT.FetchDataBengkel("SELECT Convert(VarChar(12), getdate(), 101) as Tanggal,Convert(VarChar(5), getdate(), 108) as Jam")
          
201       If Not rsGetDate.EOF Then
202           sSQLi = Replace(sSQLi, "Convert(VarChar(12), getdate(), 101)", "'" & CStr(rsGetDate.Fields(0).value) & "'")
203           sSQLi = Replace(sSQLi, "Convert(VarChar(5), getdate(), 108)", "'" & CStr(rsGetDate.Fields(1).value) & "'")
204       End If
          
          ' UPDATE DISC DI JOBADD2
205       If rsJob.State <> adStateClosed Then
206           If Not rsJob.BOF Then
207               rsJob.MoveFirst
208               Do While Not rsJob.EOF
209                   If rsJob("ADDJOB") Then
210                       Call objHT.UpdateData_Bengkel("UPDATE JOBADD2 SET OFFN = " & CDbl(rsJob("HRG")) & ", DISC = " & Replace(CDbl(rsJob("DISC")), ",", ".") & ", VALID = '" & Trim(rsJob("VALID")) & "', FREEC = '" & Trim(IIf(rsJob("FREEC"), 1, 0)) & "' WHERE NO_WO = '" & noWO & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND REP = '" & Trim(rsJob("JOB")) & "'")
211                   End If
212                   rsJob.MoveNext
213               Loop
214           End If
215       End If
          
          'UPDATE TUTUP ORDER SOHDR
216       If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(43)), "DD/MM/YYYY") Then
217           Call objHT.InsertData_ACC("UPDATE SOHDR SET TUTUPORDER = 1 WHERE NOSO = '" & Trim(text(0)) & "' ")
218       End If
          
          ' UPDATE NO INV DI JOBDTA
219       Call objHT.UpdateData_Bengkel("UPDATE JOBDTA set noinv = '" & text(20) & "', INVDATE = CONVERT(varchar(10),GETDATE(),1), RPR_COND = 'S' WHERE no_wo = '" & Trim(noWO) & "'")
          
          ' UPDATE NO INV DI DIAGDTA
220       Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET NOINV = '" & Trim(text(20)) & "' WHERE no_wo = '" & Trim(noWO) & "'")
221       Call objHT.UpdateData_Bengkel("UPDATE DIAGSUBLET SET NOINV = '" & Trim(text(20)) & "' WHERE no_wo = '" & Trim(noWO) & "'")
          
          'Update pajak
          Dim vTotalNonPPn As Double
          Dim vTPPN As Double
222       vTotalNonPPn = Round(Koma2Titik(text(27)) * (100 / 110), 0) + Round((Koma2Titik(text(22)) * (100 / 110)), 0) + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)
223       vTPPN = (Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)) - vTotalNonPPn

      '    Call objHT.UpdateData_Bengkel("update T_INVOICE set FN_PAJAK = " & Round((CDbl(Text(29)) - CDbl(Text(42))) * 10 / 110) & " " _
                              & "where NO_WO ='" & NoWo & "' and (DELIV='PRC' or DELIV='OKE') and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'")
224       Call objHT.UpdateData_Bengkel("update T_INVOICE set FN_PAJAK = " & vTPPN & " " _
                              & "where NO_WO ='" & noWO & "' and (DELIV='PRC' or DELIV='OKE') and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'")
          
      '    nSQLo(0) = "update T_INVOICEnew set FN_PAJAK = " & Round((CDbl(Text(29)) - CDbl(Text(42))) * 10 / 110) & " " _
                              & "where NO_WO ='" & NoWo & "' and (DELIV='PRC' or DELIV='OKE') and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'"
225       nSQLo(0) = "update T_INVOICEnew set FN_PAJAK = " & vTPPN & " " _
                              & "where NO_WO ='" & noWO & "' and (DELIV='PRC' or DELIV='OKE') and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'"
          
          
226       Call objHT.UpdateData_Bengkel("update CUSTDAT set TGL_REP = '" & TglYYYYMMDD(text(1)) & "', FS_KM_AKHIR = '" & text(37) & "' where NO_CHASSIS = '" & text(17) & "'")
          
          ' UPDATE TGL REP RC TERAKHIR
227       If Trim(gsSql2(41)) <> "" Then
228           Call objHT.UpdateData_Bengkel("update CUSTDAT set RCDATE = '" & TglYYYYMMDD(text(1)) & "' where NO_CHASSIS = '" & text(17) & "'")
229       End If
          
230       Set rsHlp = objO.FetchData("SELECT noSTALL FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & noWO & "|%'")
231       If Not rsHlp.EOF Then
232           Call objHT.UpdateData_Bengkel("UPDATE STALL SET FLAG = 'I' WHERE KETERANGAN LIKE '%|" & noWO & "|%'")
233       End If
          
          ' CEK VALIDASI UNTUK MRS
          Dim hCekMrs As String
          Dim rsMRS As ADODB.Recordset
      '    Set rsTemp = objO.FetchData("SELECT * FROM JOBdta WHERE NO_WO = '" & Trim(text(0)) & "' AND NO_CHASSIS = '" & text(17) & "'")
      '    If Not rsTemp.EOF Then
234           hCekMrs = cekMRS(text(17), CDbl(text(37)), Trim(text(33)), 2, , Trim(JobGrid.Columns(0)))
235           Call objHT.UpdateData_Bengkel("UPDATE JOBDTA SET SBE = '" & hCekMrs & "' WHERE NO_WO = '" & Trim(noWO) & "' AND NO_CHASSIS = '" & text(17) & "'")
236           If hCekMrs = "GRP" Then
237               Call objHT.UpdateData_Bengkel("INSERT INTO M_TOSBE SELECT DISTINCT '" & text(17) & "', '" & text(37) & "', '" & text(64) & "', CONVERT(varchar(10),GETDATE(),1), '" & gSA & "'")
238               Exit Sub
239           ElseIf hCekMrs = "SBE" Then
                  
                  'Set oRs = objO.FetchData("SELECT FS_NILAI FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC = '" & Trim(Text(33)) & "'")
                  'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
240               Set oRs = objHT.FetchDataBengkel("SELECT FS_NILAI FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC = '" & Trim(text(33)) & "'")
241               If Not oRs.EOF Then
242                   ExecMRSDo = True
243                   ExecMRS = "EXEC " & Trim(hTipePenghitunganDealer) & " '" & text(17) & "', '" & Format(text(43), "MM/DD/YYYY") & "', '" & CDbl(text(37)) & "', '" & oRs("FS_NILAI") & "', '" & IIf(nHeavy, 5000, 10000) & "', '" & Trim(text(20)) & "', '" & Trim(noWO) & "','" & text(4) & "', '" & txt5 & "','" & text(16) & "', '" & text(7) & "', '" & text(8) & "', '" & text(9) & "'"
                      'Call objHT.UpdateData_Bengkel("EXEC " & Trim(hTipePenghitunganDealer) & " '" & text(17) & "', '" & Format(text(43), "MM/DD/YYYY") & "', '" & CDbl(text(37)) & "', '" & oRs("FS_NILAI") & "', '" & IIf(nHeavy, 5000, 10000) & "', '" & Trim(text(20)) & "', '" & trim(NoWo) & "','" & text(4) & "', '" & txt5 & "','" & text(16) & "', '" & text(7) & "', '" & text(8) & "', '" & text(9) & "'")
244               End If
245           End If
      '    End If
          
          ' INPUT UNTUK TRANSAKSI PIUTANG
246       If tipeByr = "T" And gVerifikator <> "" Then
247           Call objHT.InsertData_Bengkel("INSERT INTO PIUTANGTRN VALUES ('" & Format(isNowBilling, "MM/DD/YYYY") & "', '" & Trim(noWO) & "', '" & text(20) & "', '" & text(17) & "', '" & text(16) & "', " & CDbl(text(29)) & ", 'U', '" & gVerifikator & "')")
              
      '        Set rsHlp = objO.FetchData("SELECT * FROM T_PETUGAS WITH (NOLOCK) WHERE FS_KD_PETUGAS = '" & gVerifikator & "'")
              'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
248           Set rsHlp = objHT.FetchDataBengkel("SELECT * FROM T_PETUGAS WITH (NOLOCK) WHERE FS_KD_PETUGAS = '" & gVerifikator & "'")
249           If Not rsHlp.EOF Then
250               If rsHlp("TIPE") = "KACAB" Then
251                   Call objHT.UpdateData_Bengkel("UPDATE T_PROFIL SET HUTANGKACAB = HUTANGKACAB + " & CDbl(text(29)) & "")
252               ElseIf rsHlp("TIPE") = "KABENG" Then
253                   Call objHT.UpdateData_Bengkel("UPDATE T_PROFIL SET HUTANGKABENG = HUTANGKABENG + " & CDbl(text(29)) & "")
254               End If
255           End If
               
256       End If

         'On Error GoTo 0
257      Exit Sub

simpanInvoice_Error:
258       erTranAll = True
259       ErrMsg.judul = ""
260       ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanInvoice of Form frmBilling")
261       ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanInvoice of Form frmBilling")
End Sub

Private Sub simpanInvoice2(noWO As String)
      '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
      Dim rsHlp As ADODB.Recordset
          
1        On Error GoTo simpanInvoice2_Error
            Dim fn_pajak As Double
2         ClearsSql
3         If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
4           If Trim(gNpwp) = Trim(gInvNPWP2) Then
               '' jika afiliasi
5              subJob = subJob * 10 / 11
6              fn_pajak = 0
7           Else
8              fn_pajak = Round((subJob * 1 / 11))
9           End If
10        Else
11          fn_pajak = Round((subJob * 1 / 11))
            
12        End If
13        gsSql1(0) = "NOINV"
14        gsSql2(0) = text(20) & IIf(nDuoInv, "-B", "")
15        gsSql1(1) = "CUSTOMER"
16        gsSql2(1) = gInvCust2
17        gsSql1(2) = "CUST_ADRES"
18        gsSql2(2) = gInvAlmt2
19        gsSql1(3) = "KOTA"
20        gsSql2(3) = gInvKota2
21        gsSql1(4) = "TELPON"
22        gsSql2(4) = ""
23        gsSql1(5) = "NO_POLISI"
24        gsSql2(5) = text(16)
25        gsSql1(6) = "NO_ENGINE"
26        gsSql2(6) = text(18)
27        gsSql1(7) = "NO_CHASSIS"
28        gsSql2(7) = text(17)
29        gsSql1(8) = "MDL"
30        gsSql2(8) = text(12)
31        gsSql1(9) = "MODEL"
32        gsSql2(9) = text(13)
33        gsSql1(10) = "FMDL"
34        gsSql2(10) = text(14)
35        gsSql1(11) = "COLOR"
36        gsSql2(11) = text(15)
37        gsSql1(12) = "FS_KD_DEALER"
38        gsSql2(12) = Trim(gKdCbg)
39        gsSql1(13) = "FS_NM_DEALER"
40        gsSql2(13) = Trim(gNmCbg)
41        gsSql1(14) = "PPN"
42        gsSql2(14) = 10
          
43        gsSql1(15) = "DISC"
44        gsSql2(15) = 0
45        gsSql1(16) = "DISCT"
46        gsSql2(16) = 0
          
47        gsSql1(17) = "FRONT"
48        gsSql2(17) = Trim(text(19))
49        gsSql1(18) = "MECH"
50        gsSql2(18) = text(38)
51        gsSql1(19) = "KM"
52        gsSql2(19) = text(37)
          
53        gsSql1(20) = "PART_G"
54        gsSql2(20) = 0
55        gsSql1(21) = "PART_NG"
56        gsSql2(21) = 0
          
57        gsSql1(22) = "MTRL"
58        gsSql2(22) = 0
59        gsSql1(23) = "OIL_G"
60        gsSql2(23) = 0
61        gsSql1(24) = "OIL_NG"
62        gsSql2(24) = 0
          
63        gsSql1(25) = "MET"
64        gsSql2(25) = Val(CDbl(Text42.text))
65        gsSql1(26) = "SBLT"
66        gsSql2(26) = 0
67        gsSql1(27) = "LABOR"
68        gsSql2(27) = IIf(lblPDS.Visible, Val(subJob), IIf(nDuoInv, Val(subJob), Val(CDbl(text(22)))))
69        gsSql1(28) = "BODY"
70        gsSql2(28) = 0
71        gsSql1(29) = "TOTAL"
72        gsSql2(29) = Val(gsSql2(25)) + Val(gsSql2(27))
                  
      '    Set rsHlp = objO.FetchData("SELECT *, GETDATE() AS tglServ FROM jobdta where no_wo = '" & Trim(text(0)) & "' AND no_chassis = '" & text(17) & "'")
      '    If Not rsHlp.EOF Then
73            gsSql1(30) = "[DATE]"
74            gsSql2(30) = TglYYYYMMDD(rsJobDta("DATE"))
75            gsSql1(31) = "TIME"
76            gsSql2(31) = rsJobDta("TIME")
77            gsSql1(32) = "TIME_M"
78            gsSql2(32) = rsJobDta("TIME_M")
79            gsSql1(33) = "TIME_S"
80            gsSql2(33) = rsJobDta("TIME_S")
81            gsSql1(34) = "TIME_A"
82            gsSql2(34) = rsJobDta("TIME_A")
83            gsSql1(35) = "PAID"
84            gsSql2(35) = "T"
85            gsSql1(36) = "STATUS"
86            gsSql2(36) = vTipeCustomer
87            gsSql1(37) = "NOFAK"
88            gsSql2(37) = rsJobDta("NOFAK")
89            gsSql1(38) = "REMARK1"
90            gsSql2(38) = ""
91            gsSql1(39) = "SUGGEST"
92            gsSql2(39) = ""
93            gsSql1(40) = "FN_RETURN"
94            gsSql2(40) = rsJobDta("FN_RETURN")
95            gsSql1(41) = "FS_RC"
96            gsSql2(41) = Nul(Trim(rsJobDta("FS_RC")))
97            nRC = IIf(Trim(gsSql2(41)) <> "", True, False)
98            gsSql1(42) = "RPR_COND"
99            gsSql2(42) = "S" 'Nul(rsJobdta("RPR_COND"))
              
              Dim hlp As String
100           For i = 1 To 5
101               hlp = "REQ" & i
102               gsSql1(42 + i) = hlp
103               gsSql2(42 + i) = rsJobDta(hlp)
104           Next
      '    End If
          
105       gsSql1(48) = "SLIP"
106       gsSql2(48) = ""
107       gsSql1(49) = "INVDATE"
108       gsSql2(49) = "getdate()"
109       gsSql1(50) = "NO_WO"
110       gsSql2(50) = noWO
111       gsSql1(51) = "DELIV"
112       gsSql2(51) = "OKE"
113       gsSql1(52) = "FS_TH_MOBIL"
114       gsSql2(52) = text(10)
115       gsSql1(53) = "FS_CATATAN"
116       gsSql2(53) = Trim(text(30))
117       gsSql1(54) = "FS_LOG_ADD"
118       gsSql2(54) = gSA
119       gsSql1(55) = "FS_KD_PLG"
120       gsSql2(55) = gKodePlg2
121       gsSql1(56) = "FS_KD_WIL"
122       gsSql2(56) = Left(Trim(gKodePlg2), 3)
123       gsSql1(57) = "THS"
124       gsSql2(57) = IIf(nTHS, 1, 0)
125       gsSql1(58) = "PDS"
126       gsSql2(58) = IIf(lblPDS.Visible, 1, 0)
127       gsSql1(59) = "MOU"
128       gsSql2(59) = Trim(text(34))
129       gsSql1(60) = "TPSVC"
130       gsSql2(60) = IIf(opWO(0).value, "", IIf(opWO(1).value, "P", "W"))
131       gsSql1(61) = "KODETIPE"
132       gsSql2(61) = Trim(text(35))
133       gsSql1(62) = "INVTIME"
134       gsSql2(62) = Format(isNowBilling, "HH:NN")
135       gsSql1(63) = "NOFAKTUR"
136       gsSql2(63) = ""
137       gsSql1(64) = "FS_KREDIT"
138       gsSql2(64) = 1
          
      '    Set rsHlp = objO.FetchData("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE no_wo = '" & Trim(Text(0)) & "' AND no_chassis = '" & Text(17) & "'")
          'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
139       Set rsHlp = objHT.FetchDataBengkel("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE no_wo = '" & Trim(text(0)) & "' AND no_chassis = '" & text(17) & "'")
140       If Not rsHlp.EOF Then
141           For i = 1 To 5
142               hlp = "CMPL" & i
143               gsSql1(64 + i) = hlp
144               gsSql2(64 + i) = Trim(rsHlp(hlp))
145           Next
146       End If
           
          
147       If rsJob.State <> adStateClosed Then
148           If Not rsJob.BOF Then
149               i = 1
150               K = 0
151               rsJob.MoveFirst
152               Do While Not rsJob.EOF
153                   If rsJob("addJob") Then
154                       K = K + 1
155                   End If
156                   rsJob.MoveNext
157               Loop
158               rsJob.MoveFirst
159               Do While Not rsJob.EOF
160                   If Not rsJob("ADDJOB") Then
161                       gsSql1(69 + i) = "REP" & i
162                       gsSql2(69 + i) = Trim(rsJob("JOB"))
163                       gsSql1(69 + (1 * (rsJob.RecordCount - K)) + i) = "RTE" & i
164                       gsSql2(69 + (1 * (rsJob.RecordCount - K)) + i) = Replace(CDbl(rsJob("RATE")), ",", ".")
165                       gsSql1(69 + (2 * (rsJob.RecordCount - K)) + i) = "ACTRATE" & i
166                       gsSql2(69 + (2 * (rsJob.RecordCount - K)) + i) = Replace(CDbl(rsJob("ACTRT")), ",", ".")
167                       gsSql1(69 + (3 * (rsJob.RecordCount - K)) + i) = "KLPKERJA" & i
168                       gsSql2(69 + (3 * (rsJob.RecordCount - K)) + i) = Trim(rsJob("GRP"))
169                       gsSql1(69 + (4 * (rsJob.RecordCount - K)) + i) = "MEKANIK" & i
170                       gsSql2(69 + (4 * (rsJob.RecordCount - K)) + i) = Trim(rsJob("MECH"))
171                       gsSql1(69 + (5 * (rsJob.RecordCount - K)) + i) = "OFF" & i
172                       gsSql2(69 + (5 * (rsJob.RecordCount - K)) + i) = CDbl(rsJob("HRG")) 'IIf(Trim(rsJob("FREEC")), 0, CDbl(rsJob("HRG")))
173                       gsSql1(69 + (6 * (rsJob.RecordCount - K)) + i) = "DISC" & i
174                       gsSql2(69 + (6 * (rsJob.RecordCount - K)) + i) = IIf(Trim(rsJob("FREEC")), 100, 0)
175                       gsSql1(69 + (7 * (rsJob.RecordCount - K)) + i) = "VALID" & i
176                       gsSql2(69 + (7 * (rsJob.RecordCount - K)) + i) = Trim(rsJob("VALID"))
177                       gsSql1(69 + (8 * (rsJob.RecordCount - K)) + i) = "FREEC" & i
178                       gsSql2(69 + (8 * (rsJob.RecordCount - K)) + i) = IIf(Trim(rsJob("FREEC")), 1, 0)
179                       i = i + 1
180                   End If
181                   rsJob.MoveNext
182               Loop
183           End If
184       End If
          
          'Set rsTemp = objO.FetchData("SELECT * FROM t_INVOICE WHERE noinv = '" & Text(20) & "' AND no_wo = '" & NoWo & "' AND no_chassis = '" & Text(17) & "'")
          'If Not rsTemp.EOF Then
          '    sSQL = GensSql2("T_INVOICE", True) + "WHERE no_chassis = '" & Text(17) & "' AND no_WO = '" & NoWo & "' AND noInv = '" & Text(20) & "' AND FS_LOG_DEL = ''"
          'Else
185       sSQL = GensSql2("T_INVOICE", False)
          'End If
186       Call objHT.UpdateData_Bengkel(sSQL)
187       sSQLi2 = Replace(sSQL, "T_INVOICE", "T_INVOICENew")
          'If rsTemp.State <> adStateClosed Then rsTemp.Close
          
          'Update pajak
      '    Call objHT.UpdateData_Bengkel("update T_INVOICE set FN_PAJAK = " & Round(CDbl(subJob) * 10 / 110) & " " _
      '         & "where NO_WO ='" & noWO & "' and (DELIV='PRC' or DELIV='OKE') AND NOINV  = '" & Text(20) & IIf(nDuoInv, "-B", "") & "' and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'")
      '    nSQLo(1) = "update T_INVOICEnew set FN_PAJAK = " & Round(CDbl(subJob) * 10 / 110) & " " _
      '         & "where NO_WO ='" & noWO & "' and (DELIV='PRC' or DELIV='OKE') AND NOINV  = '" & Text(20) & IIf(nDuoInv, "-B", "") & "' and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'"
            '' 2016-03-29 || revisi fn_pajak dari afiliasi non ppn
188           Call objHT.UpdateData_Bengkel("update T_INVOICE set FN_PAJAK = " & Round(CDbl(fn_pajak)) & " " _
               & "where NO_WO ='" & noWO & "' and (DELIV='PRC' or DELIV='OKE') AND NOINV  = '" & text(20) & IIf(nDuoInv, "-B", "") & "' and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'")
189            nSQLo(1) = "update T_INVOICEnew set FN_PAJAK = " & Round(CDbl(fn_pajak)) & " " _
               & "where NO_WO ='" & noWO & "' and (DELIV='PRC' or DELIV='OKE') AND NOINV  = '" & text(20) & IIf(nDuoInv, "-B", "") & "' and FS_NO_RETUR = '' AND STATUS <> 'AFFNONPPN'"
        
         'On Error GoTo 0
190      Exit Sub

simpanInvoice2_Error:
191       erTranAll = True
192       ErrMsg.judul = ""
193       ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanInvoice2 of Form frmBilling")
194       ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanInvoice2 of Form frmBilling")
End Sub

Private Sub simpanHistory(nWO As String)
      Dim rsHlp As ADODB.Recordset

          ' SIMPAN MASTER HISTORY
1        On Error GoTo simpanHistory_Error

2         ClearsSql
          
3         gsSql1(0) = "NOINV"
4         gsSql2(0) = text(20)
5         gsSql1(1) = "CUSTOMER"
6         gsSql2(1) = text(4)
7         gsSql1(2) = "CUST_ADRES"
8         gsSql2(2) = txt5
9         gsSql1(3) = "KOTA"
10        gsSql2(3) = text(6)
11        gsSql1(4) = "NO_POLISI"
12        gsSql2(4) = text(16)
13        gsSql1(5) = "NO_ENGINE"
14        gsSql2(5) = text(18)
15        gsSql1(6) = "NO_CHASSIS"
16        gsSql2(6) = text(17)
17        gsSql1(7) = "FS_KD_DEALER"
18        gsSql2(7) = gKdCbg
19        gsSql1(8) = "FS_NM_DEALER"
20        gsSql2(8) = gNmCbg
21        gsSql1(9) = "FRONT"
22        gsSql2(9) = Trim(text(19))
23        gsSql1(10) = "MECH"
24        gsSql2(10) = text(38)
25        gsSql1(11) = "KM"
26        gsSql2(11) = text(37)
          
27        gsSql1(12) = "DISC"
28        gsSql2(12) = setKoma(CDbl(text(21)))
29        gsSql1(13) = "DISCT"
30        gsSql2(13) = CDbl(text(23))
          
31        gsSql1(14) = "PART_G"
32        gsSql2(14) = Val(subPartGenuine * (1 - (CDbl(text(23)) / 100))) 'CDbl(Text(24))     'subPart    '
33        gsSql1(15) = "PART_NG"
34        gsSql2(15) = Val(subPartNGenuine * (1 - (CDbl(text(23)) / 100)))
          
35        gsSql1(16) = "MTRL"
36        gsSql2(16) = Val(CDbl(text(28)))
37        gsSql1(17) = "OIL_G"
38        gsSql2(17) = Val(subOliGenuine * (1 - (CDbl(text(23)) / 100))) 'CDbl(Text(25))     'subOli     '
39        gsSql1(18) = "OIL_NG"
40        gsSql2(18) = Val(subOliNGenuine * (1 - (CDbl(text(23)) / 100)))
          
41        gsSql1(19) = "MET"
42        gsSql2(19) = Val(CDbl(Text42.text))
43        gsSql1(20) = "SBLT"
44        gsSql2(20) = Val(CDbl(text(27)))     'subSublet  '
45        gsSql1(21) = "LABOR"
46        gsSql2(21) = IIf(lblPDS.Visible, Val(subJob), IIf(nDuoInv, Val(subJob), Val(CDbl(text(22)))))
47        gsSql1(22) = "BODY"
48        gsSql2(22) = Val(CDbl(text(39)))     'subBody    '
49        gsSql1(23) = "TOTAL"
50        gsSql2(23) = Round(CDbl((subPartGenuine * (1 - (CDbl(text(23)) / 100))) + (subPartNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(text(28)) + (subOliGenuine * (1 - (CDbl(text(23)) / 100))) + (subOliNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(Text42.text) + CDbl(text(27)) + IIf(lblPDS.Visible, subJob, CDbl(text(22))) + CDbl(text(39))), 0)
          'CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)) + IIf(lblPDS.Visible, subJob, CDbl(Text(22))) + CDbl(Text(39)) + CDbl(Text(28)) + CDbl(text42.text)
          'subPart + subOli + subSublet + subBody + CDbl(Text(28)) + CDbl(text42.text) 'CDbl(Text(29))
          
      '    Set rsHlp = objO.FetchData("SELECT * FROM jobdta where no_wo = '" & Trim(text(0)) & "' AND no_chassis = '" & text(17) & "'")
      '    If Not rsHlp.EOF Then
51            gsSql1(24) = "[DATE]"
52            gsSql2(24) = TglYYYYMMDD(rsJobDta("DATE"))

              Dim hlp As String
              Dim hlp2 As String
              Dim hlp3 As String
53            For i = 1 To 7
54                hlp = "REP" & i
55                hlp2 = "MEKANIK" & i
56                hlp3 = "KLPKERJA" & i
57                gsSql1(24 + i) = hlp
58                gsSql2(24 + i) = rsJobDta(hlp)
59                gsSql1(31 + i) = hlp2
60                gsSql2(31 + i) = rsJobDta(hlp2)
61                gsSql1(38 + i) = hlp3
62                gsSql2(38 + i) = rsJobDta(hlp3)
63            Next
              
64            gsSql1(46) = "PAID"
65            gsSql2(46) = rsJobDta("PAID")
66            gsSql1(47) = "STATUS"
67            gsSql2(47) = rsJobDta("STATUS")
68            gsSql1(48) = "NOFAK"
69            gsSql2(48) = rsJobDta("NOFAK")
70            gsSql1(49) = "REMARK1"
71            gsSql2(49) = rsJobDta("REMARK1")
72            gsSql1(61) = "SUGGEST"
73            gsSql2(61) = rsJobDta("SUGGEST")
              
74            gsSql1(50) = "FN_RETURN"
75            gsSql2(50) = rsJobDta("FN_RETURN")
76            gsSql1(51) = "FS_RC"
77            gsSql2(51) = Nul(rsJobDta("FS_RC"))
78            gsSql1(52) = "SBE"
79            gsSql2(52) = Nul(rsJobDta("SBE"))
80            gsSql1(53) = "TIPEWO"
81            gsSql2(53) = rsJobDta("TPSVC")
82            gsSql1(54) = "KODETIPE"
83            gsSql2(54) = rsJobDta("KODETIPE")
      '    End If
          
84        gsSql1(55) = "NO_WO"
85        gsSql2(55) = nWO
86        gsSql1(56) = "MOU"
87        gsSql2(56) = Trim(text(34))
88        gsSql1(57) = "FS_KD_PLG"
89        gsSql2(57) = Trim(text(2))
90        gsSql1(58) = "FS_KD_WIL"
91        gsSql2(58) = Trim(text(3))
92        gsSql1(59) = "PDS"
93        gsSql2(59) = IIf(lblPDS.Visible, 1, 0)
          
94        gsSql1(60) = "THS"
95        gsSql2(60) = IIf(nTHS, 1, 0)
96        gsSql1(62) = "INVDATE"
97        gsSql2(62) = Format(text(1), "MM/DD/YYYY")
          
          
      '    Set rsTemp = objO.FetchData("SELECT * FROM HISTORY WITH (NOLOCK) WHERE noinv = '" & Text(20) & "' AND no_wo = '" & nWO & "' AND no_chassis = '" & Text(17) & "' AND FS_LOG_DEL = ''")
          'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
98        Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM HISTORY WITH (NOLOCK) WHERE noinv = '" & text(20) & "' AND no_wo = '" & nWO & "' AND no_chassis = '" & text(17) & "' AND FS_LOG_DEL = ''")
          
99        If Not rsTemp.EOF Then
100           PSQL = GensSql2("HISTORY", True) + "WHERE no_chassis = '" & text(17) & "' AND no_WO = '" & nWO & "' AND noInv = '" & text(20) & "' AND FS_LOG_DEL = ''"
101       Else
102           PSQL = GensSql2("HISTORY", False)
103       End If
104       Call objHT.UpdateData_Bengkel(PSQL)
          
          ' SIMPAN HISTORY JOB ADD
          '03/10/2012
          'Call objHT.InsertData_Bengkel("INSERT INTO HIST_JOBADD SELECT DISTINCT *, '','','" & text(20) & IIf(nDuoInv, "-A", "") & "' FROM JOBADD2 WHERE no_wo = '" & nWO & "' AND no_chassis = '" & text(17) & "'")
105       Call objHT.InsertData_Bengkel("INSERT INTO HIST_JOBADD SELECT DISTINCT *, '','','" & text(20) & "'+CASE WHEN FREEC = 0 THEN '" & IIf(nDuoInv, "-B", "") & "' ELSE '" & IIf(nDuoInv, "-A", "") & "'  END FROM JOBADD2 WITH (NOLOCK) WHERE no_wo = '" & nWO & "' AND no_chassis = '" & text(17) & "'")
106       Call objHT.InsertData_Bengkel("INSERT INTO HIST_JOBADDHDR SELECT DISTINCT * FROM JOBADDHDR WITH (NOLOCK) WHERE no_wo = '" & nWO & "' AND no_chassis = '" & text(17) & "'")
          
          
          
          ' SIMPAN HISTORY SUBLET
107       Call objHT.UpdateData_Bengkel("UPDATE jobsublet SET no_inv = '" & text(20) & IIf(nDuoInv, "-A", "") & "' WHERE NO_WO = '" & nWO & "' AND no_chassis = '" & text(17) & "'")
108       Call objHT.DeleteData_Bengkel("DELETE HIST_SUBLET WHERE NO_WO = '" & nWO & "' AND noINV = '" & text(20) & IIf(nDuoInv, "-A", "") & "' AND no_chassis = '" & text(17) & "'")
109       Call objHT.InsertData_Bengkel("INSERT INTO HIST_SUBLET SELECT NOPO, NO_WO, NO_INV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, PRICE, NONOTA, KETERANGAN, FLAG, SUPPLIER, TAGIH " _
                                      & "FROM JOBSUBLET WHERE EXISTS (SELECT * FROM JOBSUBLET WHERE NO_WO = '" & nWO & "' AND no_chassis = '" & text(17) & "') AND  NO_WO = '" & nWO & "' AND no_chassis = '" & text(17) & "'")
          
          ' SIMPAN HISTORY PART
110       With rsPart
111           If .State <> adStateClosed Then
112               If Not .BOF Then
113                   .MoveFirst
114                   Do While Not .EOF
                          
115                       ClearsSql
116                       gsSql1(0) = "NO_WO"
117                       gsSql2(0) = nWO
118                       gsSql1(1) = "KODEITEM"
119                       gsSql2(1) = rsPart("KODEITEM")
120                       gsSql1(2) = "KETERANGAN"
121                       gsSql2(2) = rsPart("NAMAITEM")
122                       gsSql1(3) = "QTY"
123                       gsSql2(3) = rsPart("QTY")
124                       gsSql1(4) = "HARGA"
125                       gsSql2(4) = setKoma(IIf(rsPart("HARGA") = "", 0, CDbl(rsPart("HARGA"))))
126                       gsSql1(5) = "OLI"
127                       gsSql2(5) = IIf(rsPart("OLI"), 1, 0)
128                       gsSql1(6) = "NOINV"
129                       gsSql2(6) = text(20) & IIf(nDuoInv, "-A", "")
130                       gsSql1(7) = "FN_EXT_DISKON"
131                       gsSql2(7) = Replace(CDbl(rsPart("EXTDISC")), ",", ".")
                          
      '                    Set rsTemp = objO.FetchData("SELECT * FROM T_PART WHERE NO_WO = '" & nWO & "' AND KODEITEM = '" & rsPart("KODEITEM") & "' and FS_LOG_DEL = ''")
      '                    If Not rsTemp.EOF Then
      '                        sSQL = GensSql2("T_PART", True) + "WHERE NO_WO = '" & nWO & "' AND KODEITEM = '" & rsPart("KODEITEM") & "'"
      '                    Else
132                           sSQL = GensSql2("T_PART", False)
      '                    End If
133                       Call objHT.UpdateData_Bengkel(sSQL)
134                       .MoveNext
135                   Loop
136               End If
137           End If
138       End With

         'On Error GoTo 0
139      Exit Sub

simpanHistory_Error:
140       erTranAll = True
141       ErrMsg.judul = ""
142       ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanHistory of Form frmBilling")
143       ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanHistory of Form frmBilling")
End Sub

Private Sub simpanTrnMOU(nWO As String)
          
1        On Error GoTo simpanTrnMOU_Error

2         ClearsSql
3         gsSql1(0) = "NO_WO"
4         gsSql2(0) = Trim(text(0))
5         gsSql1(1) = "NO_MOU"
6         gsSql2(1) = text(34)
7         gsSql1(2) = "NOINV"
8         gsSql2(2) = text(20)
9         gsSql1(3) = "NO_CHASSIS"
10        gsSql2(3) = text(17)
11        gsSql1(4) = "NO_POLISI"
12        gsSql2(4) = text(16)
13        gsSql1(5) = "FS_KD_DEALER"
14        gsSql2(5) = gKdCbg
15        gsSql1(6) = "TANGGAL"
16        gsSql2(6) = Format(isNowBilling, "MM/DD/YYYY")
17        gsSql1(7) = "STATUS"
18        gsSql2(7) = "U"
          
          ' UPDATE TERHUTANG
19        Call objHT.UpdateData_Bengkel("UPDATE MOUHDR SET HUTANG = HUTANG + " & Val(CDbl(text(29))) & " WHERE NO_MOU = '" & text(34) & "'")
          ' SIMPAN DI MOUTRN
20        Call objHT.InsertData_Bengkel("INSERT INTO MOUTRN VALUES ('" & Format(innow, "MM/DD/YYYY") & "', '" & text(34) & "', '" & Trim(text(0)) & "', '" & text(20) & "', '" & text(17) & "', '" & text(16) & "', '" & gKdCbg & "', " & Val(CDbl(text(29))) & ", 'U','','01/01/3000')")

         'On Error GoTo 0
21       Exit Sub

simpanTrnMOU_Error:
22        erTranAll = True
23        ErrMsg.judul = ""
24        ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanTrnMOU of Form frmBilling")
25        ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanTrnMOU of Form frmBilling")
End Sub

Private Function simpanACCNonAkunPenyeragaman(nWO As String) As Boolean
'On Error GoTo vbError
    simpanACCNonAkunPenyeragaman = False
    
    Dim rsSearch As ADODB.Recordset
    Dim SsqlNis As String
    Dim vNamaPlg As String
    'Dim vWilayah As String
    'Dim vSyaratBayar As String
    'Dim vStatusPersh As String
    
    Dim vSyaratByr As String
    'Dim vJmlHari As Double
    Dim vTglSkr As Date
    'Dim vTglTempo As Date
    Dim vTPenjualan As Double
    Dim vTPenjNonMet As Double
    Dim vTPenjSublet As Double
    Dim vTPembSublet As Double
    Dim vTPenjPart As Double            ' nilai penjualan part
    Dim vTPPN As Double
    Dim vTotal As Double
    Dim vTotalNonPPn As Double
    Dim vTMeterai As Double
    'Dim vKeterangan As String
    'Dim vKodeLgn As String
    
    Dim vPiutangGL As String
    Dim vPenjualanGL As String
    Dim vPenjSubletGL As String
    Dim vPenjPartGL As String           ' kode akun gl
    Dim vPpnGL As String
    Dim vMeteraiGL As String
    Dim vDelIndex As String
    Dim vHPPSublet As String
    Dim vPDPSublet As String
    
    Dim NilaiGLNonPart As Double
    
    If gConnecAAS = True Then
        vDelIndex = IIf(opWO(2), "TW", gKode) & Left(McDelIndex(7), 8)
        
        If gJurnalPart Then
            '' NILAI PART DITRANSFER KE ACC
            vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
            vTPenjSublet = Round(Koma2Titik(text(27)) * (100 / 110), 0)           ' sublet non ppn
            vTPenjNonMet = Round((Koma2Titik(text(22)) * (100 / 110)), 0)          ' jasa non ppn
            If nDuoInv Then
                vTMeterai = 0
            Else
                vTMeterai = Koma2Titik(Text42.text)                         ' materai
            End If
            vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)  ' jasa non ppn + meterai
            
            vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)
            vTPPN = vTotal - vTotalNonPPn

'            vTPPN = Round((vTotal * 10 / 110), 0)
'            vTPenjualan = (vTotal - vTPPN) + vTMeterai
        Else
            '' NILAI PART TIDAK DITRANSFER KE ACC
            vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
            vTPenjSublet = Round(Koma2Titik(text(27)) * (100 / 110), 0)    ' sublet non ppn
            vTPenjNonMet = Round(vTotal * (100 / 110), 0) - vTPenjSublet    ' jasa non ppn
            If nDuoInv Then
                vTMeterai = 0
            Else
                vTMeterai = Koma2Titik(Text42.text)                                ' materai
            End If
            vTPenjualan = vTPenjNonMet + vTMeterai + vTPenjSublet        ' jasa non ppn + meterai
            vTPPN = Round(vTotal * (10 / 110), 0)                          ' ppn
        End If
        
        'Proses Simpan Akses
        
            'sSQL = "select KodeLgn,NamaLgn,KodeWil, StatusPersh, SaldoJalan, SyaratBayar from Armst where KodeLgn = '" & Trim(text(2)) & "'"
                ' 26/09/2012...MARK !!  tidak ada coding lain yang menggunakan hal di bawah
'                vJmlHari = Val(Right(rsSearch("SyaratBayar"), Len(rsSearch("SyaratBayar")) - 1))
'                vnamalgn = Trim(rsSearch("NamaLgn").value)
'                vWilayah = Trim(rsSearch("KodeWil").value)
'                vSyaratBayar = Trim(rsSearch("SyaratBayar").value)
'                vStatusPersh = Trim(rsSearch("StatusPErsh").value)
            
            vTglSkr = gTanggal 'Format(isNowBilling, "MM/DD/YYYY")
            'vTglTempo = DateAdd(LCase(Left(text(41), 1)), Val(Right(text(41), Len(text(41)) - 1)), vTglSkr)
                        
            ' KODE AKUN
            Dim rsAkun As ADODB.Recordset
            Set rsAkun = objO.FetchData("select * from T_AKUN WITH (NOLOCK)")
            If Trim(text(40)) = "INT" Then
                If Not rsAkun.BOF Then
                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_INT"))
                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_INT"))
                    vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_SUBLET_INT"))
                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_INT"))
                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_INT"))
                    vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_INT"))
                    vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
                    If Trim(vPenjSubletGL) = "" Then
                        vTPenjNonMet = vTPenjNonMet + vTPenjSublet
                        vTPenjSublet = 0
                    End If
                End If
            ElseIf Trim(text(40)) = "EXT" Then
                If Not rsAkun.BOF Then
                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_EXT"))
                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_EXT"))
                    vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_SUBLET_EXT"))
                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_EXT"))
                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_EXT"))
                    vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_EXT"))
                    vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
                    If Trim(vPenjSubletGL) = "" Then
                        vTPenjNonMet = vTPenjNonMet + vTPenjSublet
                        vTPenjSublet = 0
                    End If
                End If
            ElseIf Trim(text(40)) = "AFFNONPPN" Then
                If Not rsAkun.BOF Then
                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_AFFNONPPN"))
                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_AFFNONPPN"))
                    vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_SUBLET_AFFNONPPN"))
                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_AFFNONPPN"))
                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_AFFNONPPN"))
                    vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_AFFNONPPN"))
                    vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
                    'vTPenjualan = vTPenjualan + vTPPN
                    'vTPenjNonMet = vTPenjNonMet + vTPPN
                    vTotal = vTotal - vTPPN
                    vTPPN = 0
                    If Trim(vPenjSubletGL) = "" Then
                        vTPenjNonMet = vTPenjNonMet + vTPenjSublet
                        vTPenjSublet = 0
                    End If
                End If
            ElseIf Trim(text(40)) = "AFFPPN" Then
                If Not rsAkun.BOF Then
                    vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_AFFPPN"))
                    vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_AFFPPN"))
                    vPenjSubletGL = cekKodeAkun(rsAkun("FS_KD_SUBLET_AFFPPN"))
                    vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_AFFPPN"))
                    vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_AFFPPN"))
                    vHPPSublet = cekKodeAkun(rsAkun("FS_KD_HPP_AFFPPN"))
                    vPDPSublet = cekKodeAkun(rsAkun("FS_KD_PDP"))
                    If Trim(vPenjSubletGL) = "" Then
                        vTPenjNonMet = vTPenjNonMet + vTPenjSublet
                        vTPenjSublet = 0
                    End If
                End If
            End If
            If rsAkun.State <> adStateClosed Then rsAkun.Close
            
            If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
                pesanIT = pesanIT + "Kode GL belum lengkap !"
                GoTo vbError
            End If
            If Trim(text(40)) <> "AFFNONPPN" And vPpnGL = "" Then
                pesanIT = pesanIT + "Kode GL belum lengkap !"
                GoTo vbError
            End If
            
            ' INSERT KE AR
            'If vTPenjualan <> 0 Then
                ' PENJUALAN
                sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                      & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) " _
                      & "values " _
                      & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', 'C'," _
                      & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                      & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                      & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', 'Penjualan Kredit No. WO : " & Trim(Trim(text(0))) & "'," _
                      & " " & SimpanKoma2Titik(vTPenjualan) & ", '" & gSA & "'," _
                      & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "')"
                Call objHT.InsertData_ACC(sSQL1)
                
                'PPN
                If vTPPN <> 0 Then
                    sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                          & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                          & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) values " _
                          & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', 'F'," _
                          & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', 'PPN Keluar No. WO : " & Trim(Trim(text(0))) & "'," _
                          & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "'," _
                          & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                End If
    
                'Update Armst Saldo berjalan
                sSQL1 = "Update Armst set SaldoJalan = SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(vKodeLgn) & "'"
                Call objHT.UpdateData_ACC(sSQL1)
            'End If
        
            'Membuat transaksi di GL
            'If vTotal + vTMeterai <> 0 Then
              sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                    & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, nowo, nofps) values " _
                    & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                    & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                    & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Call objHT.InsertData_ACC(sSQL1)
              NilaiGLNonPart = vTotal + vTMeterai
              sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                    & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                    & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                    & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                    & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
              Call objHT.InsertData_ACC(sSQL1)
              NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
                If vTPenjSublet <> 0 Then
                    sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                    NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
                End If
                If vTMeterai <> 0 Then
                    sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                    NilaiGLNonPart = NilaiGLNonPart - vTMeterai
                End If
                If vTPPN <> 0 Then
                    sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                    NilaiGLNonPart = NilaiGLNonPart - vTPPN
                End If
                
                
                If nDuoInv = True Then
                    Set rsAkun = objO.FetchData("select * from T_AKUN WITH (NOLOCK)")
                    If Trim(vTipeCustomer) = "INT" Then
                        If Not rsAkun.BOF Then
                            vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_INT"))
                            vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_INT"))
                            vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_INT"))
                            vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_INT"))
                        End If
                    ElseIf Trim(vTipeCustomer) = "EXT" Then
                        If Not rsAkun.BOF Then
                            vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_EXT"))
                            vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_EXT"))
                            vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_EXT"))
                            vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_EXT"))
                        End If
                    ElseIf Trim(vTipeCustomer) = "AFFNONPPN" Then
                        If Not rsAkun.BOF Then
                            vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_AFFNONPPN"))
                            vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_AFFNONPPN"))
                            vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_AFFNONPPN"))
                            vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_AFFNONPPN"))
                            vTotal = vTotal - vTPPN
                            vTPPN = 0
                        End If
                    ElseIf Trim(vTipeCustomer) = "AFFPPN" Then
                        If Not rsAkun.BOF Then
                            vPiutangGL = cekKodeAkun(rsAkun("FS_KD_PIUTANG_AFFPPN"))
                            vPenjualanGL = cekKodeAkun(rsAkun("FS_KD_LABOUR_AFFPPN"))
                            vPpnGL = cekKodeAkun(rsAkun("FS_KD_PPN_AFFPPN"))
                            vMeteraiGL = cekKodeAkun(rsAkun("FS_KD_METERAI_AFFPPN"))
                        End If
                    End If
                End If
                
                
                If gJurnalPart Then
                    '' NILAI SPAREPART
                    vTPenjPart = 0
                    Dim rsSearchPart As ADODB.Recordset
                    If rsPart.State <> adStateClosed Then
                        rsPart.MoveFirst
                        Do While Not rsPart.EOF
                            vTPenjPart = vTPenjPart + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
                            rsPart.MoveNext
                        Loop
                        
                        If NilaiGLNonPart <> vTPenjPart Then
                            If NilaiGLNonPart - vTPenjPart = 1 Or NilaiGLNonPart - vTPenjPart = -1 Then
                                vTPenjPart = NilaiGLNonPart
                            End If
                        End If
                        
                        ' AMBIL KODE AKUN PENJUALAN PART
                        Set rsSearchPart = objO.FetchDataACC("SELECT NoJSales, NoRSales FROM SIJURNAL WITH (NOLOCK) WHERE KODEJURNAL = '" & Trim(text(46)) & "'")
                        If Not rsSearchPart.EOF Then
                            vPenjPartGL = Trim(rsSearchPart("NoJSales"))
                        End If
                        
                        If Trim(vPenjPartGL) = "" Then
                            pesanIT = pesanIT + "Kode Akun Part belum lengkap !"
                            GoTo vbError
                        End If
        
'                        Set rsSearchPart = objO.FetchDataACC("SELECT * FROM GLTRN WHERE NOBUKTI = '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "' AND KODEGL = '" & Trim(vPenjPartGL) & "' AND DELINDEX = '" & Trim(vDelIndex) & "'")
'                        If Not rsSearchPart.EOF Then
'                            sSQL1 = "UPDATE GLTRN SET JLHKREDIT = JLHKREDIT + " & SimpanKoma2Titik(vTPenjPart) & " WHERE NOBUKTI = '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "' AND KODEGL = '" & Trim(vPenjPartGL) & "' AND DELINDEX = '" & Trim(vDelIndex) & "'"
'                            Call objHT.UpdateData_ACC(sSQL1)
'                        Else
                        
                        '---------------------------------------------------------
                        '-
                        '---------------------------------------------------------
                            sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                  & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                  & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                  & " " & SimpanKoma2Titik(vTPenjPart) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                  & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                            Call objHT.InsertData_ACC(sSQL1)
'                        End If
                        
                        sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTPenjPart) & ", " _
                              & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPenjPart) & " " _
                              & " where KodeGl = '" & vPenjPartGL & "'"
                        Call objHT.InsertData_ACC(sSQL1)
                    End If
                End If
                
                '' SUBLET
                ' AMBIL HPP SUBLET
                Dim vTPpnS As String
                Dim vTotalS As Double
                Dim vTPpn1 As String
                Dim vTBeli1 As Double
                Dim vTotal1 As Double
                Dim nonJurnal As Boolean
                Dim vKeterangan1 As String
                
                vTPembSublet = 0
                vTPenjSublet = 0
                'Set rsSearch = objO.FetchData("SELECT PKP, NAMABARANG, NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, PRICE*QTY AS HARGAJUAL, NONJURNAL FROM JOBSUBLET a INNER JOIN T_SUPPLIER B ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(Trim(text(0))) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
                If rsInvSublet.State <> adStateClosed Then
                    If Not rsInvSublet.EOF Then
                        Do While Not rsInvSublet.EOF
                            nonJurnal = IIf(Trim(rsInvSublet("NONJURNAL")) = 1, True, False)
                            If Not nonJurnal Then
                                vTotal1 = Round(CDbl(rsInvSublet("HPP")), 0)
                                vTPpn1 = Round(vTotal1 * 10 / 110, 0)
                                vTBeli1 = Round(vTotal1 - vTPpn1, 0)
                                vKeterangan1 = vKeterangan & " Sublet " '" (" & rsinvsublet("NAMABARANG") & ") " 'No. WO :" & Trim(TRIM(text(0)))
                                If Trim(rsInvSublet("PKP")) = "1" Then
                                    vTPpn1 = 0
                                    vTotal1 = Round(vTBeli1, 0)
                                End If
                                vTPembSublet = vTPembSublet + vTotal1
                            End If
                            rsInvSublet.MoveNext
                        Loop
                    End If
                End If

                'If Not nonJurnal Then
                    If vTPembSublet <> 0 Then
                        sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vHPPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                        
                        sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPDPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                    End If
                    
                    If vTPembSublet <> 0 Then
                        ' HPP SUBLET TERHADAP PDP
                        sSQL1 = "update glmst set JbulanIni = JbulanIni +" & (vTPembSublet) & ", " _
                              & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                              & " where KodeGl = '" & vHPPSublet & "'"
                        Call objHT.UpdateData_ACC(sSQL1)
                        
                        sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPembSublet) & ", " _
                              & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                              & " where KodeGl = '" & vPDPSublet & "'"
                        Call objHT.UpdateData_ACC(sSQL1)
                    End If
                'End If
                
                sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
                      & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                      & " where KodeGl = '" & vPiutangGL & "'"
                Call objHT.UpdateData_ACC(sSQL1)
                
                sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                      & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                      & " where KodeGl = '" & vPenjualanGL & "'"
                Call objHT.UpdateData_ACC(sSQL1)
                
                sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjSublet) & ", " _
                      & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
                      & " where KodeGl = '" & vPenjSubletGL & "'"
                Call objHT.UpdateData_ACC(sSQL1)
                
                sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTMeterai) & ", " _
                      & " JKreditBi = JKreditBi +" & vTMeterai & " " _
                      & " where KodeGl = '" & vMeteraiGL & "'"
                Call objHT.UpdateData_ACC(sSQL1)
                
                sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPPN) & ", " _
                      & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPPN) & " " _
                      & " where KodeGl = '" & vPpnGL & "'"
                Call objHT.UpdateData_ACC(sSQL1)
            'End If
            
            If nDuoInv Then
                ' UNTUK INVOICE KEDUA -B KE DEALER PENJUAL
                ' MEMBUAT TRANSAKSI DUO INV DI AR
                vTotal = subJob
                vTPenjNonMet = Round(vTotal * (100 / 110), 0)  'dikurangi ppn
                vTMeterai = Koma2Titik(Text42.text)               'materai
                vTPenjualan = vTPenjNonMet + vTMeterai         'non ppn + meterai
                vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          'ppn
                'vTPenjualan = (vTotal - vTPPN) + vTMeterai
                vKeterangan = gInsCust2 & ". AN/ " & vKeterangan
                
                
                If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
                    pesanIT = pesanIT + "Kode GL belum lengkap !"
                    GoTo vbError
                End If
                If Trim(vTipeCustomer) <> "AFFNONPPN" And vPpnGL = "" Then
                    pesanIT = pesanIT + "Kode GL belum lengkap !"
                    GoTo vbError
                End If
        
                    ' PENJUALAN
                    sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                          & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) " _
                          & "values " _
                          & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', 'C'," _
                          & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                          & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', 'Penjualan Kredit No. WO : " & Trim(Trim(text(0))) & "'," _
                          & " " & SimpanKoma2Titik(vTPenjualan) & ", '" & gSA & "'," _
                          & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                    
                    'PPN
                    If vTPPN <> 0 Then
                        sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                              & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                              & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry) values " _
                              & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', 'F'," _
                              & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                              & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                              & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', 'PPN Keluar No. WO : " & Trim(Trim(text(0))) & "'," _
                              & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "'," _
                              & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                    End If
        
                    'Update Armst Saldo berjalan
                    sSQL1 = "Update Armst set SaldoJalan = SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(gKdPlgDealer2) & "'"
                    Call objHT.UpdateData_ACC(sSQL1)
                'End If
                
                ' MEMBUAT TRANSAKSI DUO INV DI GL
                'If vTotal + vTMeterai <> 0 Then
                    sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                  
                    sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                    Call objHT.InsertData_ACC(sSQL1)
                
                    If vTMeterai <> 0 Then
                        sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                    End If
                    
                    If vTPPN <> 0 Then
                        sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                              & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                              & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                              & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                              & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                        Call objHT.InsertData_ACC(sSQL1)
                    End If
                    
                    sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
                          & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                          & " where KodeGl = '" & vPiutangGL & "'"
                    Call objHT.UpdateData_ACC(sSQL1)
                    
                    sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                          & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                          & " where KodeGl = '" & vPenjualanGL & "'"
                    Call objHT.UpdateData_ACC(sSQL1)
                    
                    sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTMeterai) & ", " _
                          & " JKreditBi = JKreditBi +" & vTMeterai & " " _
                          & " where KodeGl = '" & vMeteraiGL & "'"
                    Call objHT.UpdateData_ACC(sSQL1)
                    
                    sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPPN) & ", " _
                          & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPPN) & " " _
                          & " where KodeGl = '" & vPpnGL & "'"
                    Call objHT.UpdateData_ACC(sSQL1)
                'End If
            End If
            
            If erTranAll = False Then
                Call objHT.CloseTransactionAll
                Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 1)
                simpanACCNonAkunPenyeragaman = True
            Else
                Call objHT.CloseTransactionAll_Err
                Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
                MsgBox "Ada Error, Pembuatan Invoice dibatalkan, buat invoice kembali / Hubungi EDP!"
                simpanACCNonAkunPenyeragaman = False
            End If
    Else
        Call objHT.CloseTransactionAll_Err
        Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
        MsgBox "Ada Error, Pembuatan Invoice dibatalkan, buat invoice kembali / Hubungi EDP!"
        simpanACCNonAkunPenyeragaman = False
        MsgBox "Ada Error, Pembuatan Invoice dibatalkan, buat invoice kembali / Hubungi EDP!"
    End If
    Exit Function
    
vbError:
    simpanACCNonAkunPenyeragaman = False
    Call objHT.CloseTransactionAll_Err
    Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
    Exit Function
End Function

Private Function simpanACC2(nWO As String) As Boolean
          'On Error GoTo vbError
1         simpanACC2 = False
          Dim isSudahDikuragi As Boolean
2         isSudahDikurangi = False
          
          Dim rsSearch As ADODB.Recordset
          Dim SsqlNis As String
          Dim vNamaPlg As String
          'Dim vWilayah As String
          'Dim vSyaratBayar As String
          'Dim vStatusPersh As String
          
          Dim vSyaratByr As String
          'Dim vJmlHari As Double
          Dim vTglSkr As Date
          'Dim vTglTempo As Date
          Dim vTPenjualan As Double
          Dim vTPenjNonMet As Double
          Dim vTPenjSublet As Double
          Dim vTPembSublet As Double
          Dim vTPenjPart As Double            ' nilai penjualan part
          
          Dim vJSPPN As Double                ' NILAI PPN JASA + SUBLET
          Dim vPaPPN As Double                ' NILAI PPN Part
          Dim vOlPPN As Double                ' NILAI PPN Oli
          Dim vAcPPN As Double                ' NILAI PPN Accs
          
          Dim vTPPN As Double
          Dim vTotal As Double
          Dim vTotalNonPPn As Double
          Dim vTMeterai As Double
          'Dim vKeterangan As String
          'Dim vKodeLgn As String
              
          Dim vPiutangGL As String
          Dim vPenjualanGL As String
          Dim vPenjSubletGL As String
          Dim vPenjPartGL As String           ' kode akun gl
          Dim vPenjOlieGL As String           ' kode akun gl
          Dim vPenjAccsGL As String           ' kode akun gl
          Dim vPpnGL As String
          Dim vPpnPartGL As String
          Dim vPpnAccsGL As String
          Dim vMeteraiGL As String
          Dim vDelIndex As String
          Dim vHPPSublet As String
          Dim vPDPSublet As String
          
          Dim NilaiGLNonPart As Double
          
          Dim totalKreditTemp As Double
          Dim totalDebitTemp As Double
          
          Dim adaSelisih As Boolean
          Dim kreditSelisih As Boolean
          
3         kreditSelisih = False
4         adaSelisih = False
              
5         If gConnecAAS = True Then
6             vDelIndex = IIf(opWO(2), "TW", gKode) & Left(McDelIndex(7), 8)
              
7             If gJurnalPart Then
                  '' NILAI PART DITRANSFER KE ACC
8                 vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
9                 vTPenjSublet = Round(Koma2Titik(text(27)) * (100 / 110), 0)           ' sublet non ppn
10                vTPenjNonMet = Round((Koma2Titik(text(22)) * (100 / 110)), 0)          ' jasa non ppn
11                If nDuoInv Then
12                    vTMeterai = 0
13                Else
14                    vTMeterai = Koma2Titik(Text42.text)                         ' materai
15                End If
16                vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)  ' jasa non ppn + meterai
                  
17                vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)
18                vTPPN = vTotal - vTotalNonPPn
                  ' REVISI THANIA 09/05/2014
19                vJSPPN = vTotal - Round(((Koma2Titik(text(24)) + Koma2Titik(text(25)))), 0) - vTPenjSublet - vTPenjNonMet

      '            vTPPN = Round((vTotal * 10 / 110), 0)
      '            vTPenjualan = (vTotal - vTPPN) + vTMeterai
20            Else
                  '' NILAI PART TIDAK DITRANSFER KE ACC
21                vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
22                vTPenjSublet = Round(Koma2Titik(text(27)) * (100 / 110), 0)    ' sublet non ppn
23                vTPenjNonMet = Round(vTotal * (100 / 110), 0) - vTPenjSublet    ' jasa non ppn
24                If nDuoInv Then
25                    vTMeterai = 0
26                Else
27                    vTMeterai = Koma2Titik(Text42.text)                                ' materai
28                End If
29                vTPenjualan = vTPenjNonMet + vTMeterai + vTPenjSublet        ' jasa non ppn + meterai
30                vTPPN = Round(vTotal * (10 / 110), 0)                          ' ppn
31            End If
                  
32                vTglSkr = gTanggal 'Format(isNowBilling, "MM/DD/YYYY")
                  'vTglTempo = DateAdd(LCase(Left(text(41), 1)), Val(Right(text(41), Len(text(41)) - 1)), vTglSkr)
                              
              'AKUNBARU
                  Dim isDyna As Boolean
33                isDyna = IIf(InStr(1, UCase(text(14).text), "DYNA") > 0, True, False)
                  Dim isCust As String
                  Dim isExAflAR As Boolean
                  Dim isExAflSI As Boolean
                  Dim isExIntSI As Boolean
                  Dim nKodeSi As String
              
34                nKodeSi = vKodeLgn 'Trim(text(2)) 'vKodeLgn
35                isCust = Trim(text(40))
36                isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
37                If isExAflAR = True Then  'AFL
38                    isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
39                    If isExAflSI = True Then
40                        isCust = "AFFPPN"
41                    Else
42                        ErrMsg.judul = ""
43                        ErrMsg.message = "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting"
44                        ErrMsg.Show
                          ''MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
45                        GoTo vbError
46                    End If
47                Else
48                    If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
49                        ErrMsg.judul = ""
50                        ErrMsg.message = "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting"
51                        ErrMsg.Show
                          
                          '''MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
52                        GoTo vbError
53                    End If
54                    If isCust = "INT" Then
55                        isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
56                        If isExIntSI = False Then
57                            ErrMsg.judul = ""
58                            ErrMsg.message = "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting"
59                            ErrMsg.Show
                              ''MsgBox "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
60                            GoTo vbError
61                        End If
62                    End If
63                End If
              
              'AKUNBARU
              Dim rsGL As ADODB.Recordset
64            Select Case isCust
                  Case "EXT"
65                    Set rsGL = objO.Get_AkunGL_NEW("OTHERS", isDyna, False)
66                Case Else  'INT/AFL
67                    Set rsGL = objO.Get_AkunGL_NEW(nKodeSi, isDyna, False)
68            End Select
                  
      '''piutang gr
      '''  penj bengkel
      '''  penj sublet
      '''  penj Part
      '''  penj oli
      '''  penj Accs
      '''  PPn Jasa
      '''  PPn Part + oli
      '''  PPn Accs
      '''  Materai
      '''HPP sublet
      ''' Persediaan sublet
       
                  
69                If Not rsGL.BOF Then
70                    vPiutangGL = cekKodeAkun(rsGL.Fields(0).value)
71                    vPenjualanGL = cekKodeAkun(rsGL.Fields(1).value)
72                    vPenjSubletGL = cekKodeAkun(rsGL.Fields(2).value)
73                    vPenjPartGL = cekKodeAkun(rsGL.Fields(3).value)
74                    vPenjOlieGL = cekKodeAkun(rsGL.Fields(4).value)
75                    vPenjAccsGL = cekKodeAkun(rsGL.Fields(5).value)
76                    vPpnGL = cekKodeAkun(rsGL.Fields(6).value)
77                    vPpnPartGL = cekKodeAkun(rsGL.Fields(7).value)
78                    vPpnAccsGL = cekKodeAkun(rsGL.Fields(8).value)
79                    vMeteraiGL = cekKodeAkun(rsGL.Fields(9).value)
80                    vHPPSublet = cekKodeAkun(rsGL.Fields(10).value)
81                    vPDPSublet = cekKodeAkun(rsGL.Fields(11).value)
      ''                If Trim(vPenjSubletGL) = "" Then
      ''                    vTPenjNonMet = vTPenjNonMet + vTPenjSublet
      ''                    vTPenjSublet = 0
      ''                End If
82                End If

83                If rsGL.State <> adStateClosed Then rsGL.Close
                  
84                If vPiutangGL = "" Or vPenjualanGL = "" Or vPenjSubletGL = "" Or vPenjPartGL = "" Or vPenjOlieGL = "" Or vPenjAccsGL = "" _
                     Or vPpnGL = "" Or vPpnPartGL = "" Or vPpnAccsGL = "" Or vMeteraiGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Then
85                    pesanIT = pesanIT + "Kode GL belum lengkap !"
86                    GoTo vbError
87                End If
      ''''            If Trim(text(40)) <> "AFFNONPPN" And vPpnGL = "" Then
      ''''                pesanIT = pesanIT + "Kode GL belum lengkap !"
      ''''                GoTo vbError
      ''''            End If

                  ' ------------- PROSES PENGECEKAN -------------
88                totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTotal + vTMeterai))
89                NilaiGLNonPart = vTotal + vTMeterai
                  
90                totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPenjNonMet))
91                NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
                  
92                totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPenjSublet))
93                NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
                  
94                If vTMeterai <> 0 Then
95                    totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTMeterai))
96                    NilaiGLNonPart = NilaiGLNonPart - vTMeterai
97                End If
                  
98                If (vTPenjNonMet + vTPenjSublet) <> 0 Then
99                    totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)))
100                   NilaiGLNonPart = NilaiGLNonPart - (Round(((vTPenjNonMet + vTPenjSublet) / 10), 0))
101               End If
                      
102               If nDuoInv = True Then
103                   nKodeSi = Trim(gKdPlgDealer2)
104                   isCust = Trim(text(40))
105                   isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
106                   If isExAflAR = True Then  'AFL
107                       isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
108                       If isExAflSI = True Then
109                           isCust = "AFFPPN"
110                       End If
111                   Else
112                       If isCust = "INT" Then
113                           isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
114                       End If
115                   End If
          
116               End If
              
117               If gJurnalPart Then
                          Dim SumPartTemp As Double
                          Dim SumOlieTemp As Double
                          Dim SumAccsTemp As Double
                          Dim SumPartFullTemp As Double
                          Dim SumOlieFullTemp As Double
                          Dim SumAccsFullTemp As Double
                          '' NILAI SPAREPART
118                       vTPenjPart = 0
                          Dim rsSearchPartTemp As ADODB.Recordset
119                       If rsPart.State <> adStateClosed Then
120                           SumPartTemp = 0
121                           SumOlieTemp = 0
122                           SumAccsTemp = 0
123                           SumPartFullTemp = 0
124                           SumOlieFullTemp = 0
125                           SumAccsFullTemp = 0
126                           vPaPPN = 0
127                           vOlPPN = 0
128                           vAcPPN = 0
129                           If rsPart.RecordCount > 0 Then
130                               rsPart.MoveFirst
131                               Do Until rsPart.EOF
132                                   Select Case Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH (NOLOCK) Inner Join Profilacc..SCC B WITH (NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsPart("KodeItem")) & "'")), 3)
                                      Case "PAR"
133                                       SumPartTemp = SumPartTemp + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
134                                       SumPartFullTemp = SumPartFullTemp + CDbl(rsPart("HARGA"))
135                                   Case "OLI"
136                                       SumOlieTemp = SumOlieTemp + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
137                                       SumOlieFullTemp = SumOlieFullTemp + CDbl(rsPart("HARGA"))
138                                   Case "ACC"
139                                       SumAccsTemp = SumAccsTemp + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
140                                       SumAccsFullTemp = SumAccsFullTemp + CDbl(rsPart("HARGA"))
141                                   End Select
142                                   rsPart.MoveNext
143                               Loop
144                           End If
145                           vTPenjPart = SumPartTemp + SumOlieTemp + SumAccsTemp
                                                  
146                           If NilaiGLNonPart <> vTPenjPart Then
147                               If NilaiGLNonPart - vTPenjPart = 1 Then
148                                   If SumPartTemp <> 0 Then
149                                       SumPartTemp = SumPartTemp + 1
150                                   Else
151                                       If SumOlieTemp <> 0 Then
152                                           SumOlieTemp = SumOlieTemp + 1
153                                       Else
154                                           If SumAccsTemp <> 0 Then
155                                               SumAccsTemp = SumAccsTemp + 1
156                                           End If
157                                       End If
158                                   End If
159                                   vTPenjPart = NilaiGLNonPart
160                               ElseIf NilaiGLNonPart - vTPenjPart = -1 Then
161                                   If SumPartTemp <> 0 Then
162                                       SumPartTemp = SumPartTemp - 1
163                                   Else
164                                       If SumOlieTemp <> 0 Then
165                                           SumOlieTemp = SumOlieTemp - 1
166                                       Else
167                                           If SumAccsTemp <> 0 Then
168                                               SumAccsTemp = SumAccsTemp - 1
169                                           End If
170                                       End If
171                                   End If
172                                   vTPenjPart = NilaiGLNonPart
173                               End If
174                           End If
                              
175                           vPaPPN = SumPartFullTemp - SumPartTemp
176                           vOlPPN = SumOlieFullTemp - SumOlieTemp
177                           vAcPPN = SumAccsFullTemp - SumAccsTemp
                              
178                           If SumPartTemp <> 0 Then
179                               totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(SumPartTemp))
180                           End If
                              
181                           If SumOlieTemp <> 0 Then
182                               totalKreditTemp = totalKreditTemp + SimpanKoma2Titik(SumOlieTemp)
183                           End If
                              
184                           If (SumPartTemp + SumOlieTemp) <> 0 Then
185                               totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)))
186                           End If
                              
                              
187                           If SumAccsTemp <> 0 Then
188                               totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(SumAccsTemp))
                                        
189                               totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(Round((vAcPPN), 0)))
190                           End If
191                       End If
192                   End If
                      
                      '' SUBLET
                      ' AMBIL HPP SUBLET
                      Dim vTPpnSTemp As String
                      Dim vTotalSTemp As Double
                      Dim vTPpn1Temp As String
                      Dim vTBeli1Temp As Double
                      Dim vTotal1Temp As Double
                      Dim nonJurnalTemp As Boolean
                      Dim vKeterangan1Temp As String
                  
193                   vTPembSublet = 0
                      'vTPenjSublet = 0
                      
194                   If rsInvSublet.State <> adStateClosed Then
195                       If Not rsInvSublet.EOF Then
196                           Do While Not rsInvSublet.EOF
197                               nonJurnalTemp = IIf(Trim(rsInvSublet("NONJURNAL")) = 1, True, False)
198                               If Not nonJurnalTemp Then
199                                   vTotal1Temp = Round(CDbl(rsInvSublet("HPP")), 0)
200                                   vTPpn1Temp = Round(vTotal1Temp * 10 / 110, 0)
201                                   vTBeli1Temp = Round(vTotal1Temp - vTPpn1Temp, 0)
202                                   vKeterangan1Temp = vKeterangan & " Sublet " '" (" & rsinvsublet("NAMABARANG") & ") " 'No. WO :" & Trim(TRIM(text(0)))
203                                   If Trim(rsInvSublet("PKP")) = "1" Then
204                                       vTPpn1Temp = 0
205                                       vTotal1Temp = Round(vTBeli1Temp, 0)
206                                   End If
207                                   vTPembSublet = vTPembSublet + vTotal1Temp
208                               End If
209                               rsInvSublet.MoveNext
210                           Loop
211                       End If
212                   End If
                      
213                   If vTPembSublet <> 0 Then
214                       totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTPembSublet))
215                       totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPembSublet))
216                   End If
                      
217                   If nDuoInv Then
                          ' UNTUK INVOICE KEDUA -B KE DEALER PENJUAL
                          ' MEMBUAT TRANSAKSI DUO INV DI AR
218                       vTotal = subJob
219                       vTPenjNonMet = Round(vTotal * (100 / 110), 0)  'dikurangi ppn
220                       vTMeterai = Koma2Titik(Text42.text)               'materai
                                   'non ppn + meterai
                          'vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          'ppn
                          ''2015-10-19 || AKTIF AFFPPN
221                       If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
222                           If Trim(gNpwp) = Trim(gInvNPWP2) Then
                                  '' 2016-03-29 || subjob sudah dikurangi ppn pada saat invoice
223                               vTotal = subJob ''vTPenjNonMet
224                               vTPenjNonMet = vTotal

225                               vTPPN = 0     'Round(vTotal * (10 / 110), 0)
226                           Else
227                               vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
228                           End If
229                       Else
230                           vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
231                       End If
232                       vTPenjualan = vTPenjNonMet + vTMeterai
                          'vTPenjualan = (vTotal - vTPPN) + vTMeterai
233                       vKeterangan = gInsCust2 & ". AN/ " & vKeterangan
                          
                          ' MEMBUAT TRANSAKSI DUO INV DI GL
                          
234                       totalDebitTemp = totalDebitTemp + (SimpanKoma2Titik(vTotal + vTMeterai))
                          
235                       totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPenjNonMet))
                          
236                       If vTMeterai <> 0 Then
237                           totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTMeterai))
238                       End If
                          
239                       If vTPPN <> 0 Then
240                           totalKreditTemp = totalKreditTemp + (SimpanKoma2Titik(vTPPN))
241                       End If
242                   End If
                  
243                   If totalDebitTemp <> totalKreditTemp Then
244                       If (totalDebitTemp - totalKreditTemp) = -1 Or (totalDebitTemp - totalKreditTemp) = 1 Then
245                           If (totalDebitTemp - totalKreditTemp) = -1 Then
246                               kreditSelisih = True
247                           End If
248                           adaSelisih = True
249                       End If
250                   End If
                      
                  
                  '------------------ AKHIR PENGECEKAN ---------------------------
                  
                  
251               If gJurnalPart Then
                      '' NILAI PART DITRANSFER KE ACC
252                   vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)
253                   vTPenjSublet = Round(Koma2Titik(text(27)) * (100 / 110), 0)           ' sublet non ppn
254                   vTPenjNonMet = Round((Koma2Titik(text(22)) * (100 / 110)), 0)          ' jasa non ppn
255                   If nDuoInv Then
256                       vTMeterai = 0
257                   Else
258                       vTMeterai = Koma2Titik(Text42.text)                         ' materai
259                   End If
260                   vTPenjualan = Round(vTPenjNonMet + vTMeterai + vTPenjSublet + ((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)  ' jasa non ppn + meterai
                      
261                   vTotalNonPPn = vTPenjSublet + vTPenjNonMet + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)
262                   vTPPN = vTotal - vTotalNonPPn
                      ' REVISI THANIA 09/05/2014
263                   vJSPPN = vTotal - Round(((Koma2Titik(text(24)) + Koma2Titik(text(25)))), 0) - vTPenjSublet - vTPenjNonMet
          
          '            vTPPN = Round((vTotal * 10 / 110), 0)
          '            vTPenjualan = (vTotal - vTPPN) + vTMeterai
264               Else
                      '' NILAI PART TIDAK DITRANSFER KE ACC
265                   vTotal = Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)), 0)
266                   vTPenjSublet = Round(Koma2Titik(text(27)) * (100 / 110), 0)    ' sublet non ppn
267                   vTPenjNonMet = Round(vTotal * (100 / 110), 0) - vTPenjSublet    ' jasa non ppn
268                   If nDuoInv Then
269                       vTMeterai = 0
270                   Else
271                       vTMeterai = Koma2Titik(Text42.text)                                ' materai
272                   End If
273                   vTPenjualan = vTPenjNonMet + vTMeterai + vTPenjSublet        ' jasa non ppn + meterai
274                   vTPPN = Round(vTotal * (10 / 110), 0)                          ' ppn
275               End If
              
                  ' INSERT KE AR
                  'If vTPenjualan <> 0 Then
                      ' PENJUALAN
                      
                      '------------------------------------------------------------------
                      '- 2015-04-09 | RUDI | PENGECEKAN TANGGAL BUKAN HARI INI          -
                      '------------------------------------------------------------------
                      
276                   If DateDiff("D", Format(vTglSkr, "MM/DD/YYYY"), Format(gTglServer, "MM/DD/YYYY")) > 0 Then
277                       MsgBox "Terjadi kesalahan tanggal saat menyimpan data. Ulangi proses penyimpanan data", vbInformation
278                       GoTo vbError
279                   End If
                      
280                   sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                            & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) " _
                            & "values " _
                            & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', 'C'," _
                            & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                            & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                            & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', 'Penjualan Kredit No. WO : " & Trim(Trim(text(0))) & "'," _
                            & " " & SimpanKoma2Titik(vTPenjualan) & ", '" & gSA & "'," _
                            & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "', '" & vPiutangGL & "')"
281                   Call objHT.InsertData_ACC(sSQL1)
                      
                      'PPN
282                   If vTPPN <> 0 Then
283                       sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                                & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                                & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                                & " ('" & Trim(vKodeLgn) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', 'F'," _
                                & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                                & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                                & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat, isNowBilling), "DD/MM/YYYY")) & "', 'PPN Keluar No. WO : " & Trim(Trim(text(0))) & "'," _
                                & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "'," _
                                & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
284                       Call objHT.InsertData_ACC(sSQL1)
285                   End If
          
                      'Update Armst Saldo berjalan
286                   sSQL1 = "Update Armst set SaldoJalan = SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(vKodeLgn) & "'"
287                   Call objHT.UpdateData_ACC(sSQL1)
                  'End If
              
                  'Membuat transaksi di GL
                  'If vTotal + vTMeterai <> 0 Then
                    'PIUTANG GR
288                 sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                          & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, nowo, nofps) values " _
                          & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                          & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
289                 Call objHT.InsertData_ACC(sSQL1)
290                 NilaiGLNonPart = vTotal + vTMeterai
                      'PENJ JASA
291                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                            & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                            & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                            & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                            & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
292                   Call objHT.InsertData_ACC(sSQL1)
293                   NilaiGLNonPart = NilaiGLNonPart - vTPenjNonMet
                      'PENJ SUBLET
294                   If vTPenjSublet <> 0 Then
295                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjSubletGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTPenjSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
296                       Call objHT.InsertData_ACC(sSQL1)
297                       NilaiGLNonPart = NilaiGLNonPart - vTPenjSublet
298                   End If
                      'MATERAI
299                   If vTMeterai <> 0 Then
300                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
301                       Call objHT.InsertData_ACC(sSQL1)
302                       NilaiGLNonPart = NilaiGLNonPart - vTMeterai
303                   End If
                      'PPn Jasa + Sublet
304                   If (vTPenjNonMet + vTPenjSublet) <> 0 Then
      '                    sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                          & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                          & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                          & " " & SimpanKoma2Titik(Round(((vTPenjNonMet + vTPenjSublet) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                          & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                          
                          ' REVISI THANIA 09/05/2014
                          'sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vJSPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                          
                          ' REVISI ACEP 10/06/2014
305                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                          '----------------- coba selisih -------------
306                       If adaSelisih Then
307                           If kreditSelisih Then
308                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
309                           Else
310                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
311                           End If
312                           isSudahDikuragi = True
313                       End If
                          '-----------------
314                       Call objHT.InsertData_ACC(sSQL1)
                          
                         
                                            
315                       NilaiGLNonPart = NilaiGLNonPart - (Round(((vTPenjNonMet + vTPenjSublet) / 10), 0))
316                   End If
                      
317                   If nDuoInv = True Then
                      
                      
      '                    nKodeSi = Trim(gKdPlgDealer2)
      '                    isCust = Trim(Text(40))
      '                    isExAflAR = objO.Fetch_TipeCust_Afiliasi_ARMST(nKodeSi)  'cex dari ARMST apakah AFL / not AFL
      '                    If isExAflAR = True Then  'AFL
      '                        isExAflSI = objO.Fetch_TipeCust_Afiliasi_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah AFL terdaftar
      '                        If isExAflSI = True Then
      '                            isCust = "AFFPPN"
      '                        Else
      '                            MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar di SI_JURNAL_NEW !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
      '                            GoTo vbError
      '                        End If
      '                    Else
      '                        If isCust = "AFFPPN" Then  'cex jika tidak sesuai antara ARMST dan M_Asuransi
      '                            MsgBox "Data Customer '" & nKodeSi & " - Afiliasi' tidak Terdaftar sebagai Afiliasi !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
      '                            GoTo vbError
      '                        End If
      '                        If isCust = "INT" Then
      '                            isExIntSI = objO.Fetch_TipeCust_Internal_SIJURNALNEW(nKodeSi) 'cex dari SIJURNALNEW apakah INT terdaftar
      '                            If isExIntSI = False Then
      '                                MsgBox "Data Customer '" & nKodeSi & " - Internal' tidak Terdaftar sebagai Internal !" & Chr(13) & "Silahkan Hubungi PIC Accounting", vbInformation + vbOKOnly, "Billing"
      '                                GoTo vbError
      '                            End If
      '                        End If
      '                    End If
              
                          Dim rsGL2 As ADODB.Recordset
318                       Select Case isCust
                              Case "EXT"
319                               Set rsGL2 = objO.Get_AkunGL_NEW("OTHERS", isDyna, False)
320                           Case Else  'INT/AFL
321                               Set rsGL2 = objO.Get_AkunGL_NEW(nKodeSi, isDyna, False)
322                       End Select

323                       If Not rsGL2.BOF Then
324                           vPiutangGL = cekKodeAkun(rsGL2.Fields(0).value)
325                           vPenjualanGL = cekKodeAkun(rsGL2.Fields(1).value)
326                           vPpnGL = cekKodeAkun(rsGL2.Fields(6).value)
327                           vMeteraiGL = cekKodeAkun(rsGL2.Fields(9).value)
328                       End If
329                   End If
330                   If gJurnalPart Then
                          Dim SumPart As Double
                          Dim SumOlie As Double
                          Dim SumAccs As Double
                          Dim SumPartFull As Double
                          Dim SumOlieFull As Double
                          Dim SumAccsFull As Double
                          '' NILAI SPAREPART
331                       vTPenjPart = 0
                          Dim rsSearchPart As ADODB.Recordset
332                       If rsPart.State <> adStateClosed Then
333                           SumPart = 0
334                           SumOlie = 0
335                           SumAccs = 0
336                           SumPartFull = 0
337                           SumOlieFull = 0
338                           SumAccsFull = 0
339                           vPaPPN = 0
340                           vOlPPN = 0
341                           vAcPPN = 0
342                           rsPart.MoveFirst
343                           Do Until rsPart.EOF
344                               Select Case Left(UCase(objO.FetchData_String_Acc("select B.Jenis from Inmst A WITH (NOLOCK) Inner Join Profilacc..SCC B WITH (NOLOCK) On A.SCC_Code = B.Scc_Code Where A.KodeItem = '" & Trim(rsPart("KodeItem")) & "'")), 3)
                                  Case "PAR"
345                                   SumPart = SumPart + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
346                                   SumPartFull = SumPartFull + CDbl(rsPart("HARGA"))
347                               Case "OLI"
348                                   SumOlie = SumOlie + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
349                                   SumOlieFull = SumOlieFull + CDbl(rsPart("HARGA"))
350                               Case "ACC"
351                                   SumAccs = SumAccs + Round((CDbl(rsPart("HARGA")) * 100 / 110), 0)
352                                   SumAccsFull = SumAccsFull + CDbl(rsPart("HARGA"))
353                               End Select
354                               rsPart.MoveNext
355                           Loop
356                           vTPenjPart = SumPart + SumOlie + SumAccs
                                                  
357                           If NilaiGLNonPart <> vTPenjPart Then
358                               If NilaiGLNonPart - vTPenjPart = 1 Then
359                                   If SumPart <> 0 Then
360                                       SumPart = SumPart + 1
361                                   Else
362                                       If SumOlie <> 0 Then
363                                           SumOlie = SumOlie + 1
364                                       Else
365                                           If SumAccs <> 0 Then
366                                               SumAccs = SumAccs + 1
367                                           End If
368                                       End If
369                                   End If
370                                   vTPenjPart = NilaiGLNonPart
371                               ElseIf NilaiGLNonPart - vTPenjPart = -1 Then
372                                   If SumPart <> 0 Then
373                                       SumPart = SumPart - 1
374                                   Else
375                                       If SumOlie <> 0 Then
376                                           SumOlie = SumOlie - 1
377                                       Else
378                                           If SumAccs <> 0 Then
379                                               SumAccs = SumAccs - 1
380                                           End If
381                                       End If
382                                   End If
383                                   vTPenjPart = NilaiGLNonPart
384                               End If
385                           End If
                              
386                           vPaPPN = SumPartFull - SumPart
387                           vOlPPN = SumOlieFull - SumOlie
388                           vAcPPN = SumAccsFull - SumAccs
                              
389                           If SumPart <> 0 Then
390                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                        & " " & SimpanKoma2Titik(SumPart) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                        & " '" & vDelIndex & "','" & vKeterangan & " - PART','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
391                                     Call objHT.InsertData_ACC(sSQL1)
392                               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(SumPart) & ", " _
                                        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(SumPart) & " " _
                                        & " where KodeGl = '" & vPenjPartGL & "'"
393                                     Call objHT.InsertData_ACC(sSQL1)
394                           End If
395                           If SumOlie <> 0 Then
396                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjOlieGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                        & " " & SimpanKoma2Titik(SumOlie) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                        & " '" & vDelIndex & "','" & vKeterangan & " - OLI','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
397                                     Call objHT.InsertData_ACC(sSQL1)
398                               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(SumOlie) & ", " _
                                        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(SumOlie) & " " _
                                        & " where KodeGl = '" & vPenjOlieGL & "'"
399                                     Call objHT.InsertData_ACC(sSQL1)
400                           End If
                              
401                               If (SumPart + SumOlie) <> 0 Then
      '                                sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                                      & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                                      & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                                      & " " & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                                      & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
      '                                      Call objHT.InsertData_ACC(sSQL1)
      '                                sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 1)) & ", " _
      '                                      & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & " " _
      '                                      & " where KodeGl = '" & vPpnPartGL & "'"
      '                                      Call objHT.InsertData_ACC(sSQL1)
402                                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                            & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                            & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                            & " " & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                            & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
                                            
403                                   If adaSelisih Then
404                                       If isSudahDikuragi = False Then
405                                           If kreditSelisih Then
406                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                  & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                                  & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                                  & " " & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0) - 1) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                  & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
407                                           Else
408                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                  & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                                  & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnPartGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                                  & " " & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0) + 1) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                  & " '" & vDelIndex & "','" & vKeterangan & "- PPn Part" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
409                                           End If
410                                       End If
411                                   End If
                                      
412                                   Call objHT.InsertData_ACC(sSQL1)
                                      
413                                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
                                            & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0)) & " " _
                                            & " where KodeGl = '" & vPpnPartGL & "'"
414                                   If adaSelisih Then
415                                       If isSudahDikuragi = False Then
416                                           If kreditSelisih Then
417                                               sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
                                                  & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0) - 1) & " " _
                                                  & " where KodeGl = '" & vPpnPartGL & "'"
418                                           Else
419                                               sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(Round(((SumPart + SumOlie) / 10), 0)) & ", " _
                                                  & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round((vPaPPN + vOlPPN), 0) + 1) & " " _
                                                  & " where KodeGl = '" & vPpnPartGL & "'"
420                                           End If
421                                       End If
422                                   End If
423                                   Call objHT.InsertData_ACC(sSQL1)
                                            
424                               End If

425                           If SumAccs <> 0 Then
426                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPenjAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                        & " " & SimpanKoma2Titik(SumAccs) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                        & " '" & vDelIndex & "','" & vKeterangan & " - OLI','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
427                                     Call objHT.InsertData_ACC(sSQL1)
428                               sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(SumAccs) & ", " _
                                        & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(SumAccs) & " " _
                                        & " where KodeGl = '" & vPenjAccsGL & "'"
429                                     Call objHT.InsertData_ACC(sSQL1)
                                        
      '                            sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
      '                                  & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
      '                                  & " ('" & Trim(gKodeFaktur) & Trim(Text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
      '                                  & " " & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
      '                                  & " '" & vDelIndex & "','" & vKeterangan & "- PPn Accs" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(Text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
      '                                  Call objHT.InsertData_ACC(sSQL1)
      '                            sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", " _
      '                                  & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & " " _
      '                                  & " where KodeGl = '" & vPpnAccsGL & "'"
      '                                  Call objHT.InsertData_ACC(sSQL1)
                                        
430                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                        & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPpnAccsGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                        & " " & SimpanKoma2Titik(Round((vAcPPN), 0)) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                        & " '" & vDelIndex & "','" & vKeterangan & "- PPn Accs" & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
431                                     Call objHT.InsertData_ACC(sSQL1)
432                               sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(Round(((SumAccs) / 10), 0)) & ", " _
                                        & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(Round(vAcPPN, 0)) & " " _
                                        & " where KodeGl = '" & vPpnAccsGL & "'"
433                                     Call objHT.InsertData_ACC(sSQL1)
434                           End If
435                       End If
436                   End If
                      
                      '' SUBLET
                      ' AMBIL HPP SUBLET
                      Dim vTPpnS As String
                      Dim vTotalS As Double
                      Dim vTPpn1 As String
                      Dim vTBeli1 As Double
                      Dim vTotal1 As Double
                      Dim nonJurnal As Boolean
                      Dim vKeterangan1 As String
                      
437                   vTPembSublet = 0
                      'vTPenjSublet = 0
                      'Set rsSearch = objO.FetchData("SELECT PKP, NAMABARANG, NO_WO, NO_CHASSIS, HARGA*QTY AS HPP, PRICE*QTY AS HARGAJUAL, NONJURNAL FROM JOBSUBLET a INNER JOIN T_SUPPLIER B ON A.SUPPLIER = B.KODEAP WHERE NO_WO = '" & Trim(Trim(text(0))) & "' AND NO_CHASSIS = '" & Trim(text(17)) & "' AND B.FS_LOG_DEL = ''")
438                   If rsInvSublet.State <> adStateClosed Then
439                       If rsInvSublet.EOF And Not rsInvSublet.BOF Then rsInvSublet.MoveFirst
                          
440                       If Not rsInvSublet.EOF Then
441                           Do While Not rsInvSublet.EOF
442                               nonJurnal = IIf(Trim(rsInvSublet("NONJURNAL")) = 1, True, False)
443                               If Not nonJurnal Then
444                                   vTotal1 = Round(CDbl(rsInvSublet("HPP")), 0)
445                                   vTPpn1 = Round(vTotal1 * 10 / 110, 0)
446                                   vTBeli1 = Round(vTotal1 - vTPpn1, 0)
447                                   vKeterangan1 = vKeterangan & " Sublet " '" (" & rsinvsublet("NAMABARANG") & ") " 'No. WO :" & Trim(TRIM(text(0)))
448                                   If Trim(rsInvSublet("PKP")) = "1" Then
449                                       vTPpn1 = 0
450                                       vTotal1 = Round(vTBeli1, 0)
451                                   End If
452                                   vTPembSublet = vTPembSublet + vTotal1
453                               End If
454                               rsInvSublet.MoveNext
455                           Loop
456                       End If
457                   End If

                      'If Not nonJurnal Then
458                       If vTPembSublet <> 0 Then
459                           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                    & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                    & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vHPPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                    & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                    & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
460                           Call objHT.InsertData_ACC(sSQL1)
                              
461                           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                    & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                    & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "', '" & vPDPSublet & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                    & " " & SimpanKoma2Titik(vTPembSublet) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                    & " '" & vDelIndex & "','" & vKeterangan1 & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
462                           Call objHT.InsertData_ACC(sSQL1)
463                       End If
                          
464                       If vTPembSublet <> 0 Then
                              ' HPP SUBLET TERHADAP PDP
465                           sSQL1 = "update glmst set JbulanIni = JbulanIni +" & (vTPembSublet) & ", " _
                                    & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                                    & " where KodeGl = '" & vHPPSublet & "'"
466                           Call objHT.UpdateData_ACC(sSQL1)
                              
467                           sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPembSublet) & ", " _
                                    & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPembSublet) & " " _
                                    & " where KodeGl = '" & vPDPSublet & "'"
468                           Call objHT.UpdateData_ACC(sSQL1)
469                       End If
                      'End If
                      
470                   sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
                            & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                            & " where KodeGl = '" & vPiutangGL & "'"
471                   Call objHT.UpdateData_ACC(sSQL1)
                      
472                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                            & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                            & " where KodeGl = '" & vPenjualanGL & "'"
473                   Call objHT.UpdateData_ACC(sSQL1)
                      
474                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjSublet) & ", " _
                            & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjSublet) & " " _
                            & " where KodeGl = '" & vPenjSubletGL & "'"
475                   Call objHT.UpdateData_ACC(sSQL1)
                      
476                   sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTMeterai) & ", " _
                            & " JKreditBi = JKreditBi +" & vTMeterai & " " _
                            & " where KodeGl = '" & vMeteraiGL & "'"
477                   Call objHT.UpdateData_ACC(sSQL1)
                      
      '                sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPPN) & ", " _
      '                      & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPPN) & " " _
      '                      & " where KodeGl = '" & vPpnGL & "'"
      '                Call objHT.UpdateData_ACC(sSQL1)
                  'End If
478                       sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", " _
                          & " JKreditBi = JKreditBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & " " _
                          & " where KodeGl = '" & vPpnGL & "'"
                          
479                       If adaSelisih Then
480                           If kreditSelisih Then
481                               sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", " _
                                  & " JKreditBi = JKreditBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet + 1)) & " " _
                                  & " where KodeGl = '" & vPpnGL & "'"
482                           Else
483                               sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet)) & ", " _
                                  & " JKreditBi = JKreditBi +" & SimpanKoma2Titik((Round(Koma2Titik(text(22)) + Koma2Titik(text(27)), 0)) - (vTPenjNonMet + vTPenjSublet - 1)) & " " _
                                  & " where KodeGl = '" & vPpnGL & "'"
484                           End If
485                       End If
                          
486                       Call objHT.InsertData_ACC(sSQL1)
487               If nDuoInv Then
                      ' UNTUK INVOICE KEDUA -B KE DEALER PENJUAL
                      ' MEMBUAT TRANSAKSI DUO INV DI AR
488                   vTotal = subJob
489                   vTPenjNonMet = Round(vTotal * (100 / 110), 0)  'dikurangi ppn
490                   vTMeterai = Koma2Titik(Text42.text)               'materai
                             'non ppn + meterai
                      ''2015-10-19 || AKTIF AFFPPN
                      ''2015-10-19 || AKTIF AFFPPN
491                       If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isNowBilling, "mm/dd/yyyy")) >= 0 Then
492                           If Trim(gNpwp) = Trim(gInvNPWP2) Then
                                 ''2016-03-30 || subjob sudah non ppn
                                 
493                               vTotal = SpyRound(subJob)   ''vTPenjNonMet
494                               vTPenjNonMet = vTotal
495                               vTPPN = 0     'Round(vTotal * (10 / 110), 0)
496                           Else
497                               vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
498                           End If
499                       Else
500                           vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)
501                       End If
502                   vTPenjualan = vTPenjNonMet + vTMeterai
503                   vTPPN = vTotal - vTPenjNonMet      'Round(vTotal * (10 / 110), 0)          'ppn
                      'vTPenjualan = (vTotal - vTPPN) + vTMeterai
504                   vKeterangan = gInsCust2 & ". AN/ " & vKeterangan
                      
                      
505                   If vPiutangGL = "" Or vPenjualanGL = "" Or vHPPSublet = "" Or vPDPSublet = "" Or vMeteraiGL = "" Then
506                       pesanIT = pesanIT + "Kode GL belum lengkap !"
507                       GoTo vbError
508                   End If
509                   If Trim(vTipeCustomer) <> "AFFNONPPN" And vPpnGL = "" Then
510                       pesanIT = pesanIT + "Kode GL belum lengkap !"
511                       GoTo vbError
512                   End If
              
                          ' PENJUALAN
513                       sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) " _
                                & "values " _
                                & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', 'C'," _
                                & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                                & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                                & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', 'Penjualan Kredit No. WO : " & Trim(Trim(text(0))) & "'," _
                                & " " & SimpanKoma2Titik(vTPenjualan) & ", '" & gSA & "'," _
                                & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
514                       Call objHT.InsertData_ACC(sSQL1)
                          
                          'PPN
515                       If vTPPN <> 0 Then
516                           sSQL1 = "insert into Artrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn," _
                                    & " TglJthTmp, TglTrnFaktur, TglJtpFaktur, Keterangan," _
                                    & " JumlahTrn, Kodeuser, DelIndex, Kodesumber, nowo, nopolisi, tglEntry,kodejurnal) values " _
                                    & " ('" & Trim(gKdPlgDealer2) & "', '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', 'F'," _
                                    & " '" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & FormatTgl(vTglSkr) & "'," _
                                    & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', '" & FormatTgl(vTglSkr) & "'," _
                                    & " '" & FormatTgl(Format(DateAdd("D", nRangeSyarat2, isNowBilling), "DD/MM/YYYY")) & "', 'PPN Keluar No. WO : " & Trim(Trim(text(0))) & "'," _
                                    & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "'," _
                                    & " '" & vDelIndex & "','" & gKode & "','" & Trim(Trim(text(0))) & "','" & Trim(text(16)) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & vPiutangGL & "')"
517                           Call objHT.InsertData_ACC(sSQL1)
518                       End If
              
                          'Update Armst Saldo berjalan
519                       sSQL1 = "Update Armst set SaldoJalan = SaldoJalan+" & SimpanKoma2Titik(vTPenjualan + vTPPN) & " where KodeLgn= '" & Trim(gKdPlgDealer2) & "'"
520                       Call objHT.UpdateData_ACC(sSQL1)
                      'End If
                      
                      ' MEMBUAT TRANSAKSI DUO INV DI GL
                      'If vTotal + vTMeterai <> 0 Then
521                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vPiutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTotal + vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
522                       Call objHT.InsertData_ACC(sSQL1)
                        
523                       sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vPenjualanGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                & " " & SimpanKoma2Titik(vTPenjNonMet) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                                & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
524                       Call objHT.InsertData_ACC(sSQL1)
                      
525                       If vTMeterai <> 0 Then
526                           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                    & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                    & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vMeteraiGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                    & " " & SimpanKoma2Titik(vTMeterai) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                                    & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
527                           Call objHT.InsertData_ACC(sSQL1)
528                       End If
                          
529                       If vTPPN <> 0 Then
530                           sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                    & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated, NoWo, NoFps) values " _
                                    & " ('" & Trim(gKodeFaktur) & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', '" & "C" & "', " _
                                    & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & gKdPlgDealer2 & "', " _
                                    & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Trim(text(0))) & "','" & IIf(Trim(noFaktur1) = "", Trim(noFaktur2), Trim(noFaktur1)) & "')"
531                           Call objHT.InsertData_ACC(sSQL1)
532                       End If
                          
533                       sSQL1 = "update glmst set JbulanIni = JbulanIni +" & SimpanKoma2Titik(vTotal + vTMeterai) & ", " _
                                & " JDebitBi = JDebitBi +" & SimpanKoma2Titik(vTotal + vTMeterai) & " " _
                                & " where KodeGl = '" & vPiutangGL & "'"
534                       Call objHT.UpdateData_ACC(sSQL1)
                          
535                       sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPenjNonMet) & ", " _
                                & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPenjNonMet) & " " _
                                & " where KodeGl = '" & vPenjualanGL & "'"
536                       Call objHT.UpdateData_ACC(sSQL1)
                          
537                       sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTMeterai) & ", " _
                                & " JKreditBi = JKreditBi +" & vTMeterai & " " _
                                & " where KodeGl = '" & vMeteraiGL & "'"
538                       Call objHT.UpdateData_ACC(sSQL1)
                          
539                       sSQL1 = "update glmst set JbulanIni = JbulanIni -" & SimpanKoma2Titik(vTPPN) & ", " _
                                & " JKreditBi = JKreditBi +" & SimpanKoma2Titik(vTPPN) & " " _
                                & " where KodeGl = '" & vPpnGL & "'"
540                       Call objHT.UpdateData_ACC(sSQL1)
                      'End If
541               End If
              
              
542               If erTranAll = False Then
543                   Call objHT.CloseTransactionAll
544                   Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 1)
545                   simpanACC2 = True
546               Else
547                   Call objHT.CloseTransactionAll_Err
548                   Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
549                   simpanACC2 = False
550                   MsgBox "Ada Error, Pembuatan Invoice dibatalkan, buat invoice kembali / Hubungi EDP!"
                      
551               End If
552       Else
553           Call objHT.CloseTransactionAll_Err
554           Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
555           simpanACC2 = False
556           MsgBox "Ada Error, Pembuatan Invoice dibatalkan, buat invoice kembali / Hubungi EDP!"
              ''MsgBox "Ada Error, Pembuatan Invoice dibatalkan, buat invoice kembali / Hubungi EDP!"
557       End If
558       Exit Function
          
vbError:
559       simpanACC2 = False
560       Call objHT.CloseTransactionAll_Err
561       Call saveLogMenu(gSA, ComptName, gVersi, "frmBilling", "simpanInvoice", 2)
562       If pesanIT <> "" Then
563         MsgBox pesanIT, vbInformation
564       Else
565         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanACC2 of Form frmBilling")
566       End If
          
567       Exit Function
End Function

Private Function createTemp(vNamaTabel As String) As Boolean
      Dim rsHlp As ADODB.Recordset

          'If Get_Connectiontemp("tempdb", gUser, gPwd, gServerDb) Then
1        On Error GoTo createTemp_Error

2             createTemp = False
3             Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ") ' ")
4             If Not rsHlp.EOF Then
5                 createTemp = True
6             Else
7                 Call BuatTblInv(vNamaTabel)
8                 createTemp = True
9             End If
          'End If

10       'On Error GoTo 0
11       Exit Function

createTemp_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createTemp of Form frmBilling")
End Function

Public Sub BuatTblInv(vNamaTabel As String)
1        On Error GoTo BuatTblInv_Error

2         sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ) " _
               & " drop table [tempdb]..[" & vNamaTabel & "]"
3         Call objO.UpdateData(sSQL)
          
4         sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
               & "[REV_CETAK] [decimal] (18,2) NOT NULL DEFAULT 0, "
5         sSQL = sSQL & "[REV_HITUNG] [decimal] (18,2) NOT NULL DEFAULT 0, " _
               & "[NO_CHASSIS] [varchar] (25) NOT NULL DEFAULT '', " _
               & "[NOINV] [varchar] (20) NOT NULL DEFAULT '', " _
               & "[INVDATE] [datetime] NOT NULL DEFAULT '01/01/1900', " _
               & "[FS_KD_PLG] [varchar] (15) NOT NULL DEFAULT '', " _
               & "[CUSTOMER] [varchar] (50) NOT NULL DEFAULT '', " _
               & "[CUST_ADRES] [varchar] (250) NOT NULL DEFAULT '', " _
               & "[KOTA] [varchar] (100) NOT NULL DEFAULT '', " _
               & "[NPWP] [varchar] (50) NOT NULL DEFAULT '', " _
               & "[NO_POLISI] [varchar] (15) NOT NULL DEFAULT '', " _
               & "[FMDL] [varchar] (50) NOT NULL DEFAULT '', " _
               & "[NO_ENGINE] [varchar] (15) NOT NULL DEFAULT '', " _
               & "[NO_WO] [varchar] (15) NOT NULL DEFAULT ''," _
               & "[FRONT] [varchar] (10) NOT NULL DEFAULT '', " _
               & "[STATUS] [varchar] (10) NOT NULL DEFAULT '', " _
               & "[MECH] [varchar] (10) NOT NULL DEFAULT '', " _
               & "[FS_LOG_ADD] [varchar] (10) NOT NULL DEFAULT '', " _
               & "[SBLT] [decimal] (18,2) NOT NULL DEFAULT 0, " _
               & "[PART] [decimal] (18,2) NOT NULL DEFAULT 0, " _
               & "[OLI] [decimal] (18,2) NOT NULL DEFAULT 0, " _
               & "[MTRL] [decimal] (18,2) NOT NULL DEFAULT 0, " _
               & "[MET] [decimal] (18,2) NOT NULL DEFAULT 0, " _
               & "[TOTAL] [decimal] (18,2) NOT NULL DEFAULT 0, " _
         & "[SLIP] [varchar] (100) NOT NULL DEFAULT '', "
6         sSQL = sSQL + "[FS_CATATAN] [varchar] (1000) NOT NULL DEFAULT '', " _
                      & "[REP] [varchar] (100) NOT NULL DEFAULT '', " _
                      & "[KLPKERJA] [varchar] (5) NOT NULL DEFAULT '', " _
                      & "[RTE] [varchar] (5) NOT NULL DEFAULT '', " _
                      & "[OFF] [decimal] (18,2) NOT NULL DEFAULT 0, " _
                      & "[DISC] [decimal] (18,4) NOT NULL DEFAULT 0, " _
                      & "[CTKDISC] [char] (1) NOT NULL DEFAULT 0, " _
                      & "[BP] [decimal] (18,2) NOT NULL DEFAULT 0, [FN_PAJAK] [decimal] (18,2) NOT NULL DEFAULT 0) ON [PRIMARY]"
7         Call objO.UpdateData(sSQL)

         'On Error GoTo 0
8        Exit Sub

BuatTblInv_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure BuatTblInv of Form frmBilling")
End Sub

Private Sub cetakInv1()
          Dim rsRepo As New ADODB.Recordset
          Dim rsTamp As New ADODB.Recordset
          Dim rpt As Object
1         Select Case gPrintCetak
              Case "4"
2                 Set rpt = New RptFormInvoiceN1_4           ' inv blanko kosong
3                 rpt.txtDealer.SetText support.kopOrder(1)
4                 rpt.txtAlamat.SetText support.kopOrder(2)
5                 rpt.txtTelp.SetText support.kopOrder(3)
6                 rpt.txtBank.SetText support.kopBank(0)
7             Case Else
8                 Set rpt = New RptFormInvoiceN1_3           ' inv ALL
9         End Select
          Dim hlp1 As String
          Dim hlp2 As String
          Dim hlp3 As String
          Dim hlp4 As String
          Dim hlp6 As String
          Dim jPPN As Double
          Dim vNamaTabel As String
          Dim gcat As String
          Dim infoBill As String
          
          Dim onlySublet As Boolean
          Dim total As Double
          
          ''2017-10-26
          ''--------------------------------------------------------------------
10        If Log.loyalityCustomer = 1 Then
11            Set rsTemp = objO.FetchData("SELECT * FROM RPLNASAPP..MASTERUSER WHERE isnull(idUser,'') = '" & Trim(text(2)) & "@apps.nasmoco" & "'")
12            If Not rsTemp.EOF Then
13                rpt.txtLoyalityUser.SetText Trim(rsTemp!idUser)
                   
14                rpt.txtLoyalityPassword.SetText Trim(rsTemp!Password)
15                rpt.txtKeteranganLoyality.SetText Trim(Log.ketNasApp)
16            End If
17        End If
          
          ''********************************************************************
          '' 2017-06-16 || PENANDA UNTUK UNIT NCM DITAMBAHKAN KETERANGAN NOTIF
          ''====================================================================
18        If Log.notifInvNCM = 1 Then
19            rpt.boxNCM.Suppress = True
20            rpt.txtNCM1.Suppress = True
21            rpt.txtNCM2.Suppress = True
              
22            If isJualNRM = True Then
23                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
24                If Not rsTemp.EOF Then
                  
25                    rpt.boxNCM.Suppress = False
26                    rpt.txtNCM1.Suppress = False
27                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
28                    rpt.txtNCM2.Suppress = False
29                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
      '
      '                rpt.txtNCM3.Suppress = False
      '                rpt.txtNCM3.SetText Trim(rsTemp!Text3)
      '
      '                rpt.txtNCM4.Suppress = False
      '                rpt.txtNCM4.SetText Trim(rsTemp!Text4)
      '
      '                rpt.txtNCM5.Suppress = False
      '                rpt.txtNCM5.SetText Trim(rsTemp!Text5)
                   
30                End If
31            Else
32                Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
33                If Not rsTemp.EOF Then
                  
34                    rpt.boxNCM.Suppress = False
35                    rpt.txtNCM1.Suppress = False
36                    rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                      
37                    rpt.txtNCM2.Suppress = False
38                    rpt.txtNCM2.SetText Trim(rsTemp!Text2)
39                End If
40            End If
41        End If
          
      '    If Log.notifInvNonNCM = 1 Then
      '        If isJualNRM = False Then
      '            Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
      '            If Not rsTemp.EOF Then
      '
      '                rpt.boxNCM.Suppress = False
      '                rpt.txtNCM1.Suppress = False
      '                rpt.txtNCM1.SetText Trim(rsTemp!Text1)
      '
      '                rpt.txtNCM2.Suppress = False
      '                rpt.txtNCM2.SetText Trim(rsTemp!Text2)
      '
      '                rpt.txtNCM3.Suppress = False
      '                rpt.txtNCM3.SetText Trim(rsTemp!Text3)
      '
      '                rpt.txtNCM4.Suppress = False
      '                rpt.txtNCM4.SetText Trim(rsTemp!Text4)
      '
      '                rpt.txtNCM5.Suppress = False
      '                rpt.txtNCM5.SetText Trim(rsTemp!Text5)
      '
      '            End If
      '        Else
      '            rpt.boxNCM.Suppress = True
      '            rpt.txtNCM1.Suppress = True
      '            rpt.txtNCM2.Suppress = True
      '            rpt.txtNCM3.Suppress = True
      '            rpt.txtNCM4.Suppress = True
      '            rpt.txtNCM5.Suppress = True
      '        End If
      '    End If
          
42        gKodePlg = IIf(Trim(gKodePlg) = "", text(2), gKodePlg)
43        gInvNo = text(20) & IIf(nDuoInv, "-A", "") & "/" & text(40) 'IIf(opWO(2).value, vTipeCustomer, text(40))
44        gInvTgl = TglDDMMYYYY(isnow)
45        gInvCust = IIf(gInvCust = "", text(4), gInvCust)
46        gInvAlmt = IIf(gInvAlmt = "", txt5, gInvAlmt)
47        gInvKota = IIf(gInvKota = "", text(6) & "-" & text(7), gInvKota) 'vNoTelp
48        gInvNPWP = IIf(gInvNPWP = "", "", gInvNPWP) 'Text(14)
49        vNamaTabel = "INVOICE" & Trim(text(20))
          
          ' BELUM ANTZTEIN
          
          'If Get_Connectiontemp("tempdb", gUser, gPwd, gServerDb) Then
50            sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')) " _
                   & " drop table [tempdb]..[" & vNamaTabel & "]"
51            Call objO.UpdateData(sSQL)
          'End If
          
          'gStlRubah = IIf(DateDiff("d", Format(TglYYYYMMDD(Text(1)), "MM/DD/YYYY"), Format(gTglRubah, "MM/DD/YYYY")) >= 0, True, False)
          ' INVOICE
52        sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
               & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
               & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
               & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
               & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, FN_PAJAK " _
               & "FROM T_INVOICE WITH (NOLOCK) WHERE NOINV = '" & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "' AND NO_WO = '" & Trim(Trim(text(0))) & "' ORDER BY NOINV"
           
53        Set rsTamp = objO.FetchData(sSQL)
            Dim totalA As Double
            Dim ppnB As Double
            
54        If Not rsTamp.EOF Then
55            If createTemp(vNamaTabel) Then
56                onlySublet = True
57                total = total + CDbl(rsTamp("PART")) + CDbl(rsTamp("OLI")) + CDbl(rsTamp("MTRL")) + CDbl(rsTamp("SBLT")) + IIf(nDuoInv, 0, CDbl(rsTamp("MET"))) + CDbl(rsTamp("BODY"))
58                ppnB = CDbl(rsTamp("fn_pajak"))
59                Do While Not rsTamp.EOF
60                    For i = 1 To 8
61                        hlp1 = "REP" & i
62                        hlp2 = "RTE" & i
63                        hlp3 = "OFF" & i
64                        hlp4 = "DISC" & i
65                        hlp6 = "KLPKERJA" & i
                          
66                        If Trim(rsTamp(hlp1)) <> "" Then
67                            jPPN = Round((10 / 110) * CDbl(rsTamp(hlp3)))
68                            Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES (round((" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & RepRp(CDbl(rsTamp(hlp4))) & " * 0.01))),0,1),((" & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & RepRp(CDbl(rsTamp(hlp4))) & " * 0.01)))),'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                           & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                           & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "'," & RepRp(CDbl(rsTamp("SBLT"))) & "," & RepRp(CDbl(rsTamp("PART"))) & "," & RepRp(CDbl(rsTamp("OLI"))) & "," & RepRp(CDbl(rsTamp("MTRL"))) & ", " _
                                           & "" & IIf(nDuoInv, 0, RepRp(CDbl(rsTamp("MET")))) & "," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "'," & IIf(ChkCtkRate, RepRp(CDbl(rsTamp(hlp2))), 0) & "," & Round(CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & "," & RepRp(CDbl(rsTamp(hlp4))) & ",'" & IIf(chCtkDisc, 0, 1) & "'," & CDbl(rsTamp("BODY")) & "," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
69                            total = total + Round((CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp(hlp3)) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp(hlp4)) / 100)))
70                            onlySublet = False
71                        End If
72                    Next
                      
73                    If onlySublet Then
74                        Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                       & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                       & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "'," & RepRp(CDbl(rsTamp("SBLT"))) & "," & RepRp(CDbl(rsTamp("PART"))) & "," & RepRp(CDbl(rsTamp("OLI"))) & "," & RepRp(CDbl(rsTamp("MTRL"))) & ", " _
                                       & " " & IIf(nDuoInv, 0, RepRp(CDbl(rsTamp("MET")))) & "," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Order Sublet','','0',0,0,'" & IIf(chCtkDisc, 0, 1) & "'," & CDbl(rsTamp("BODY")) & "," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
75                    End If
76                    rsTamp.MoveNext
77                Loop
78            End If
79        End If
80        If rsTamp.State <> adStateClosed Then rsTamp.Close
          
          ' ADD JOB
          'AND B.NOINV NOT LIKE '%-B'
81        Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                   & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, A.*,fn_pajak FROM HIST_JOBADD A WITH (NOLOCK) INNER JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL AND A.NO_INV = B.NOINV  WHERE A.NO_WO = '" & Trim(Trim(text(0))) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = ''")
82        If Not rsTamp.EOF Then
83            If createTemp(vNamaTabel) Then
84                rsTamp.MoveFirst
85                Do While Not rsTamp.EOF
86                    jPPN = Round((10 / 110) * CDbl(rsTamp("OFFN")))
      '                Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES ('" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                             & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                             & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "'," & RepRp(CDbl(rsTamp("SBLT"))) & "," & RepRp(CDbl(rsTamp("PART"))) & "," & RepRp(CDbl(rsTamp("OLI")) & "," & RepRp(CDbl(rsTamp("MTRL")) & ", " _
      '                             & " " & RepRp(IIf(nDuoInv, 0, CDbl(rsTamp("MET")))) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "'," & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & "," & Replace(CDbl(rsTamp("DISC")), ",", ".") & ",'" & IIf(chCtkDisc, 0, 1) & "', " & CDbl(rsTamp("BODY")) & ") ")
                      
      '                Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES ('" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                             & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                             & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
      '                             & "'" & IIf(nDuoInv, 0, CDbl(rsTamp("MET"))) & "'," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(Text(40)) = "AFFNONPPN", jPPN, 0)) & "','" & Replace(CDbl(rsTamp("DISC")), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "', '" & CDbl(rsTamp("BODY")) & "') ")
      '
      '                total = total + Round((CDbl(rsTamp("OFFN")) - IIf(Trim(Text(40)) = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(Trim(Text(40)) = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp("DISC")) / 100)))
      '
                      '' @@@@@@@@@@@@@ 2016-03-28 || penambahan fn_pajak
                      
87                    Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES (round((" & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & Replace(CDbl(rsTamp("DISC")), ",", ".") & " * 0.01))),0,1),((" & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & " * (1-(" & Replace(CDbl(rsTamp("DISC")), ",", ".") & " * 0.01)))),'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                   & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                   & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
                                   & "'" & IIf(nDuoInv, 0, CDbl(rsTamp("MET"))) & "'," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, CDbl(rsTamp("RTE")), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) & "','" & Replace(CDbl(rsTamp("DISC")), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "', '" & CDbl(rsTamp("BODY")) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
                      
88                    total = total + Round((CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(Trim(text(40)) = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp("DISC")) / 100)))
89                    rsTamp.MoveNext
90                Loop
91            End If
92        End If

93        Set rsRepo = objO.FetchData("SELECT DISTINCT *,(SBLT+PART+OLI+MTRL+MET) as nonLabor FROM [tempdb]..[" & vNamaTabel & "] ORDER BY disc desc , (SBLT+PART+OLI+MTRL+MET) asc, KLPKERJA DESC, [OFF]") 'objO.FetchData("SELECT DISTINCT * FROM [tempdb]..[" & vNamaTabel & "] ORDER BY disc desc ,KLPKERJA DESC, [OFF]")
94        rsRepo.MoveLast
95        If rsRepo.Fields(18).value = 0 And rsRepo.Fields(19).value = 0 And rsRepo.Fields(20).value = 0 Then
96            Set rsRepo = objO.FetchData("SELECT DISTINCT *,(SBLT+PART+OLI+MTRL+MET) as nonLabor FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY disc desc , (SBLT+PART+OLI+MTRL+MET) asc,KLPKERJA ASC, REP DESC, [OFF]")
97        Else
98            rsRepo.MoveFirst
99        End If

100       rpt.database.SetDataSource rsRepo
          
101       If nRC Then
              Dim rsHlp As ADODB.Recordset
              Dim nBulan As String
              Dim nSbe As String
102           nSbe = ""
103           Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
104           If Not rsHlp.EOF Then
105               nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
106               nSbe = Trim(rsHlp("SBENEXT"))
107           End If
108           rsHlp.Close
109           If Trim(nSbe) <> "" Then
                  'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah kelipatan 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah kelipatan 10.000 KM atau bulan " & nBulan)
110               gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
111               gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
112           Else
113               Call objO.InsertData("INSERT INTO ESTIMASISALAH (norangka, tglservice, kmsvc, sbesvc, kelipatan, noinv, nowo,customer, alamat,nopolisi, telpon, gsm,cdma) VALUES " & _
                                 "('" & Trim(text(17)) & "','" & Format(text(43), "MM/DD/YYYY") & "'," & CDbl(text(37)) & ",'" & Trim(text(33)) & "','" & IIf(nHeavy, 5000, 10000) & "', " & _
                                 "'" & Trim(text(20)) & "','" & Trim(Trim(text(0))) & "','" & Trim(text(4)) & "','" & Trim(txt5) & "','" & Trim(text(16)) & "','" & Trim(text(7)) & "','" & Trim(text(8)) & "','" & Trim(text(9)) & "')")
114           End If
115       End If
          Dim rsTerbilang As ADODB.Recordset
          Dim totalTerbilang As Double
          
          'Set rsTerbilang = objO.FetchData("SELECT NOINV, ROUND(SBLT+PART+OLI+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI")
          'Set rsTerbilang = objO.FetchData("SELECT total, NOINV, ROUND(SBLT+PART+OLI+MET+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI, MET,total")
          
          '--------------------------------------
          '- 2015-07-30 | GANTI QUERY TERBILANG -
          '--------------------------------------
116       If nDuoInv = True Or (text(24) > 0 And text(25) > 0) Or text(27) > 0 Then
      '        Set rsTerbilang = objO.FetchData("SELECT distinct TOTAL,A.noinv, a.rev+SBLT+round(PART+OLI,0)+MET AS REV FROM (SELECT NOINV, ROUND(SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV) A INNER JOIN [TEMPDB]..[" & vNamaTabel & "] B ON A.NOINV = B.NOINV ")
               ''2016-05-20 || pergantian query
117           Set rsTerbilang = objO.FetchData("SELECT NOINV, ROUND(SUM([OFF]*((100-DISC)/100)) + sblt + Part + oli + met,0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV,sblt,part,oli,met")
              
118       Else
119           Set rsTerbilang = objO.FetchData("SELECT total, NOINV, ROUND(SBLT+PART+OLI+MET+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI, MET,total")
120       End If
121       If Not rsTerbilang.EOF Then
122           Do While Not rsTerbilang.EOF
123               totalTerbilang = totalTerbilang + rsTerbilang("REV")
124               rsTerbilang.MoveNext
125           Loop
126           rsTerbilang.MoveFirst
      '95            If totalTerbilang <> CDbl(rsTerbilang("total")) Then
      '96                totalTerbilang = CDbl(rsTerbilang("total"))
      '97            End If
127       End If
          
          ''2016-09-06 || text booking online
128       Set rsHlp = objO.FetchData("SELECT ISNULL(INFOBILL,'') AS INFO FROM T_PROFIL WITH (NOLOCK)")
129       If Not rsHlp.EOF Then
130         infoBill = Trim(rsHlp!info)
131       End If
132       rpt.txtTerbilang.SetText Trim(Angka2Huruf(totalTerbilang, 1))
          'rpt.txtTerbilang.SetText Trim(Angka2Huruf(Round(total, 0), 1)) 'Trim(txtTotalRp)
133       rpt.txtManager.SetText Trim(gSecmanager)
134       rpt.txtKet.SetText Trim(gcat)
135       rpt.txtInfoBilling.SetText Trim(infoBill)
          
          ' ==============================================================================================
          ' 08 Maret 2017 | Thania | penambahan informasi PWC utk cetakkan invoice & Cetak Ulang
          ' ==============================================================================================
              'rpt.txtRetur.SetText ""
136           rpt.txtRetur.SetText IIf(chPWC.value, "PWC", "")
          ' ==============================================================================================
          
137       rpt.txtJT.SetText "Jatuh Tempo : " & JTBayar(gKodePlg, gInvTgl)
          
138       With CRViewer1
139           .ReportSource = rpt
140           .ViewReport
141           .Zoom 100
142           .Top = 0
143           .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
144           .Width = Frame2.Width + 100
145           .Height = asxPowerBanner1.Height
              
146           cmdTutup.Top = .Top + 430
147           cmdTutup.Left = .Left
148           cmdTutup.Visible = True
149           .Visible = True
150       End With
151       cmdTutup.Visible = True
152       Screen.MousePointer = vbDefault
153       Exit Sub
End Sub


Private Sub cetakInv2()
          'Dim rpt As New RptFormInvoiceN2_1      ' INV KGW   '23 AGUSTUS 2010 GANTI FORMAT BARU
          'Dim rpt As New RptFormInvoiceN2_2      ' inv mlati
          Dim rpt As Object
          Dim met As Double
1         met = CDbl(Trim(Text42.text))
          
2         Select Case gPrintCetak
              Case "4"
3                 Set rpt = New RptFormInvoiceN2_4           ' inv blanko kosong
4                 rpt.txtDealer.SetText support.kopOrder(1)
5                 rpt.txtAlamat.SetText support.kopOrder(2)
6                 rpt.txtTelp.SetText support.kopOrder(3)
7                 rpt.txtBank.SetText support.kopBank(0)
8             Case Else
      '            Set rpt = New RptFormInvoiceN2_3           ' inv ALL
9                 If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) >= 0 Then
10                    If Trim(gNpwp) = Trim(gInvNPWP2) Then
11                        Set rpt = New RptFormInvoiceN2_3_nonPPn
12                    Else
13                        Set rpt = New RptFormInvoiceN2_3           ' inv ALL
14                    End If
15                Else
16                    Set rpt = New RptFormInvoiceN2_3           ' inv ALL
17                End If
18        End Select
          Dim rsRepo As ADODB.Recordset
          Dim rsTamp As ADODB.Recordset
          Dim hlp1 As String
          Dim hlp2 As String
          Dim hlp3 As String
          Dim hlp4 As String
          Dim hlp6 As String
          Dim hlp7 As String
          Dim jPPN As Double
          Dim sTipeCust As String
          Dim vNamaTabel As String
          Dim gcat As String
          Dim onlySublet As Boolean
          Dim total As Double
          
          ''2015-10-23 || ambil nama customer dari t_invoice
19        Set rsRepo = objO.FetchData("SELECT CUSTOMER,CUST_ADRES,KOTA,STATUS,FS_KD_PLG,NPWP FROM T_INVOICE WHERE NOINV = '" & text(20) & IIf(nDuoInv, "-B", "") & "'")
20        If Not rsRepo.EOF Then
21            rsRepo.MoveFirst
              
22            sTipeCust = Trim(rsRepo!Status) ''IIf(vTipeCustomer = "", text(40), vTipeCustomer)
23            gKodePlg = Trim(rsRepo!FS_KD_PLG) ''IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
24            gInvNo = text(20) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
25            gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
26            gInvCust = Trim(rsRepo!CUSTOMER) ''IIf(gInvCust2 = "", text(4), gInvCust2)
27            gInvAlmt = Trim(rsRepo!cust_adres) ''IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
28            gInvKota = Trim(rsRepo!Kota) ''IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
29            gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
30        Else
31            sTipeCust = IIf(vTipeCustomer = "", text(40), vTipeCustomer)
32            gKodePlg = IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
33            gInvNo = text(20) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
34            gInvTgl = vTglInv 'TglDDMMYYYY(isnow) '
35            gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
36            gInvAlmt = IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
37            gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
38            gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
39        End If
      '    sTipeCust = IIf(vTipeCustomer = "", text(40), vTipeCustomer)
      '    gKodePlg = IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
      '    gInvNo = text(20) & IIf(nDuoInv, "-B", "") & "/" & IIf(vTipeCustomer = "", text(40), vTipeCustomer)
      '    gInvTgl = TglDDMMYYYY(isnow) 'vTglInv
      '    gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
      '    gInvAlmt = IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
      '    gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
      '    gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
40        vNamaTabel = "INVOICE2-" & Trim(text(20))
          
41        If Get_Connectiontemp("tempdb", gUser, gPwd, gServerDb) Then
42            sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ) " _
                   & " drop table [tempdb]..[" & vNamaTabel & "]"
43            objO.UpdateData (sSQL)
44        End If
          
45        total = 0
          'gStlRubah = IIf(DateDiff("d", Format(TglYYYYMMDD(text(1)), "MM/DD/YYYY"), Format(gTglRubah, "MM/DD/YYYY")) >= 0, True, False)
46        sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
               & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
               & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
               & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
               & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, " _
               & "FREEC1, FREEC2, FREEC3, FREEC4, FREEC5, FREEC6, FREEC7, FREEC8, FN_PAJAK " _
               & "FROM T_INVOICE WITH (NOLOCK) WHERE NOINV = '" & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "' AND NO_WO = '" & Trim(Trim(text(0))) & "'"
          
          Dim PPn As Double
          
47        Set rsTamp = objO.FetchData(sSQL)
48        If Not rsTamp.EOF Then
49            If createTemp(vNamaTabel) Then
50                onlySublet = True
51                total = total + (CDbl(rsTamp("MET")))
52                PPn = CDbl(rsTamp!fn_pajak)
                  
53                Do While Not rsTamp.EOF
54                    For i = 1 To 8
55                        hlp1 = "REP" & i
56                        hlp2 = "RTE" & i
57                        hlp3 = "OFF" & i
58                        hlp4 = "DISC" & i
59                        hlp6 = "KLPKERJA" & i
60                        hlp7 = "FREEC" & i
                          
61                        If Trim(rsTamp(hlp1)) <> "" Then
62                            jPPN = Round((10 / 110) * CDbl(rsTamp(hlp3)))
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
63                            Call objO.InsertData("INSERT INTO[tempdb]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                           & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                           & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & CDbl(rsTamp("SBLT")) & "','" & CDbl(rsTamp("PART")) & "','" & CDbl(rsTamp("OLI")) & "','" & CDbl(rsTamp("MTRL")) & "', " _
                                           & "'" & CDbl(met) & "'," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','" & Trim(rsTamp(hlp1)) & "','" & Trim(rsTamp(hlp6)) & "','" & IIf(ChkCtkRate, CDbl(rsTamp(hlp2)), "") & "','" & Round(CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & setKoma(CDbl(rsTamp(hlp4))) & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
64                            total = total + Round((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp(hlp4)) / 100)))
65                            onlySublet = False
66                        End If
67                    Next
                      
68                    If onlySublet Then
69                        Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                       & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                       & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & CDbl(rsTamp("SBLT")) & "','" & CDbl(rsTamp("PART")) & "','" & CDbl(rsTamp("OLI")) & "','" & CDbl(rsTamp("MTRL")) & "', " _
                                       & "'" & CDbl(met) & "'," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Order Sublet','','0','0','0','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
70                    End If
71                    rsTamp.MoveNext
72                Loop
73            End If
74        End If
          
          ' ADD JOB
          'AND B.NOINV NOT LIKE '%-B'
75        Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                   & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, fn_pajak , TOTAL, A.* FROM HIST_JOBADD A WITH (NOLOCK) INNER JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  AND A.NO_INV = B.NOINV  WHERE A.NO_WO = '" & Trim(Trim(text(0))) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = ''")
76        If Not rsTamp.EOF Then
77            If createTemp(vNamaTabel) Then
78                Do While Not rsTamp.EOF
79                    jPPN = Round((10 / 110) * CDbl(rsTamp("OFFN")))
                      '--:: 25 AGUSTUS 2010  ::--
      '                Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
      '                             & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
      '                             & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
      '                             & "'" & RepRp(CDbl(met)) & "'," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, setKoma(CDbl(rsTamp("RTE"))), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "'," & RepRp(CDbl(rsTamp("fn_pajak"))) & ") ")
                                   
                     ''2016-05-20
80                   Call objO.InsertData("INSERT INTO [tempdb]..[" & vNamaTabel & "] VALUES (0,0,'" & Trim(rsTamp("NO_CHASSIS")) & "','" & Trim(rsTamp("NOINV")) & "','" & Format(rsTamp("INVDATE"), "MM/DD/YYYY") & "','" & Trim(rsTamp("FS_KD_PLG")) & "','" & Trim(rsTamp("CUSTOMER")) & "','" & Trim(rsTamp("CUST_ADRES")) & "', " _
                                   & "'" & Trim(rsTamp("KOTA")) & "','" & Trim(rsTamp("NPWP")) & "','" & Trim(rsTamp("NO_POLISI")) & "','" & Trim(rsTamp("FMDL")) & "','" & Trim(rsTamp("NO_ENGINE")) & "','" & Trim(rsTamp("NO_WO")) & "','" & Trim(rsTamp("FRONT")) & "', " _
                                   & "'" & Trim(rsTamp("STATUS")) & "','" & Trim(rsTamp("MECH")) & "','" & Trim(rsTamp("FS_LOG_ADD")) & "','" & RepRp(CDbl(rsTamp("SBLT"))) & "','" & RepRp(CDbl(rsTamp("PART"))) & "','" & RepRp(CDbl(rsTamp("OLI"))) & "','" & RepRp(CDbl(rsTamp("MTRL"))) & "', " _
                                   & "'" & RepRp(CDbl(met)) & "'," & CDbl(rsTamp("TOTAL")) & ",'" & Trim(rsTamp("SLIP")) & "','" & Trim(rsTamp("FS_CATATAN")) & "','Add. " & Trim(rsTamp("REP")) & "','" & Trim(rsTamp("KLPKERJA")) & "','" & IIf(ChkCtkRate, setKoma(CDbl(rsTamp("RTE"))), "") & "','" & Round(CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) & "','" & Replace(100 - IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC"))), ",", ".") & "','" & IIf(chCtkDisc, 0, 1) & "','" & CDbl(rsTamp("BODY")) & "'," & RepRp(CDbl(PPn)) & ") ")
                     
81                    total = total + Round((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (100 - IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC")))) / 100))
82                    rsTamp.MoveNext
83                Loop
84            End If
85        End If
          
86        Set rsRepo = objO.FetchData("SELECT DISTINCT * FROM [tempdb]..[" & vNamaTabel & "] ORDER BY KLPKERJA DESC,[OFF]")
87        rpt.database.SetDataSource rsRepo
          
88        If nRC Then
              Dim rsHlp As ADODB.Recordset
              Dim nBulan As String
              Dim nSbe As String
89            nSbe = ""
90            Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(text(17)) & "' AND TGLSVCNOW = '" & Format(text(43), "MM/DD/YYYY") & "'")
91            If Not rsHlp.EOF Then
92                nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
93                nSbe = Trim(rsHlp("SBENEXT"))
94            End If
              
95            If Trim(nSbe) <> "" Then
96                gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
97                gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
98            End If
99        End If
          Dim rsTerbilang As ADODB.Recordset
          Dim totalTerbilang As Double
          'Set rsTerbilang = objO.FetchData("SELECT NOINV, ROUND(SBLT+PART+OLI+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI")
          'Set rsTerbilang = objO.FetchData("SELECT total, NOINV, ROUND(SBLT+PART+OLI+MET+SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV, SBLT, PART, OLI, MET,total")
          
          '------------------------------------
          '2015-07-30 | GANTI QUERY TERBILANG -
          '------------------------------------
100       Set rsTerbilang = objO.FetchData("SELECT distinct met,fn_pajak,total, A.noinv, a.rev+SBLT+round(PART+OLI,0)+MET AS REV FROM (SELECT NOINV, ROUND(SUM([OFF]*((100-DISC)/100)),0) AS REV from [TEMPDB]..[" & vNamaTabel & "] GROUP  BY NOINV) A INNER JOIN [TEMPDB]..[" & vNamaTabel & "] B ON A.NOINV = B.NOINV WHERE SBLT+PART+OLI = 0 and DISC = 0")

101       If Not rsTerbilang.EOF Then
102           Do While Not rsTerbilang.EOF
103               totalTerbilang = totalTerbilang + rsTerbilang("REV")
104               rsTerbilang.MoveNext
105           Loop
106           rsTerbilang.MoveFirst
107           If totalTerbilang <> CDbl(rsTerbilang("total")) Then
                  'totalTerbilang = CDbl(rsTerbilang("total"))
108           End If
109       End If
      '109       If DateDiff("d", Format(Log.StartAffPpn, "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) >= 0 Then
      '110           If Trim(gNpwp) = Trim(gInvNPWP2) Then
      '                  ''hilangkan ppn afiliasi
      '111               totalTerbilang = totalTerbilang - CDbl(rsTerbilang!MET)
      '112               totalTerbilang = Round((totalTerbilang * 10 / 11), 0) ''totalTerbilang * 10 / 11
      '
      '113           End If
      '114       End If
110       rpt.txtTerbilang.SetText Trim(Angka2Huruf(totalTerbilang, 1))
          
          'rpt.txtTerbilang.SetText Trim(Angka2Huruf(Round(total, 0), 1)) 'Trim(txtTotalRp)
111       rpt.txtManager.SetText Trim(gSecmanager)
112       rpt.txtKet.SetText Trim(gcat)
113       rpt.txtRetur.SetText ""
114       rpt.txtJT.SetText "Jatuh Tempo : " & JTBayar(gKodePlg, gInvTgl)
          
115       With CRViewer2
116           .ReportSource = rpt
117           .ViewReport
118           .Zoom 100
119           .Top = 0
120           .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
121           .Width = Frame2.Width + 100
122           .Height = asxPowerBanner1.Height
              
123           cmdTutup2.Top = .Top + 430
124           cmdTutup2.Left = .Left
125           cmdTutup2.Visible = True
126           .Visible = True
127       End With
128       cmdTutup2.Visible = True
129       Screen.MousePointer = vbDefault
130       Exit Sub
End Sub

Private Sub Cetak()
    Dim rpt As New RptFromInvoice
    Dim gInvUrai(8) As String
    Dim gInvRt(8) As String
    Dim gInvRp(8) As String
    Dim gInvRup(8) As String
    Dim gcat As String
    Dim gDisc As String
    Dim gDiscount(8) As String
    Dim gRateAdd As Double
    Dim gRpAdd As Double
    Dim gRupAdd As Double
    Dim adaAddJob As Boolean
    
    
    gKodePlg = IIf(Trim(gKodePlg) = "", text(2), gKodePlg)
    gInvNo = text(20) & IIf(nDuoInv, "-A", "") & "/" & text(40)
    gInvTgl = TglDDMMYYYY(isnow)
    gInvCust = IIf(gInvCust = "", text(4), gInvCust)
    gInvAlmt = IIf(gInvAlmt = "", txt5, gInvAlmt)
    gInvKota = IIf(gInvKota = "", text(6) & "-" & text(7), gInvKota) 'vNoTelp
    gInvNPWP = IIf(gInvNPWP = "", "", gInvNPWP) 'Text(14)
    gInvNoPolisi = text(16)
    gInvModel = text(12) & " (" & text(13) & ")"
    gInvNoMesin = text(18)
    gInvNoChas = text(17)
    gInvNoOrder = Trim(text(0))
    gInvSA = text(19)
    gInvStatus = text(40)
    gInvMech = text(38)
    gInvPrint = Format(isnow, "HH:NN")
    gInvBilling = gSA
    
    For i = 1 To 7
        gInvRt(i) = ""
        gInvUrai(i) = ""
        gInvRup(i) = ""
        gDiscount(i) = 0
    Next
    gDisc = ""
    gRateAdd = 0
    gRupAdd = 0
    gRpAdd = 0
    adaAddJob = False
    
    If rsJob.State <> adStateClosed Then
        rsJob.MoveFirst
        i = 1
        Do While Not rsJob.EOF
            If i < 8 Then
            'If Not rsJob("ADDJOB") Then
                gInvUrai(i) = IIf(rsJob("ADDJOB"), "Add. ", "") & rsJob("JOB")
                gInvRt(i) = rsJob("RATE")
                gInvRp(i) = Format(CDbl(rsJob("HRG")), "#,#0")
                gInvRup(i) = Format(CDbl(rsJob("HRG")) * ((100 - rsJob("DISC")) / 100), "#,#0")
                gDiscount(i) = Round(CDbl(rsJob("DISC")), 2)
                i = i + 1
            End If
            rsJob.MoveNext
        Loop
    End If
    
    'text(29) = text(29) - IIf(nDuoInv, CDbl(text42.text), 0)
    gInvTRp1 = text(22)     ' JASA
    gInvTRp2 = text(27)     ' SUBLET
    gInvTRp3 = text(24)     ' PART
    gInvTRp4 = text(25)     ' OLI
    gInvTRp5 = text(28)     ' MATERIAL
    gInvTRp6 = IIf(nDuoInv, 0, Text42.text)    ' METERAI
    gInvTRp7 = text(39)     ' BODY REPAIR
    gGInvTRp = Format(text(29), "#,##0")
    gInvTerbilang = txtTotalRp
    gInvSS = text(32)
    gInvCatatan = text(30)
    gInvSM = gSecmanager
    
    If nRC Then
        Dim rsHlp As ADODB.Recordset
        Dim nBulan As String
        Set rsHlp = objO.FetchData("SELECT JT_TGL, JT_BLN FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & gInvNoChas & "' AND TGLSVCNOW = '" & Format(isnow, "MM/DD/YYYY") & "'")
        If Not rsHlp.EOF Then
            nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
        End If
        gcat = IIf(nHeavy, "Kembali untuk service berkala setelah 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah 10.000 KM atau bulan " & nBulan)
    End If
    
    rpt.Field65.DataValue = gKodePlg
    rpt.Field39.DataValue = gInvNo
    rpt.Field40.DataValue = gInvTgl

    rpt.Field1.DataValue = gInvCust
    rpt.Field4.DataValue = gInvAlmt
    rpt.Field28.DataValue = gInvKota
    rpt.Field7.DataValue = gInvNPWP

    rpt.Field2.DataValue = gInvNoPolisi
    rpt.Field5.DataValue = gInvModel
    rpt.Field12.DataValue = gInvNoMesin
    rpt.Field14.DataValue = gInvNoChas
    
    rpt.Field29.DataValue = gInvNoOrder
    rpt.Field30.DataValue = gInvSA

    rpt.Field31.DataValue = gInvStatus
    rpt.Field32.DataValue = gInvMech

    rpt.Field33.DataValue = gInvPrint
    rpt.Field38.DataValue = gInvBilling
    
    rpt.Field16.DataValue = gInvUrai(1)
    rpt.Field17.DataValue = gInvUrai(2)
    rpt.Field18.DataValue = gInvUrai(3)
    rpt.Field19.DataValue = gInvUrai(4)
    rpt.Field20.DataValue = gInvUrai(5)
    rpt.Field21.DataValue = gInvUrai(6)
    rpt.Field22.DataValue = gInvUrai(7)
    rpt.Field74.DataValue = IIf(cmdSublet.Visible, "SUBLET", "") 'IIf(adaAddJob, "Additional Job", "")
    
    If ChkCtkRate = Checked Then
        rpt.Field41.DataValue = gInvRt(1)
        rpt.Field42.DataValue = gInvRt(2)
        rpt.Field43.DataValue = gInvRt(3)
        rpt.Field44.DataValue = gInvRt(4)
        rpt.Field45.DataValue = gInvRt(5)
        rpt.Field46.DataValue = gInvRt(6)
        rpt.Field47.DataValue = gInvRt(7)
        rpt.Field75.DataValue = "" 'IIf(adaAddJob, gRateAdd, "")
    Else
        rpt.Field41.DataValue = ""
        rpt.Field42.DataValue = ""
        rpt.Field43.DataValue = ""
        rpt.Field44.DataValue = ""
        rpt.Field45.DataValue = ""
        rpt.Field46.DataValue = ""
        rpt.Field47.DataValue = ""
        rpt.Field75.DataValue = ""
    End If
    
    If chCtkDisc = Unchecked Then
        ' CETAK DISC
        rpt.Field67.DataValue = IIf(gInvUrai(1) <> "", gDiscount(1) & " %", "")
        rpt.Field68.DataValue = IIf(gInvUrai(2) <> "", gDiscount(2) & " %", "")
        rpt.Field69.DataValue = IIf(gInvUrai(3) <> "", gDiscount(3) & " %", "")
        rpt.Field70.DataValue = IIf(gInvUrai(4) <> "", gDiscount(4) & " %", "")
        rpt.Field71.DataValue = IIf(gInvUrai(5) <> "", gDiscount(5) & " %", "")
        rpt.Field72.DataValue = IIf(gInvUrai(6) <> "", gDiscount(6) & " %", "")
        rpt.Field73.DataValue = IIf(gInvUrai(7) <> "", gDiscount(7) & " %", "")
        rpt.Field77.DataValue = "" 'IIf(adaAddJob, gDisc & " %", "")
        
        rpt.Field78.DataValue = gInvRp(1)
        rpt.Field79.DataValue = gInvRp(2)
        rpt.Field80.DataValue = gInvRp(3)
        rpt.Field81.DataValue = gInvRp(4)
        rpt.Field82.DataValue = gInvRp(5)
        rpt.Field83.DataValue = gInvRp(6)
        rpt.Field84.DataValue = gInvRp(7)
        rpt.Field85.DataValue = "" 'IIf(adaAddJob, Format(gRpAdd, "#,##0"), "")
    Else
        ' TIDAK CETAK DISC
        rpt.Field67.DataValue = ""
        rpt.Field68.DataValue = ""
        rpt.Field69.DataValue = ""
        rpt.Field70.DataValue = ""
        rpt.Field71.DataValue = ""
        rpt.Field72.DataValue = ""
        rpt.Field73.DataValue = ""
        rpt.Field77.DataValue = ""
        
        rpt.Field78.DataValue = ""
        rpt.Field79.DataValue = ""
        rpt.Field80.DataValue = ""
        rpt.Field81.DataValue = ""
        rpt.Field82.DataValue = ""
        rpt.Field83.DataValue = ""
        rpt.Field84.DataValue = ""
        rpt.Field85.DataValue = ""
    End If
    
    rpt.Field48.DataValue = gInvRup(1)
    rpt.Field49.DataValue = gInvRup(2)
    rpt.Field50.DataValue = gInvRup(3)
    rpt.Field51.DataValue = gInvRup(4)
    rpt.Field52.DataValue = gInvRup(5)
    rpt.Field53.DataValue = gInvRup(6)
    rpt.Field54.DataValue = gInvRup(7)
    rpt.Field76.DataValue = "" 'IIf(adaAddJob, Format(gRupAdd, "#,##0"), "")
    
    rpt.Field55.DataValue = gInvTRp1
    rpt.Field56.DataValue = gInvTRp2
    rpt.Field57.DataValue = gInvTRp3
    rpt.Field58.DataValue = gInvTRp4
    rpt.Field59.DataValue = gInvTRp5
    rpt.Field60.DataValue = gInvTRp6
    rpt.Field61.DataValue = gInvTRp7
    
    rpt.Field23.DataValue = "" 'gGInvRp
    rpt.Field27.DataValue = gGInvTRp
    rpt.Field25.DataValue = gInvTerbilang
    
    rpt.Field62.DataValue = gInvSS
    rpt.Field63.DataValue = gInvCatatan
    rpt.Field64.DataValue = gInvSM
    rpt.Field66.DataValue = gcat
    '9=A4
    'kwarto

    rpt.Printer.PaperSize = 256
    rpt.Printer.PaperHeight = 1440 * 11
    rpt.Printer.PaperWidth = 1440 * 8.5
    
    rpt.Printer.Orientation = 1
    
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show
    End If
End Sub

Private Sub CetakFormKecil()
    Dim rpt As New RptInvoice
    
    rpt.Label39 = gSecmanager
    rpt.Label40 = "KEPALA BENGKEL"
    rpt.Label2 = "NPWP : " & gNpwp
    rpt.Label3 = "PKP  : " & gPkp
    rpt.Label44 = "TELP : " & gTelp
    rpt.Label18 = "BIAYA JASA  : " & AddSpace(Format(text(22), "#,##0"), 13, 3, " ")
    rpt.Label19 = "SUBLET      : " & AddSpace(Format(text(27), "#,##0"), 13, 3, " ")
    rpt.Label20 = "PARTS       : " & AddSpace(Format(text(24), "#,##0"), 13, 3, " ")
    rpt.Label21 = "OIL/GRERASE : " & AddSpace(Format(text(25), "#,##0"), 13, 3, " ")
    rpt.Label22 = "MATERIAL    : " & AddSpace(Format(text(28), "#,##0"), 13, 3, " ")
    rpt.Label23 = "BODY REPAIR : " & AddSpace(Format(text(39), "#,##0"), 13, 3, " ")
    rpt.Label24 = "METERAI     : " & AddSpace(Format(Text42.text, "#,##0"), 13, 3, " ")
    rpt.Label25 = "GRAND TOTAL : " & AddSpace(Format(text(29), "#,##0"), 13, 3, " ")
    
    rpt.Label26 = AddSpace("URAIAN BIAYA JASA", 25, 1, " ") + _
                AddSpace("MKN", 4, 3, " ") + _
                AddSpace("RUPIAH", 12, 3, " ")
    rpt.Label27 = AddSpace(DGrid.TextMatrix(1, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(1, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(1, 4), "#,##0"), 12, 3, " ")
    rpt.Label28 = AddSpace(DGrid.TextMatrix(2, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(2, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(2, 4), "#,##0"), 12, 3, " ")
    rpt.Label29 = AddSpace(DGrid.TextMatrix(3, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(3, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(3, 4), "#,##0"), 12, 3, " ")
    rpt.Label30 = AddSpace(DGrid.TextMatrix(4, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(4, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(4, 4), "#,##0"), 12, 3, " ")
    rpt.Label31 = AddSpace(DGrid.TextMatrix(5, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(5, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(5, 4), "#,##0"), 12, 3, " ")
    rpt.Label32 = AddSpace(DGrid.TextMatrix(6, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(6, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(6, 4), "#,##0"), 12, 3, " ")
    rpt.Label33 = AddSpace(DGrid.TextMatrix(7, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(7, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(7, 4), "#,##0"), 12, 3, " ")
    
    rpt.Label8 = "NO. WO. : " & text(20)
    rpt.Label9 = "FRAME : " & text(17)
    rpt.Label10 = "E/G : " & text(18)
    rpt.Label11 = "NAMA : " & text(4)
    rpt.Label12 = "ALAMAT : " & txt5
    rpt.Label13 = "KOTA : " & text(6)
    rpt.Label14 = text(14) & "(" & text(13) & ")"
    rpt.Label15 = "SA : " & text(19)
    rpt.Label16 = "OPT : " & gSA
    
    rpt.Label47 = "Kode Plgn  : " & text(2)
    rpt.Label4 = "No. Inv    : " & text(20) & "/" & text(40)
    rpt.Label5 = "Tanggal    : " & vTglInv
    rpt.Label6 = "Jam        : " & Format(isnow, "HH:NN")
    rpt.Label7 = "No. Polisi : " & text(16)
    
    rpt.Label36 = txtTotalRp
    rpt.Label37 = text(32)
    rpt.Label38 = text(30)
    
    rpt.Label45 = "Bank     : " & gBank
    rpt.Label46 = "Acc. No. : " & gAcc
    
'    TRIM(text(0)) = ""
    
    rpt.Printer.PaperWidth = 1440 * 8.5
    rpt.Printer.PaperHeight = 1440 * 5.5
    rpt.Printer.Orientation = 1
    
    Call hapusForm
    Trim(text(0)) = ""
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show
    End If
End Sub

Private Sub cetakAdd_Click()
Set Report = New rptAddInvoice
Dim rsRepo As ADODB.Recordset
    
    If rsRepo.State Then rsRepo.Close
    sSQL = "SELECT A.NO_WO,NOINV, A.NO_CHASSIS, NO_POLISI, FMDL, KLPKERJA, REP, RTE, MEKANIK, OFFN, A.DISC, CUSTOMER, CUST_ADRES, KOTA, NO_ENGINE " _
         & "FROM JOBADD2 A WITH (NOLOCK) INNER JOIN JOBDTA B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS " _
         & "WHERE NOINV = '" & text(20) & "' AND A.NO_WO = '" & Trim(text(0)) & "' ORDER BY A.NO_WO, KLPKERJA, REP"
    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    Report.txtPT.SetText Trim(gPerush)
    Report.txtAlamat.SetText Trim(gAlamat)
    Report.txtKota.SetText Trim(gKotaCbg)
    Report.txtTelp.SetText "Telp. " & Trim(gTelp)
    Report.txtFax.SetText ""  '"Fax. "
    Report.txtUser.SetText Trim(gSA)
    Report.txtKot.SetText Trim(gKotaCbg)
    Report.txtKabeng.SetText Trim(gSecmanager)

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        
        .Top = 0
        .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
        .Width = Frame2.Width + 100
        .Height = Me.Height
        
        cmdTutup.Top = .Top + 430
        cmdTutup.Left = .Left
        cmdTutup.Visible = True
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
    cmdTutup.SetFocus
End Sub

Private Sub cetakSublet_Click()
Set Report = New rptSublet
Dim rsRepo As ADODB.Recordset
    '#########################################################'
    '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
    '#########################################################'
    'sSQL = "SELECT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, NOINV, A.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, PRICE, QTY, KETERANGAN " _
         & "FROM JOBSUBLET A INNER JOIN JOBDTA B ON A.NO_WO = B.NO_WO AND LEFT(A.NO_INV,8) = B.NOINV " _
         & "WHERE A.NO_INV LIKE '" & Text(20) & "%' AND A.NO_WO = '" & Trim(Text(0)) & "' AND B.NOINV NOT LIKE '%-B' ORDER BY A.NO_WO, NAMABARANG "

    sSQL = "SELECT CUSTOMER, CUST_ADRES, KOTA, A.NO_WO, A.NOINV, B.NO_POLISI, A.NO_CHASSIS, NO_ENGINE, NAMABARANG, HARGA, QTY, KETERANGAN " _
         & "FROM HIST_SUBLET A WITH (NOLOCK) INNER JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NOINV = B.NOINV  " _
         & "WHERE A.NOINV LIKE '" & text(20) & "%' AND B.NOINV NOT LIKE '%-B' AND A.NO_WO = '" & text(0) & "' ORDER BY A.NO_WO, NAMABARANG "


    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    Report.txtPT.SetText Trim(gPerush)
    Report.txtAlamat.SetText Trim(gAlamat)
    Report.txtKota.SetText Trim(gKotaCbg)
    Report.txtTelp.SetText "Telp. " & Trim(gTelp)
    Report.txtFax.SetText ""  '"Fax. "
    Report.txtKotaTgl.SetText Trim(gKotaCbg) & ", " & TglDDMMMMYYYY(Format(isnow, "DD/MM/YYYY"), 1)
    Report.txtKabeng.SetText Trim(gSecmanager)

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        
        .Top = 0
        .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
        .Width = Frame2.Width + 100
        .Height = TblKeluar.Top + TblKeluar.Height
        
        cmdTutup.Top = .Top + 430
        cmdTutup.Left = .Left
        cmdTutup.Visible = True
        
        cmdSettingPrinter.Top = cmdTutup.Top
        cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
        cmdSettingPrinter.Visible = True
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
    cmdTutup.SetFocus
End Sub

Private Sub cmdCetak_Click()
    If Not simpan Then
        MsgBox "Data belum simpan. "
        Exit Sub
    Else
        If gUkuranInvoice = False Then
            Call Cetak
        Else
            Call CetakFormKecil
        End If
    End If
End Sub

Private Sub cmdLengkap_Click()
    frmLengkap.Top = cmdLengkap.Top + cmdLengkap.Height + 50
    frmLengkap.Left = cmdLengkap.Left
    frmLengkap.Visible = True
    For i = 0 To 1
        opLengkap(i).Enabled = False
        text(115 + i).Enabled = False
    Next
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdSublet_Click()
    gTag = "billing"
    With frmOSublet
        .text(6) = Trim(text(0))
        .text(0) = Trim(text(17))
        .text(1) = Trim(text(12))
        .text(2) = Trim(text(13))
        .text(3) = Trim(text(14))
        .text(4) = Trim(text(4))
        .text(5) = Trim(text(5))
        .text(14) = Trim(text(16))
        .text(15) = Trim(text(31))
        .cmdKeluar.Caption = "Kembali ke &Billing"
        .text(13).Visible = True
        .Label11.Visible = True
        .Show vbModal
    End With
    text(100).SetFocus
End Sub

Private Sub cmdTutup_Click()
Dim vNamaTabel As String

    vNamaTabel = "INVOICE" & Trim(text(20))
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.DeleteData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    text(100).SetFocus
End Sub

Private Sub cmdTutup2_Click()
Dim vNamaTabel As String

    vNamaTabel = "INVOICE2-" & Trim(text(20))
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]') ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.DeleteData(sSQL)
    CRViewer2.Visible = False
    cmdTutup2.Visible = False
    text(100).SetFocus
End Sub

Private Sub Form_Load()
    onTimer = False
    Call hapusForm
    text(1) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
    simpan = True
    Call gTglServer
    Call CexRangeFP
   ' objHT As New oClassHOTrans
End Sub

Private Sub CexRangeFP()
          Dim objIn As New BClsBilling
1        On Error GoTo CexRangeFP_Error

2         isFP_OK = True
3         If objIn.FetchAvailable_FPNew = False Then
4             isFP_OK = False
5         End If
6         isNoFP_OK = True
7         isNoFP_New = objIn.FetchRange_FPNew
8         If objIn.FetchMinRange_FPNew >= isNoFP_New Then
9             isNoFP_OK = False
10        End If

11       'On Error GoTo 0
12       Exit Sub

CexRangeFP_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure CexRangeFP of Form frmBilling")
End Sub
Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub JobGrid_AfterColUpdate(ByVal ColIndex As Integer)
1         isUpdateCol = True
2         myChassis = Trim(text(17))
3         Select Case ColIndex
              Case 7
4                 JobGrid.Columns(9) = Round((JobGrid.Columns(7) / 100) * JobGrid.Columns(6))
                  ' BATASI DISCOUNT UTK DISC TTN
5                 If Not opWO(2).value Then
6                     With JobGrid
7                         If .Columns(ColIndex) <= gDiscMax And .Columns(ColIndex) >= 0 Then
8                             .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
9                         Else
10                            Me.Tag = ""
11                            If Log.validasiKabeng = 1 Then
12                                gBantuan = ""
13                                frmValidasi.Tag = "DISC"
14                                frmValidasi.Show vbModal
15                                Me.Tag = gBantuan
16                            End If
17                            If Me.Tag <> "validDisc" Then
                                  ''2015-10-06 || jika pogram dan dibebankan new ratna boleh melebihi ketentuan diskon
18                                If opWO(1).value = True And JobGrid.Columns(11) = False Then
                                  
                                  'If opWO(1).value = False And jobGrid.Columns(11) = False Then
19                                Else
20                                    If isLockDiscKacab = False Then
21                                        MsgBox "Discount tidak boleh lebih besar dari ketentuan / lebih kecil dari nol !!"
22                                        rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
23                                        If Not rsJob.EOF Then
24                                            rsJob("DISC") = 0
25                                            rsJob("DISCRP") = 0
26                                            .Columns(ColIndex) = 0
27                                            .Columns(ColIndex + 1) = 0
28                                            rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
29                                            .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
30                                        End If
31                                    End If
                                      
32                                End If
                                  
33                            Else
34                                rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
35                                If Not rsJob.EOF Then
36                                    rsJob("VALID") = gVerifikator
37                                End If
38                            End If
39                            rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
40                            .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
41                        End If
42                    End With
43                Else
44                    With JobGrid
45                        If .Columns(ColIndex) > 0 Then
46                            .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
47                        Else
48                            MsgBox "Discount tidak boleh lebih kecil dari nol !!"
49                            rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
50                            If Not rsJob.EOF Then
51                                rsJob("DISC") = 0
52                                rsJob("DISCRP") = 0
53                                .Columns(ColIndex) = 0
54                                .Columns(ColIndex + 1) = 0
55                                rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
56                                .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
57                            End If
58                        End If
59                    End With
60                End If
61                rsJob.MoveFirst
62                text(22) = 0
63                Do While Not rsJob.EOF
64                    text(22) = CDbl(text(22)) + CDbl(JobGrid.Columns(10))
65                    rsJob.MoveNext
66                Loop
67                text(29) = Format(Round((CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + CDbl(Text42.text)), 0), "#0,#")
68                txtTotalRp = Angka2Huruf(CDbl(text(29)), 1)
69            Case 8
70                With JobGrid
71                    If .Columns(ColIndex) Then
72                        If Log.discKacab = 1 Then
73                          .Splits(0).Columns(9).Locked = True
74                          .Splits(0).Columns(7).Locked = True
75                        Else
76                          .Splits(0).Columns(9).Locked = False
77                          .Splits(0).Columns(7).Locked = True
78                        End If
79                    Else
80                        If Log.discKacab = 1 Then
81                          .Splits(0).Columns(7).Locked = True
82                          .Splits(0).Columns(9).Locked = True
83                        Else
84                          .Splits(0).Columns(7).Locked = False
85                          .Splits(0).Columns(9).Locked = True
86                        End If


87                    End If
88                End With
89            Case 9
90                JobGrid.Columns(7) = Round((JobGrid.Columns(ColIndex) / CDbl(JobGrid.Columns(6))) * 100, 4)
                  ' BATASI DISCOUNT UTK DISC TTN
91                If Not opWO(2).value Then
92                    With JobGrid
93                        If .Columns(7) <= gDiscMax And .Columns(7) >= 0 Then
94                            .Columns(10) = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
95                        Else
96                            Me.Tag = ""
97                            If Log.validasiKabeng = 1 Then
98                                gBantuan = ""
99                                frmValidasi.Tag = "DISC"
100                               frmValidasi.Show vbModal
101                               Me.Tag = gBantuan
102                           End If
103                           If Me.Tag <> "validDisc" Then
104                               MsgBox "Discount tidak boleh lebih besar dari ketentuan / lebih kecil dari nol !!"
105                               rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
106                               If Not rsJob.EOF Then
107                                   rsJob("DISC") = 0
108                                   rsJob("DISCRP") = 0
109                                   .Columns(7) = 0
110                                   .Columns(ColIndex) = 0
111                                   rsJob("STLDISC") = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
112                                   .Columns(ColIndex + 2) = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
113                               End If
114                           Else
115                               rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
116                               If Not rsJob.EOF Then
117                                   rsJob("VALID") = gVerifikator
118                               End If
119                           End If
120                           .Columns(10) = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
121                       End If
122                   End With
123               Else
124                   With JobGrid
125                       If .Columns(7) > 0 Then
126                           .Columns(10) = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
127                       Else
128                           MsgBox "Discount tidak boleh lebih kecil dari nol !!"
129                           rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
130                           If Not rsJob.EOF Then
131                               rsJob("DISC") = 0
132                               rsJob("DISCRP") = 0
133                               .Columns(ColIndex) = 0
134                               .Columns(ColIndex + 1) = 0
135                               rsJob("STLDISC") = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
136                               .Columns(ColIndex + 2) = Round(((100 - CDbl(.Columns(7))) / 100) * CDbl(.Columns(6)))
137                           End If
138                       End If
139                   End With
140               End If
141               rsJob.MoveFirst
142               text(22) = 0
143               Do While Not rsJob.EOF
144                   text(22) = CDbl(text(22)) + CDbl(JobGrid.Columns(10))
145                   rsJob.MoveNext
146               Loop
147               text(29) = Format(Round((CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + CDbl(Text42.text)), 0), "#0,#")
148               txtTotalRp = Angka2Huruf(CDbl(text(29)), 1)
149           Case 11
                  Dim nilaiDisc As Double
150               With JobGrid
151                   If .Columns(ColIndex) = -1 Then
152                       rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
153                       If Not rsJob.EOF Then
154                           subJob = subJob - Round(CDbl(.Columns(6)))
155                           If statusSBI And (Trim(text(44)) = "" Or Trim(text(44)) = "DAN LAIN LAIN") Then
156                               .Columns(6) = hargaCustomerCBUNonTAM(CDbl(.Columns(1)), Trim(text(13)))
157                               rsJob("HRG") = Round(hargaCustomerCBUNonTAM(CDbl(.Columns(1)), Trim(text(13))))
158                           Else
159                               .Columns(6) = Round(objBilling.takeClassPrice(Trim(text(13)), 1) * CDbl(.Columns(1)))
160                               rsJob("HRG") = Round(objBilling.takeClassPrice(Trim(text(13)), 1) * CDbl(.Columns(1)))
161                           End If
162                           If Log.discKacab = 1 Then
                                  
163                               nilaiDisc = 0
164                               Set rsHlp = objO.FetchData("SELECT discJasaValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(0)) & "' and no_chassis = '" & Trim(text(17)) & "'")
165                               If Not rsHlp.EOF Then
166                                   isLockDiscKacab = True
                                      'If nDuoInv = False Then
                                        'Call giveDiscount(CInt(Nul(rsHlp!discJasaValid)), text(21), False)
167                                     nilaiDisc = CInt(Nul(rsHlp!discJasaValid))
168                                     rsJob("DISC") = CDbl(nilaiDisc)
169                                     .Columns(7) = CDbl(nilaiDisc)
170                                     JobGrid.Splits(0).Columns(7).Locked = True
171                                     JobGrid.Splits(0).Columns(8).Locked = True
172                                     JobGrid.Splits(0).Columns(9).Locked = True
                                      'End If
173                               End If

174                               rsJob("DISCRP") = (CDbl(rsJob("DISC")) / 100) * CDbl(rsJob("HRG"))
                                  '.Columns(6) = rsJob("HRG") - rsJob("DISCRP")
175                           Else
                                  'rsJob("DISC") = CDbl(text(21))
176                               rsJob("DISC") = 0
177                               rsJob("DISCRP") = 0
178                           End If


179                           rsJob("STLDISC") = Round(CDbl(.Columns(6)) - rsJob("DISCRP"))
180                           text(22) = SpyRound(CDbl(.Columns(6)))
                              '.Columns(7) = 0
181                       End If
182                       .Columns(7).Locked = False
183                       .Columns(8).Locked = False
184                   Else
185                       rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
186                       If Not rsJob.EOF Then
187                           If statusSBI And (Trim(text(44)) = "" Or Trim(text(44)) = "DAN LAIN LAIN") Then
188                               .Columns(6) = hargaDealerSBI(CDbl(.Columns(1)), Trim(text(13)), Trim(text(17)))
189                               rsJob("HRG") = Round(hargaDealerSBI(CDbl(.Columns(1)), Trim(text(13)), Trim(text(17))))
190                           Else
191                               .Columns(6) = Round(objBilling.takeClassPrice(Trim(text(13)), 2) * CDbl(.Columns(1)))
192                               rsJob("HRG") = Round(objBilling.takeClassPrice(Trim(text(13)), 2) * CDbl(.Columns(1)))
193                           End If
194                           rsJob("FREEC") = False
195                           If Log.discKacab = 1 Then
                                  
196                               nilaiDisc = 0
197                               Set rsHlp = objO.FetchData("SELECT discJasaValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(text(0)) & "' and no_chassis = '" & Trim(text(17)) & "'")
198                               If Not rsHlp.EOF Then
199                                   isLockDiscKacab = True
                                      'If nDuoInv = False Then
                                        'Call giveDiscount(CInt(Nul(rsHlp!discJasaValid)), text(21), False)
200                                     nilaiDisc = CInt(Nul(rsHlp!discJasaValid))
201                                     JobGrid.Splits(0).Columns(7).Locked = True
202                                     JobGrid.Splits(0).Columns(8).Locked = True
203                                     JobGrid.Splits(0).Columns(9).Locked = True
                                      'End If
204                               End If
205                               rsJob("DISC") = CDbl(nilaiDisc)
206                           Else
207                               rsJob("DISC") = CDbl(text(21))
208                           End If

209                           rsJob("DISCRP") = (CDbl(rsJob("DISC")) / 100) * CDbl(rsJob("HRG"))
210                           rsJob("STLDISC") = Round(((100 - CDbl(rsJob("DISC"))) / 100) * CDbl(rsJob("HRG")))
211                           subJob = subJob + (CDbl(rsJob("HRG")) * (CDbl(rsJob("DISC")) / 100))
212                           text(22) = text(22) - Round(CDbl(.Columns(6)))
213                           rsJob.Update
214                       End If
215                       .Columns(7).Locked = True
216                       .Columns(8).Locked = True
217                   End If
                      
218                   If isOnCarClass = False Then
219                       MsgBox "Kelas Kendaraan Tidak Terdaftar, Invoice di Cancel !", vbCritical + vbOKOnly, "Car Class"
220                       Call hapusForm
221                       Exit Sub
222                   End If
                                      
223                   rsJob.MoveFirst
224                   text(22) = 0
                      'Text(23) = 0
225                   Do While Not rsJob.EOF
226                       text(22) = CDbl(text(22)) + CDbl(.Columns(10))
227                       rsJob.MoveNext
228                   Loop
229                   text(29) = Format(Round((CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + CDbl(Text42.text)), 0), "#0,#")
230                   txtTotalRp = Angka2Huruf(CDbl(text(29)), 1)
231               End With
232       End Select
233       Set JobGrid.DataSource = rsJob
234       Call doHeadJob
235       isUpdateCol = False
End Sub

Private Sub JobGrid_GotFocus()
    VScroll1.value = 0
End Sub

Private Sub JobGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If simpan = False Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
                Call hapusForm
                text(0).SelLength = Len(Trim(text(0)))
                text(0).SetFocus
            End If
        End If
    End If
End Sub

Private Sub PartGrid_GotFocus()
    VScroll1.value = 1500
End Sub

Private Sub PartGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If simpan = False Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
                Call hapusForm
                text(0).SelLength = Len(text(0))
                text(0).SetFocus
            End If
        End If
    End If
End Sub

Private Sub subletGrid_GotFocus()
    VScroll1.value = 1900
End Sub

Private Sub subletGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If simpan = False Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
                Call hapusForm
                text(0).SelLength = Len(Trim(text(0)))
                text(0).SetFocus
            End If
        End If
    End If
End Sub

Private Sub TblKeluar_Click()
    onTimer = True
    Call hapusForm
    Unload Me
End Sub

Private Function cariTotal() As Double
      Dim rsTamp As ADODB.Recordset
      Dim hlp1 As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim jPPN As Double
      Dim sTipeCust As String
      Dim onlySublet As Boolean
      Dim total As Double
          
1        On Error GoTo cariTotal_Error

2         gKodePlg = IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
3         gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
4         gInvAlmt = IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
5         gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
6         gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)
          
7         total = 0
8         sSQL = "SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, NO_CHASSIS, NO_ENGINE, NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
               & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, " _
               & "REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, " _
               & "KLPKERJA1 , KLPKERJA2, KLPKERJA3, KLPKERJA4, KLPKERJA5, KLPKERJA6, KLPKERJA7, KLPKERJA8, " _
               & "Off1 , Off2, Off3, Off4, Off5, Off6, Off7, Off8, DISC1, DISC2, DISC3, DISC4, DISC5, DISC6, DISC7, DISC8, " _
               & "FREEC1, FREEC2, FREEC3, FREEC4, FREEC5, FREEC6, FREEC7, FREEC8 " _
               & "FROM T_INVOICE WITH (NOLOCK) WHERE NOINV = '" & Trim(text(20)) & IIf(nDuoInv, "-B", "") & "' AND NO_WO = '" & Trim(text(0)) & "'"
               
9         Set rsTamp = objO.FetchData(sSQL)
10        If Not rsTamp.EOF Then
11                onlySublet = True
12                total = total
13                Do While Not rsTamp.EOF
14                    For i = 1 To 8
15                        hlp1 = "REP" & i
16                        hlp2 = "RTE" & i
17                        hlp3 = "OFF" & i
18                        hlp4 = "DISC" & i
19                        hlp6 = "KLPKERJA" & i
20                        hlp7 = "FREEC" & i
                          
21                        If Trim(rsTamp(hlp1)) <> "" Then
22                            jPPN = (10 / 110) * CDbl(rsTamp(hlp3))
                              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
23                            total = total + Round((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp(hlp3)) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * (CDbl(rsTamp(hlp4)) / 100)))
24                            onlySublet = False
25                        End If
26                    Next
27                    rsTamp.MoveNext
28                Loop
29        End If
          
          ' ADD JOB
30        Set rsTamp = objO.FetchData("SELECT '" & Trim(gInvNo) & "' AS NOINV, INVDATE, '" & Trim(gKodePlg) & "' AS FS_KD_PLG, '" & Trim(gInvCust) & "' AS CUSTOMER, '" & Trim(gInvAlmt) & "' AS CUST_ADRES, '" & Trim(gInvKota) & "' AS KOTA, '" & Trim(gInvNPWP) & "' AS NPWP, NO_POLISI, FMDL, A.NO_CHASSIS, NO_ENGINE, A.NO_WO, FRONT, STATUS, MECH, FS_LOG_ADD, " _
                                   & "SBLT, PART_G+PART_NG AS PART, OIL_G+OIL_NG AS OLI, MTRL, MET, BODY, SLIP, FS_CATATAN, TOTAL, A.* FROM HIST_JOBADD A WITH (NOLOCK) INNER JOIN T_INVOICE B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL WHERE A.NO_WO = '" & Trim(text(0)) & "' AND A.NO_CHASSIS = '" & Trim(text(17)) & "' AND A.FS_LOG_DEL = ''")
31        If Not rsTamp.EOF Then
              'If createTemp(vNamaTabel) Then
32                Do While Not rsTamp.EOF
33                    jPPN = (10 / 110) * CDbl(rsTamp("OFFN"))
                      '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..
34                    total = total + Round((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) - ((CDbl(rsTamp("OFFN")) - IIf(sTipeCust = "AFFNONPPN", jPPN, 0)) * ((100 - IIf(Trim(rsTamp("FREEC")) = 1, 0, CDbl(rsTamp("DISC")))) / 100)))
35                    rsTamp.MoveNext
36                Loop
              'End If
37        End If
38        cariTotal = total

39       'On Error GoTo 0
40       Exit Function

cariTotal_Error:

41        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariTotal of Form frmBilling")
End Function

Private Function cariNPWP(noRgk As String) As String
1        On Error GoTo cariNPWP_Error

2         cariNPWP = objO.FetchData_String("SELECT NPWP FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRgk) & "'")

3        'On Error GoTo 0
4        Exit Function

cariNPWP_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariNPWP of Form frmBilling")
End Function

Private Sub simpanFakturGab(nomorNPWP As String, idx As Integer, kdLgn As String, namaFak As String, alamatFak1 As String, alamatFak2 As String, nilDPP As Double, noInvoic As String)
      Dim rsHelp As ADODB.Recordset
      Dim NoInvD As String
      Dim nStatus As String

1        On Error GoTo simpanFakturGab_Error

2         NoInvD = ""
3         If idx = 1 Then
4             Set rsHelp = objO.FetchDataACC("SELECT * FROM SIHDR WITH (NOLOCK) WHERE left(NoSo, 7) = '" & Trim(text(0)) & "'")
5             If Not rsHelp.EOF Then
6                 NoInvD = Trim(rsHelp("NOFAKTUR"))
7             End If
              
8             Set rsHelp = objO.FetchDataACC("SELECT * FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(kdLgn) & "'")
9             If Not rsHelp.EOF Then
10                nStatus = Trim(rsHelp("STATUSPERSH"))
11            End If
12        End If
          
13        objHT.InsertData_ACC ("INSERT INTO FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN (NOINVOICE,TGLINVOICE,NPWP,NPPKP,DPP,PPN,NAMAFP,ALAMAT1,ALAMAT2,KODELGN,MODUL,NOFPS,KODEENTRY,TGLENTRY,KET1,KET2,KET3,PEMERINTAH) " & _
                         "VALUES ('" & Trim(noInvoic) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & Trim(nomorNPWP) & "',''," & Round((nilDPP) * 100 / 110, 0) & "," & Round((nilDPP) * 10 / 110, 0) & ", " & _
                         "'" & Trim(namaFak) & "','" & Trim(alamatFak1) & "','" & Trim(alamatFak2) & "','" & Trim(kdLgn) & "','SVC','','" & gSA & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(text(16)) & "','" & Trim(NoInvD) & "','','" & IIf(nStatus = "P", 1, 0) & "')")

         'On Error GoTo 0
14       Exit Sub

simpanFakturGab_Error:
15        erTranAll = True
16        ErrMsg.judul = ""
17        ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanFakturGab of Form frmBilling")
18        ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanFakturGab of Form frmBilling")
End Sub

Private Sub simpanFaktur(nDuoFP As Boolean)
      Dim rsHelp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim rsFaktur As ADODB.Recordset
      Dim nFaktur As String
      Dim nFakturUP As String
      Dim boleh As Boolean
      Dim NoInvD As String
      Dim nCountFP As Double
      Dim noNPWP As String
1        On Error GoTo simpanFaktur_Error

2         nSQLUPDATECounterFP = ""
3         nSisa = 1
4         If Round((CDbl(text(29)) - CDbl(Text42.text)) * 10 / 110, 0) <= 0 Then Exit Sub
5         If OnNewFP = 0 Then
              'Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..PROFIL WITH (NOLOCK) WHERE TAHUN = '" & Right(Year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "'")
              'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
6             Set rsSearch = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..PROFIL WITH (NOLOCK) WHERE TAHUN = '" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "'")
7             ndFP = CDbl(rsSearch("NOFAKTURPAJAK"))
8             If Val(rsSearch("SISA")) > 1 Then nSisa = 2
9         Else
              'Set rsSearch = objO.FetchDataACC("SELECT *, cast(right(NoAkhir,8) as Integer) - cast(right(NoSekarang,8) as Integer) as SISA FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) and NoSekarang <> NoAkhir Order By Nomor ASC")
              'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
10            Set rsSearch = objHT.FetchDataACC("SELECT *, cast(right(NoAkhir,8) as Integer) - cast(right(NoSekarang,8) as Integer) as SISA FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) and NoSekarang <> NoAkhir Order By Nomor ASC")
              
11            ndFP = CDbl(Right(rsSearch("NOSEKARANG"), 8))
12            If Val(rsSearch("SISA")) > 1 Then nSisa = 2
13        End If
14        boleh = True
15        gInvCust = IIf(gInvCust = "", text(4), gInvCust)
16        gInvAlmt = IIf(gInvAlmt = "", txt5, gInvAlmt)
17        gInvKota = IIf(gInvKota = "", text(6) & "-" & text(7), gInvKota)
          
      '        Set rsHelp = objO.FetchDataACC("SELECT isNull(FPGABUNGAN,0) as FPGABUNGAN, NPWP FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & IIf(woStat = "SBI", gKodePlg, Trim(Text(2))) & "'")
              'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
18            Set rsHelp = objHT.FetchDataACC("SELECT isNull(FPGABUNGAN,0) as FPGABUNGAN, NPWP FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & IIf(woStat = "SBI", gKodePlg, Trim(text(2))) & "'")
19            If Not rsHelp.EOF Then
20                noNPWP = IIf(Trim(rsHelp("NPWP")) <> "", Trim(rsHelp("NPWP")), Trim(rsCustdat("NPWP")))
                  'If Trim(rsHelp("FPGABUNGAN")) = "1" Then  'TRUE'
21                If Trim(rsHelp("FPGABUNGAN")) = True Then
22                    nSisa = nSisa - 1
23                    boleh = False
                      ' PANGGIL MENU INSERT GABUNGAN
                      Dim kdPelanggan As String
                      Dim nmPelanggan As String
                      Dim alPelanggan As String
                      Dim ktPelanggan As String
24                    If Trim(woStat) = "SBI" Then
25                        kdPelanggan = gKodePlg
26                        nmPelanggan = gInvCust
27                        alPelanggan = gInvAlmt
28                        ktPelanggan = gInvKota
29                    Else
30                        kdPelanggan = text(2)
31                        nmPelanggan = text(4)
32                        alPelanggan = txt5
33                        ktPelanggan = text(6)
34                    End If
35                    If CDbl(text(29)) <> 0 Then Call simpanFakturGab(noNPWP, 1, IIf(woStat = "SBI", gKodePlg, Trim(text(2))), IIf(woStat = "SBI", gInvCust, Trim(text(4))), IIf(woStat = "SBI", gInvAlmt, Trim(txt5)), IIf(woStat = "SBI", gInvKota, Trim(text(6))), (CDbl(text(29)) - CDbl(Text42.text)), Trim(gKodeFaktur & Trim(text(20)) & IIf(nDuoInv, "-A", "")))
36                End If
37            End If
          
38        If Trim(text(40)) = "AFFNONPPN" Then
39            boleh = False
40        End If
          
41        If boleh Then
          Dim nLoop As Boolean
42        nLoop = False
cariNo:
                  'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
43                If OnNewFP = 0 Then
      '                Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..PROFIL WITH (NOLOCK) WHERE TAHUN = '" & Right(Year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "'")
44                    Set rsSearch = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..PROFIL WITH (NOLOCK) WHERE TAHUN = '" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "'")
45                Else
      '                Set rsSearch = objO.FetchDataACC("SELECT *, cast(right(NoAkhir,8) as Integer) - cast(right(NoSekarang,8) as Integer) as SISA FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) Order By Nomor ASC")
46                    Set rsSearch = objHT.FetchDataACC("SELECT *, cast(right(NoAkhir,8) as Integer) - cast(right(NoSekarang,8) as Integer) as SISA FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) Order By Nomor ASC")
47                End If
48                If Not rsSearch.EOF Then
49                    If OnNewFP = 0 Then
50                        If CDbl(text(29)) <> 0 Then
51                            nCountFP = CDbl(rsSearch("NOFAKTURPAJAK"))
52                            nCountFP = nCountFP + 1
53                            If nDuoFP Then ndFP = nCountFP + 1
54                        Else
55                            nCountFP = CDbl(rsSearch("NOFAKTURPAJAK"))
56                            nCountFP = nCountFP + 1
57                            If nDuoFP Then ndFP = nCountFP
58                            Exit Sub
59                        End If
60                    Else
61                        OnNomor = Val(rsSearch.Fields(0).value)
62                        If CDbl(text(29)) <> 0 Then
63                            nCountFP = CDbl(Right(rsSearch("NOSEKARANG"), 8))
64                            nCountFP = nCountFP + 1
65                            If nDuoFP Then ndFP = nCountFP + 1
66                        Else
67                            nCountFP = CDbl(Right(rsSearch("NOSEKARANG"), 8))
68                            nCountFP = nCountFP + 1
69                            If nDuoFP Then ndFP = nCountFP
70                            Exit Sub
71                        End If
72                    End If
73                Else
74                    If CDbl(text(29)) <> 0 Then
75                        nCountFP = 1
76                        If nDuoFP Then ndFP = nCountFP + 1
77                        Call objHT.InsertData_Faktur("INSERT INTO FAKTUR" & Right(gDb, 2) & "..PROFIL VALUES (" & Format(nCountFP, "00000000") & ",'" & Right(Format(isNowBilling, "MM/DD/YYYY"), 2) & "')")
78                    Else
79                        OnNomor = 0
80                        nCountFP = 1
81                        If nDuoFP Then ndFP = nCountFP
82                        Exit Sub
83                    End If
84                End If
85                If OnNewFP = 0 Then
                      'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
      '                Set rsFaktur = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
86                    Set rsFaktur = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
87                Else
      '                Set rsFaktur = objO.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
88                    Set rsFaktur = objHT.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) Order By Nomor ASC")
89                End If
90                If Not rsFaktur.EOF Then
91                    rsFaktur.MoveFirst
92                    If OnNewFP = 0 Then
93                        nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "." & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "."
94                        nFakturUP = nFaktur & Format(nCountFP, "#00000000")
95                    Else
96                        nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "-" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "."
97                        nFakturUP = nFaktur & Format(nCountFP, "#00000000")
98                    End If
99                End If
100               KodeCabangFaktur = Trim(rsFaktur("KODEDEALER"))
                  
                  Dim okFP As Boolean
101               okFP = False

102               Do Until okFP = True
                      'GANTI DENGAN OBJHT (BEGIN TRANS) - RUDI
      '                Set rsSearch = objO.FetchDataACC("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) WHERE NOFPS = '" & Trim(nFakturUP) & "'")
103                   Set rsSearch = objHT.FetchDataACC("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) WHERE NOFPS = '" & Trim(nFakturUP) & "'")
104                   If Not rsSearch.EOF Then
105                       nCountFP = nCountFP + 1
106                       If nDuoFP Then ndFP = nCountFP + 1
107                       nFakturUP = nFaktur & Format(nCountFP, "#00000000")
108                   Else
      '                    Set rsSearch = objO.FetchDataACC("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) WHERE NOFPS = '" & Trim(Replace(nFakturUP, "010.", "020.")) & "'")
109                       Set rsSearch = objHT.FetchDataACC("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) WHERE NOFPS = '" & Trim(Replace(nFakturUP, "010.", "020.")) & "'")
110                       If Not rsSearch.EOF Then
111                           nCountFP = nCountFP + 1
112                           If nDuoFP Then ndFP = nCountFP + 1
113                           nFakturUP = nFaktur & Format(nCountFP, "#00000000")
114                       Else
115                           okFP = True
116                           If nDuoFP Then
117                               If OnNewFP = 0 Then
118                                   nSQLUPDATECounterFP = "UPDATE FAKTUR" & Right(gDb, 2) & "..PROFIL SET NOFAKTURPAJAK = '" & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Right(Format(isNowBilling, "MM/DD/YYYY"), 2) & "'"
119                               Else
120                                   nSQLUPDATECounterFP = "UPDATE FAKTUR" & Right(gDb, 2) & "..FPKonter SET NOSEKARANG = '" & Trim(rsFaktur("KODEDEALER")) & "-" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "." & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and Nomor = " & OnNomor
121                               End If
122                           Else
123                               If OnNewFP = 0 Then
124                                   Call objHT.UpdateData_Faktur("UPDATE FAKTUR" & Right(gDb, 2) & "..PROFIL SET NOFAKTURPAJAK = '" & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Right(Format(isNowBilling, "MM/DD/YYYY"), 2) & "'")
125                               Else
126                                   Call objHT.UpdateData_Faktur("UPDATE FAKTUR" & Right(gDb, 2) & "..FPKonter SET NOSEKARANG = '" & Trim(rsFaktur("KODEDEALER")) & "-" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "." & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and Nomor = " & OnNomor)
127                               End If
128                           End If
129                       End If
130                   End If
131               Loop
132               nFaktur = nFakturUP

              'Dim NoInvD As String
133           NoInvD = ""
      '        Set rsSearch = objO.FetchData("SELECT NOFAKTUR FROM T_INVOICE WHERE NOINV LIKE '" & Trim(text(20)) & "%' AND NOINV NOT LIKE '%-B'")
      '        If Not rsSearch.EOF Then
      '            Do While Not rsSearch.EOF
134                   NoInvD = NoInvD & IIf(Trim(NoInvD) <> "", ",", "") & Trim(text(48)) 'Trim(rsSearch("NOFAKTUR"))
      '                rsSearch.MoveNext
      '            Loop
      '        End Ifn
135           Ket = "No Invoice : " & gKodeFaktur & Trim(text(20)) & IIf(nDuoInv, "-A", "") & IIf(Trim(NoInvD) <> "", " , " & Trim(NoInvD), "") & " , No Polisi : " & Trim(text(16)) '& " , CUSTOMER : " & Trim(text(4)) & ", DPP : " & setKoma(Round(CDbl(text(29))))
                  
              Dim vTotalNonPPn As Double
              Dim vTPPN As Double
136           vTotalNonPPn = Round(Koma2Titik(text(27)) * (100 / 110), 0) + Round((Koma2Titik(text(22)) * (100 / 110)), 0) + Round(((Koma2Titik(text(24)) + Koma2Titik(text(25))) * 100 / 110), 0)
137           vTPPN = (Round(Koma2Titik(text(22)) + Koma2Titik(text(28)) + Koma2Titik(text(27)) + Koma2Titik(text(24)) + Koma2Titik(text(25)), 0)) - vTotalNonPPn
                  
      '        Call objHT.InsertData_Bengkel("INSERT INTO FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR (NoFPS,TglFPS,NPWP,NPPKP,Keterangan,HargaJual,PPN,PPnBM,KodeUser,TglEntry,NamaFP,AlamatFP1,AlamatFP2,Modul,KODELGN) VALUES ('" & Trim(nFaktur) & "','" & Format(isNowBilling, "MM/DD/YYYY") & "','" & Trim(noNPWP) & "','','" & Trim(Ket) & "'," & Round((CDbl(Text(29)) - CDbl(Text(42))) * 100 / 110, 0) & "," & Round((CDbl(Text(29)) - CDbl(Text(42))) * 10 / 110, 0) & ",0,'" & Trim(gSA) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(gInvCust) & "','" & Left(Trim(gInvAlmt), 100) & "','" & Trim(gInvKota) & "','SVC','" & IIf(woStat = "SBI", gKodePlg, Trim(Text(2))) & "')")
138           Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR (NoFPS,TglFPS,NPWP,NPPKP,Keterangan,HargaJual,PPN,PPnBM,KodeUser,TglEntry,NamaFP,AlamatFP1,AlamatFP2,Modul,KODELGN) VALUES ('" & Trim(nFaktur) & "','" & Format(isNowBilling, "MM/DD/YYYY") & "','" & Trim(noNPWP) & "','','" & Trim(Ket) & "'," & vTotalNonPPn & "," & vTPPN & ",0,'" & Trim(gSA) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(gInvCust) & "','" & Left(Trim(gInvAlmt), 100) & "','" & Trim(gInvKota) & "','SVC','" & IIf(woStat = "SBI", gKodePlg, Trim(text(2))) & "')")
139           Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..INVOICE VALUES ('" & Trim(nFaktur) & "','" & gKodeFaktur & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "','" & Format(isNowBilling, "MM/DD/YYYY") & "')")
              
              '#########################################################'
              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
              '#########################################################'
140           Call objHT.UpdateData_Bengkel("UPDATE T_INVOICE SET NOFP = '" & Trim(nFaktur) & "' WHERE NOINV = '" & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "'")
141           Call objHT.UpdateData_ACC("UPDATE [" & Log.LogACC & "]..SIHDR SET NOFPS = '" & Trim(nFaktur) & "' WHERE NOFAKTUR IN ('" & NoInvD & "')")
142           nSQLo(2) = "UPDATE T_INVOICEnew SET NOFP = '" & Trim(nFaktur) & "' WHERE NOINV = '" & Trim(text(20)) & IIf(nDuoInv, "-A", "") & "' And CABANG = '" & Log.LogBengkel & "'"
              'MsgBox "No Faktur Pajak : " & Trim(nFaktur)
              
143           noFaktur1 = Trim(nFaktur)
144       End If

         'On Error GoTo 0
145      Exit Sub

simpanFaktur_Error:
146       erTranAll = True
147       ErrMsg.judul = ""
148       ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanFaktur of Form frmBilling")
149       ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanFaktur of Form frmBilling")
End Sub

Private Sub simpanFaktur2()
      Dim rsHelp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim rsSearch2 As ADODB.Recordset
      Dim rsFaktur As ADODB.Recordset
      Dim nFaktur As String
      Dim nFakturUP As String
      Dim boleh As Boolean
      Dim nCountFP As Double
      Dim noNPWP As String
      Dim nilaiPajak As Double


1        On Error GoTo simpanFaktur2_Error

2         boleh = True
3         gKodePlg = IIf(Trim(gKodePlg2) = "", text(2), gKodePlg2)
4         gInvCust = IIf(gInvCust2 = "", text(4), gInvCust2)
5         gInvAlmt = IIf(gInvAlmt2 = "", txt5, gInvAlmt2)
6         gInvKota = IIf(gInvKota2 = "", text(6) & "-" & text(7), gInvKota2) 'vNoTelp
7         gInvNPWP = IIf(gInvNPWP2 = "", "", gInvNPWP2) 'Text(14)

8         nilaiPajak = cariTotalManual  'cariTotal
9         If CDbl(nilaiPajak = 0) Then
10            boleh = False
11        Else
12                Set rsHelp = objHT.FetchDataACC("SELECT IsNull(FPGABUNGAN,0) as FPGABUNGAN, NPWP FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & gKodePlg & "'")
13                If Not rsHelp.EOF Then
14                    noNPWP = IIf(Trim(rsHelp("NPWP")) <> "", Trim(rsHelp("NPWP")), gInvNPWP)
                      'If Trim(rsHelp("FPGABUNGAN")) = "1" Then
15                    If Trim(rsHelp("FPGABUNGAN")) = True Then
16                        boleh = False
                          ' PANGGIL MENU INSERT GABUNGAN
17                        Call simpanFakturGab(noNPWP, 2, gKodePlg, Trim(gInvCust), Trim(gInvAlmt), Trim(gInvKota), nilaiPajak, Trim(gKodeFaktur & Trim(text(20)) & "-B"))
18                    End If
19                End If
20        End If
21        If Trim(text(40)) = "AFFNONPPN" Then
22            boleh = False
23        End If
          
24        If boleh = False Then
25            If nSQLUPDATECounterFP <> "" Then Call objHT.UpdateData_Faktur(nSQLUPDATECounterFP)
26        Else
          Dim nLoop As Boolean
27        nLoop = False
cariNo:
      '''            nCountFP = ndFP
      '''            If OnNewFP = 0 Then
      '''                Set rsFaktur = objO.FetchDataFP("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER")
      '''            Else
      '''                'Set rsFaktur = objO.FetchDataFP("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WHERE TAHUN = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
      '''
      '''            End If
      '''
      '''            'If Not rsFaktur.EOF Then
      '''                If OnNewFP = 0 Then
      '''                    'nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "." & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "."
      '''                    nFaktur = "010." & KodeCabangFaktur & "." & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "."
      '''                    nFakturUP = nFaktur & Format(nCountFP, "#00000000")
      '''                Else
      '''                    'nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "-" & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "."
      '''                    nFaktur = "010." & KodeCabangFaktur & "-" & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "."
      '''                    nFakturUP = nFaktur & Format(nCountFP, "#00000000")
      '''                End If
      '''            'End If
      '''
      '''            Dim okFP As Boolean
      '''            okFP = False
      '''
      '''            Do Until okFP = True
      '''                Set rsSearch = objo.FetchData("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WHERE NOFPS = '" & Trim(nFakturUP) & "'")
      '''                If Not rsSearch.EOF Then
      '''                    nCountFP = nCountFP + 1
      '''                    If nDuoFP Then ndFP = nCountFP + 1
      '''                    nFakturUP = nFaktur & Format(nCountFP, "#00000000")
      '''                Else
      '''                    Set rsSearch = objo.FetchData("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WHERE NOFPS = '" & Trim(Replace(nFakturUP, "010.", "020.")) & "'")
      '''                    If Not rsSearch.EOF Then
      '''                        nCountFP = nCountFP + 1
      '''                        If nDuoFP Then ndFP = nCountFP + 1
      '''                        nFakturUP = nFaktur & Format(nCountFP, "#00000000")
      '''                    Else
      '''                        okFP = True
      '''                        If OnNewFP = 0 Then
      '''                            Call objo.updateData("UPDATE FAKTUR" & Right(gDb, 2) & "..PROFIL SET NOFAKTURPAJAK = '" & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Right(Format(isnow, "MM/DD/YYYY"), 2) & "'")
      '''                        Else
      '''                            'Call objHT.UpdateData_Bengkel("UPDATE FAKTUR" & Right(gDb, 2) & "..FPKonter SET NOSEKARANG = '" & Trim(rsFaktur("KODEDEALER")) & "-" & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "." & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Format(isnow, "YYYY") & "' and Nomor = " & OnNomor)
      '''                            Call objo.updateData("UPDATE FAKTUR" & Right(gDb, 2) & "..FPKonter SET NOSEKARANG = '" & KodeCabangFaktur & "-" & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "." & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Format(isnow, "YYYY") & "' and Nomor = " & OnNomor)
      '''                        End If
      '''                    End If
      '''                End If
      '''            Loop
      '''            nFaktur = nFakturUP

28                If OnNewFP = 0 Then
29                    Set rsSearch = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..PROFIL WITH (NOLOCK) WHERE TAHUN = '" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "'")
30                Else
31                    Set rsSearch = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) and NoAkhir <> '" & Right(noFaktur1, 15) & "' Order By Nomor ASC")
32                End If

33                If Not rsSearch.EOF Then
34                    If OnNewFP = 0 Then
35                        If CDbl(text(29)) <> 0 Then
36                            nCountFP = CDbl(rsSearch("NOFAKTURPAJAK"))
37                            nCountFP = nCountFP + nSisa
38                        Else
39                            nCountFP = CDbl(rsSearch("NOFAKTURPAJAK"))
40                            nCountFP = nCountFP + nSisa
                              'Exit Sub
41                        End If
42                    Else
43                        OnNomor = Val(rsSearch.Fields(0).value)
44                        If CDbl(text(29)) <> 0 Then
45                            nCountFP = CDbl(Right(rsSearch("NOSEKARANG"), 8))
46                            nCountFP = nCountFP + nSisa
47                        Else
48                            nCountFP = CDbl(Right(rsSearch("NOSEKARANG"), 8))
49                            nCountFP = nCountFP + nSisa
                              'Exit Sub
50                        End If
51                    End If
52                Else
      '                If CDbl(text(29)) <> 0 Then
      '                    nCountFP = 1
      '                    objHT.InsertData_Faktur ("INSERT INTO FAKTUR" & Right(gDb, 2) & "..PROFIL VALUES (" & Format(nCountFP, "00000000") & ",'" & Right(Format(isNowBilling, "MM/DD/YYYY"), 2) & "')")
      '                Else
      '                    OnNomor = 0
      '                    nCountFP = 1
      '                    If nDuoFP Then ndFP = nCountFP
      '                    Exit Sub
      '                End If
53                End If
54                If OnNewFP = 0 Then
55                    Set rsFaktur = objHT.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH (NOLOCK)")
56                Else
57                    Set rsFaktur = objHT.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH (NOLOCK) WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and NoSekarang <> NoAkhir and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) and NoAkhir <> '" & Right(noFaktur1, 15) & "' Order By Nomor ASC")
58                End If
59                If Not rsFaktur.EOF Then
60                    rsFaktur.MoveFirst
61                    If OnNewFP = 0 Then
62                        nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "." & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "."
63                        nFakturUP = nFaktur & Format(nCountFP, "#00000000")
64                    Else
65                        nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "-" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "."
66                        nFakturUP = nFaktur & Format(nCountFP, "#00000000")
67                    End If
68                End If
69                KodeCabangFaktur = Trim(rsFaktur("KODEDEALER"))
                  
                  Dim okFP As Boolean
70                okFP = False


      '            Set rsSearch2 = objO.FetchData("SELECT NoAkhir FROM FAKTUR" & Right(gDb, 2) & "..FPKONTER WHERE NoAkhir = '" & Mid(Trim(nFakturUP), 5, 7) & Format(Val(Right(Trim(nFakturUP), 8)), "00000000") & "'")
      '            If rsSearch2.RecordCount > 0 Then
      '
      '                If OnNewFP <> 0 Then
      '                    Set rsSearch = objO.FetchData("SELECT *, cast(right(NoAkhir,8) as Integer) - cast(right(NoSekarang,8) as Integer) as SISA FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WHERE TAHUN = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir and cast(right(NoSekarang,8) as Integer) < cast(right(NoAkhir,8) as Integer) and cast(right(NoAkhir,8) as Integer) - cast(right(NoSekarang,8) as Integer) > 1 Order By Nomor ASC")
      '                    nFaktur = "010." & Trim(rsFaktur("KODEDEALER")) & "-" & Right(Year(Format(isnow, "MM/DD/YYYY")), 2) & "."
      '                    nFakturUP = nFaktur & Format(CDbl(Right(rsSearch("NOSEKARANG"), 8)), "#00000000")
      '                End If
      '            End If
                      
                      
                      
71                Do Until okFP = True
72                    Set rsSearch = objHT.FetchDataACC("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) WHERE NOFPS = '" & Trim(nFakturUP) & "'")
73                    If Not rsSearch.EOF Then
74                        nCountFP = nCountFP + 1
75                        nFakturUP = nFaktur & Format(nCountFP, "#00000000")
76                    Else
77                        Set rsSearch = objHT.FetchDataACC("SELECT NoFPS FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) WHERE NOFPS = '" & Trim(Replace(nFakturUP, "010.", "020.")) & "'")
78                        If Not rsSearch.EOF Then
79                            nCountFP = nCountFP + 1
80                            nFakturUP = nFaktur & Format(nCountFP, "#00000000")
81                        Else
82                            okFP = True
                              'REVISI 14-01-2015 RUDI
83                            If nSQLUPDATECounterFP <> "" Then Call objHT.UpdateData_Faktur(nSQLUPDATECounterFP)  'objO.UpdateData(nSQLUPDATECounterFP)
84                            If OnNewFP = 0 Then
85                                Call objHT.UpdateData_Faktur("UPDATE FAKTUR" & Right(gDb, 2) & "..PROFIL SET NOFAKTURPAJAK = '" & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Right(Format(isNowBilling, "MM/DD/YYYY"), 2) & "'")
86                            Else
87                                Call objHT.UpdateData_Faktur("UPDATE FAKTUR" & Right(gDb, 2) & "..FPKonter SET NOSEKARANG = '" & Trim(rsFaktur("KODEDEALER")) & "-" & Right(year(Format(isNowBilling, "MM/DD/YYYY")), 2) & "." & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Format(isNowBilling, "YYYY") & "' and Nomor = " & OnNomor)
88                            End If
89                        End If
90                    End If
91                Loop
92                nFaktur = nFakturUP




93            Ket = "No Invoice : " & gKodeFaktur & Trim(text(20)) & "-B , No Polisi : " & Trim(text(16))
              
94            Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR (NoFPS,TglFPS,NPWP,NPPKP,Keterangan,HargaJual,PPN,PPnBM,KodeUser,TglEntry,NamaFP,AlamatFP1,AlamatFP2,Modul,KODELGN) VALUES ('" & Trim(nFaktur) & "','" & Format(isNowBilling, "MM/DD/YYYY") & "','" & Trim(gInvNPWP) & "','','" & Trim(Ket) & "'," & Round(CDbl(nilaiPajak) * 100 / 110, 0) & "," & Round(CDbl(nilaiPajak) * 10 / 110, 0) & ",0,'" & Trim(gSA) & "','" & Format(isNowBilling, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(gInvCust) & "','" & Trim(gInvAlmt) & "','" & Trim(gInvKota) & "','SVC','" & gKodePlg & "')")
95            Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..INVOICE VALUES ('" & Trim(nFaktur) & "','" & gKodeFaktur & Trim(text(20)) & "-B" & "','" & Format(isNowBilling, "MM/DD/YYYY") & "')")
              '#########################################################'
              '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
              '#########################################################'
96            Call objHT.UpdateData_Bengkel("UPDATE T_INVOICE SET NOFP2 = '" & Trim(nFaktur) & "' WHERE NOINV = '" & Trim(text(20)) & "-B'")
97            nSQLo(3) = "UPDATE T_INVOICEnew SET NOFP2 = '" & Trim(nFaktur) & "' WHERE NOINV = '" & Trim(text(20)) & "-B'"
              'MsgBox "No Faktur Pajak : " & Trim(nFaktur)
98            noFaktur2 = Trim(nFaktur)
99        End If
100       gKodePlg = ""
101       gInvCust = ""
102       gInvAlmt = ""
103       gInvKota = ""
104       gInvNPWP = ""

         'On Error GoTo 0
105      Exit Sub

simpanFaktur2_Error:
106       erTranAll = True
107       ErrMsg.judul = ""
108       ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanFaktur2 of Form frmBilling")
109       ErrMsg.Show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanFaktur2 of Form frmBilling")
End Sub

Private Sub simpanHO(noWO As String)
      Dim rsHlp As ADODB.Recordset
      Dim PRC As Boolean
          
1         On Error GoTo vbError
          ' SIMPAN HISTORY
2         ClearsSql
3         gsSql1(0) = "NO_WO"
4         gsSql2(0) = Trim(text(0))
5         gsSql1(1) = "CUSTOMER"
6         gsSql2(1) = Trim(text(4))
7         gsSql1(2) = "CUST_ADRES"
8         gsSql2(2) = Trim(txt5)
9         gsSql1(3) = "KOTA"
10        gsSql2(3) = Trim(text(6))
11        gsSql1(4) = "NO_POLISI"
12        gsSql2(4) = Trim(text(16))
13        gsSql1(5) = "NO_ENGINE"
14        gsSql2(5) = Trim(text(18))
15        gsSql1(6) = "NO_CHASSIS"
16        gsSql2(6) = Trim(text(17))
17        gsSql1(7) = "FS_KD_DEALER"
18        gsSql2(7) = Trim(gKdCbg)
19        gsSql1(8) = "FS_NM_DEALER"
20        gsSql2(8) = Trim(gNmCbg)
21        gsSql1(9) = "FRONT"
22        gsSql2(9) = Trim(text(19))
23        gsSql1(10) = "MECH"
24        gsSql2(10) = Trim(text(38))
25        gsSql1(11) = "KM"
26        gsSql2(11) = CDbl(text(37))
27        gsSql1(12) = "DISC"
28        gsSql2(12) = CDbl(text(21))
29        gsSql1(13) = "DISCT"
30        gsSql2(13) = CDbl(text(23))
          
31        gsSql1(14) = "PART_G"
32        gsSql2(14) = Val(subPartGenuine * (1 - (CDbl(text(23)) / 100))) 'CDbl(Text(24))     'subPart    '
33        gsSql1(15) = "PART_NG"
34        gsSql2(15) = Val(subPartNGenuine * (1 - (CDbl(text(23)) / 100)))
35        gsSql1(16) = "MTRL"
36        gsSql2(16) = Val(CDbl(text(28)))
37        gsSql1(17) = "OIL_G"
38        gsSql2(17) = Val(subOliGenuine * (1 - (CDbl(text(23)) / 100))) 'CDbl(Text(25))     'subOli     '
39        gsSql1(18) = "OIL_NG"
40        gsSql2(18) = Val(subOliNGenuine * (1 - (CDbl(text(23)) / 100)))
41        gsSql1(19) = "MET"
42        gsSql2(19) = Val(CDbl(Text42.text))
43        gsSql1(20) = "SBLT"
44        gsSql2(20) = Val(CDbl(text(27)))     'subSublet  '
45        gsSql1(21) = "LABOR"
46        gsSql2(21) = IIf(lblPDS.Visible, Val(subJob), IIf(nDuoInv, Val(subJob), Val(CDbl(text(22)))))
47        gsSql1(22) = "BODY"
48        gsSql2(22) = Val(CDbl(text(39)))     'subBody    '
49        gsSql1(23) = "TOTAL"
50        gsSql2(23) = Val((subPartGenuine * (1 - (CDbl(text(23)) / 100))) + (subPartNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(text(28)) + (subOliGenuine * (1 - (CDbl(text(23)) / 100))) + (subOliNGenuine * (1 - (CDbl(text(23)) / 100))) + CDbl(Text42.text) + CDbl(text(27)) + IIf(lblPDS.Visible, subJob, CDbl(text(22))) + CDbl(text(39)))
              
51        Set rsHlp = objO.FetchData("SELECT * FROM jobdta WITH (NOLOCK) where no_wo = '" & Trim(text(0)) & "' AND no_chassis = '" & text(17) & "'")
52        If Not rsHlp.EOF Then
53            gsSql1(24) = "[DATE]"
54            gsSql2(24) = Format(rsHlp("DATE"), "MM/DD/YYYY")
55            PRC = IIf(rsHlp("FS_RC") <> "", True, False)
56            gsSql1(25) = "TIME_WO"
57            gsSql2(25) = Trim(rsHlp("TIME"))
58            gsSql1(26) = "TIME_M"
59            gsSql2(26) = Trim(rsHlp("TIME_M"))
60            gsSql1(27) = "TIME_S"
61            gsSql2(27) = Trim(rsHlp("TIME_S"))
62            gsSql1(28) = "TIME_A"
63            gsSql2(28) = Trim(rsHlp("TIME_A"))
64            gsSql1(29) = "TIME_AT"
65            gsSql2(29) = Trim(rsHlp("TIME_AT"))
66            gsSql1(30) = "TIME_SA"
67            gsSql2(30) = Trim(rsHlp("TIME_SA"))

              Dim hlp As String
              Dim hlp2 As String
              Dim hlp3 As String
              Dim hlp4 As String
              
68            For i = 1 To 7
69                hlp = "REP" & i
70                hlp2 = "MEKANIK" & i
71                hlp3 = "KLPKERJA" & i
72                hlp4 = "RTE" & i
                  
73                gsSql1(30 + i) = hlp
74                gsSql2(30 + i) = Trim(rsHlp(hlp))
75                gsSql1(37 + i) = hlp2
76                gsSql2(37 + i) = Trim(rsHlp(hlp2))
77                gsSql1(44 + i) = hlp3
78                gsSql2(44 + i) = Trim(rsHlp(hlp3))
79                gsSql1(51 + i) = hlp4
80                gsSql2(51 + i) = Trim(Replace(rsHlp(hlp4), ",", "."))
81            Next
              
82            gsSql1(59) = "PAID"
83            gsSql2(59) = Trim(rsHlp("PAID"))
84            gsSql1(60) = "STATUS"
85            gsSql2(60) = Trim(rsHlp("STATUS"))
86            gsSql1(61) = "NOFAK"
87            gsSql2(61) = Trim(rsHlp("NOFAK"))
88            gsSql1(62) = "REMARK1"
89            gsSql2(62) = Trim(rsHlp("REMARK1"))
90            gsSql1(63) = "FN_RETURN"
91            gsSql2(63) = Trim(rsHlp("FN_RETURN"))
92            gsSql1(64) = "FS_RC"
93            gsSql2(64) = Trim(Nul(rsHlp("FS_RC")))
94            gsSql1(65) = "SBE"
95            gsSql2(65) = Trim(Nul(rsHlp("SBE")))
96            gsSql1(66) = "TIPEWO"
97            gsSql2(66) = Trim(rsHlp("TPSVC"))
98            gsSql1(67) = "KODETIPE"
99            gsSql2(67) = Trim(rsHlp("KODETIPE"))
100       End If
          
101       gsSql1(68) = "NOINV"
102       gsSql2(68) = Trim(text(20))
          
          Dim memoPart As String
103       memoPart = ""
104       Set rsHlp = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' and FS_LOG_DEL = ''")
105       If Not rsHlp.EOF Then
106           Do While Not rsHlp.EOF
107               memoPart = memoPart & (rsHlp("NO_WO") & "^" & rsHlp("KODEITEM") & "^" & rsHlp("KETERANGAN") & "^" & rsHlp("QTY") & "^" & rsHlp("HARGA") & "^" & rsHlp("OLI") & "^" & rsHlp("NOINV") & "^" & rsHlp("FN_EXT_DISKON")) & "|"
108               rsHlp.MoveNext
109           Loop
110       End If
111       gsSql1(69) = "MEMOPART"
112       gsSql2(69) = Trim(memoPart)
          
          Dim memoSublet As String
113       memoSublet = ""
114       Set rsHlp = objO.FetchData("SELECT * FROM HIST_SUBLET WITH (NOLOCK) WHERE NO_WO = '" & noWO & "'")
115       If Not rsHlp.EOF Then
116           Do While Not rsHlp.EOF
117               memoSublet = memoSublet & (rsHlp("NOPO") & "^" & rsHlp("NO_WO") & "^" & rsHlp("NOINV") & "^" & rsHlp("NO_CHASSIS") & "^" & rsHlp("NOFAK") & "^" & rsHlp("KODEITEM") & "^" & rsHlp("NAMABARANG") & "^" & rsHlp("KODESATUAN") & "^" & rsHlp("QTY") & "^" & rsHlp("HARGA") & "^" & rsHlp("NONOTA") & "^" & rsHlp("KETERANGAN") & "^" & rsHlp("FLAG") & "^" & rsHlp("SUPPLIER") & "^" & rsHlp("TAGIH") & "|")
118               rsHlp.MoveNext
119           Loop
120       End If
121       gsSql1(70) = "MEMOSUBLET"
122       gsSql2(70) = Trim(memoSublet)
          
          Dim memoAddJob As String
123       memoAddJob = ""
124       Set rsHlp = objO.FetchData("SELECT * FROM HIST_JOBADD WITH (NOLOCK) WHERE NO_WO = '" & noWO & "'")
125       If Not rsHlp.EOF Then
126           Do While Not rsHlp.EOF
127               memoAddJob = memoAddJob & (rsHlp("NO_WO") & "^" & rsHlp("NO_CHASSIS") & "^" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "^" & rsHlp("KLPKERJA") & "^" & rsHlp("REP") & "^" & CDbl(rsHlp("RTE")) & "^" & rsHlp("ACTRTE") & "^" & rsHlp("MEKANIK") & "^" & rsHlp("GRP") & "^" & rsHlp("OFFN") & "^" & rsHlp("DISC") & "^" & rsHlp("TIME") & "|")
128               rsHlp.MoveNext
129           Loop
130       End If
131       gsSql1(71) = "MEMOJOBADD"
132       gsSql2(71) = Trim(memoAddJob)
          
133       gsSql1(72) = "MOU"
134       gsSql2(72) = Trim(text(34))
135       gsSql1(73) = "FS_KD_PLG"
136       gsSql2(73) = Trim(text(2))
137       gsSql1(74) = "FS_KD_WIL"
138       gsSql2(74) = Trim(text(3))
139       gsSql1(75) = "THS"
140       gsSql2(75) = Trim(IIf(nTHS, 1, 0))
141       gsSql1(76) = "PDS"
142       gsSql2(76) = Trim(IIf(lblPDS.Visible, 1, 0))
143       gsSql1(77) = "INVDATE"
144       gsSql2(77) = Format(text(1), "MM/DD/YYYY")
          
145       If Get_ConnectionHO(HODb, HOUser, HOPass, HOKomp) Then
146           erTranHO = False
147           Call objHT.OpenTransactionHO
              
              '// INPUT HISTORY
148           Set rsTemp = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH (NOLOCK) WHERE no_wo = '" & noWO & "' AND no_chassis = '" & text(17) & "' and noinv = '" & Trim(text(20)) & "'")   'noinv = '" & Text(20) & "' AND
149           If Not rsTemp.EOF Then
150               PSQL = GensSql2("HISTORY", True) + "WHERE no_chassis = '" & text(17) & "' AND no_WO = '" & noWO & "' and noinv = '" & Trim(text(20)) & "'" 'AND noInv = '" & Text(20) & "'"
151           Else
152               PSQL = GensSql2("HISTORY", False)
153           End If
154           Call objHT.UpdateData_Ho(PSQL)
                      
              '// INPUT M_ESTIMASI
155           If PRC Then
156               Set rsHlp = objO.FetchData("SELECT * FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & text(17) & "' AND NOWO = '" & noWO & "' ORDER BY SBENEXT")
157               If Not rsHlp.EOF Then
158                   Do While Not rsHlp.EOF
      '                    Set rsTemp = objO.FetchDataHO("SELECT * FROM M_ESTIMASI WHERE NORANGKA = '" & text(17) & "' AND SBENEXT = '" & rsHlp("SBENEXT") & "'")
159                       Call objO.FetchDataHO("DELETE M_ESTIMASI WHERE NORANGKA = '" & Trim(text(17)) & "' AND SBENEXT = '" & Trim(rsHlp("SBENEXT")) & "'")
      '                    If Not rsTemp.EOF Then
      '                        Call objHT.UpdateData_Ho("UPDATE M_ESTIMASI SET RATAKM = '" & rsHlp("RATAKM") & "', JT_TGL = '" & Format(rsHlp("JT_TGL"), "MM/DD/YYYY") & "', JT_BLN = '" & Format(rsHlp("JT_BLN"), "MM/DD/YYYY") & "', KMNEXT = '" & rsHlp("KMNEXT") & "', " _
      '                                       & "KMNOW = '" & rsHlp("KMNOW") & "', TGLSVCNOW = '" & Format(rsHlp("TGLSVCNOW"), "MM/DD/YYYY") & "', NOINV = '" & rsHlp("NOINV") & "', NOWO = '" & rsHlp("NOWO") & "', NAMACUST = '" & rsHlp("NAMACUST") & "', " _
      '                                       & "ALAMATCUST = '" & rsHlp("ALAMATCUST") & "', NOPOLISI = '" & rsHlp("NOPOLISI") & "', PSTN = '" & rsHlp("PSTN") & "', GSM = '" & rsHlp("GSM") & "', CDMA = '" & rsHlp("CDMA") & "' " _
      '                                       & "WHERE NORANGKA = '" & text(17) & "' AND SBENEXT = '" & rsHlp("SBENEXT") & "'")
      '                    Else
160                           Call objHT.InsertData_Ho("INSERT INTO M_ESTIMASI VALUES ('" & rsHlp("NORANGKA") & "', '" & rsHlp("RATAKM") & "', '" & Format(rsHlp("JT_TGL"), "MM/DD/YYYY") & "', '" & Format(rsHlp("JT_BLN"), "MM/DD/YYYY") & "', '" & rsHlp("KMNEXT") & "', '" & rsHlp("SBENEXT") & "', " _
                                             & "'" & rsHlp("KMNOW") & "', '" & Format(rsHlp("TGLSVCNOW"), "MM/DD/YYYY") & "', '" & rsHlp("NOINV") & "', '" & rsHlp("NOWO") & "', '" & rsHlp("NAMACUST") & "', '" & rsHlp("ALAMATCUST") & "', '" & rsHlp("NOPOLISI") & "', " _
                                             & "'" & rsHlp("PSTN") & "', '" & rsHlp("GSM") & "', '" & rsHlp("CDMA") & "','','')")
      '                    End If
161                       rsHlp.MoveNext
162                   Loop
163               End If
164           End If
              
              '// INPUT CUSTDAT
165           Set rsTemp = objO.FetchDataHO("SELECT NO_CHASSIS FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & text(17) & "'")
166           If Not rsTemp.EOF Then
167               Call objHT.DeleteData_Ho("DELETE CUSTDAT WHERE NO_CHASSIS = '" & text(17) & "'")
168           End If
              
169           Set rsTemp = objO.FetchData("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & text(17) & "'")
170           If rsTemp.EOF Then
171               Call objHT.InsertData_Ho("INSERT INTO " & HODb & "..CUSTDAT (CUSTOMER, CUST_ADRES, KOTA, KODEPOS, TELPON, NO_POLISI, NO_ENGINE, WMI, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, " _
                                 & "FS_KD_DEALER, FS_NM_DEALER, DDATE, REMARK, CODE, NPWP, PPN, DISC, KEYNO, [USER], SPK, TGL_REP, JOB, JOBCD, FS_KM_AKHIR, USEJOBDSC, RCDATE, KD_K, CR, " _
                                 & "CDK, TGL_D, CDD, CRDTCD, FS_TH_MOBIL, USEFS_NM_HOBI, FS_KD_PLG, FS_KD_WIL, DOTAM, USENAME, USEADDR, USECITY, USEPOS, USEPSTN, USEGSM, USECDMA, USET4LAHIR, " _
                                 & "USETGLLAHIR, ASURANSI, ASR_ADDR, ASR_CITY, ASR_TGLM, ASR_TGLS , ASR_LIMIT, ASR_NOTE, NOTE, GSM, CDMA, USEAGAMA, USERADD, FLAGKM, HEAVYUSE, PKP,KDPLG_ASR) VALUES (" _
                                 & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("CUST_ADRES") & "', '" & rsTemp("KOTA") & "', '" & rsTemp("KODEPOS") & "', '" & rsTemp("TELPON") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("NO_ENGINE") & "', '" & rsTemp("WMI") & "', '" & rsTemp("NO_CHASSIS") & "', " _
                                 & "'" & rsTemp("MDL") & "', '" & rsTemp("model") & "', '" & rsTemp("FMDL") & "', '" & rsTemp("Color") & "', '" & rsTemp("FS_KD_DEALER") & "', '" & rsTemp("FS_NM_DEALER") & "', '" & Format(rsTemp("ddate"), "MM/DD/YYYY") & "', '" & rsTemp("REMARK") & "', '" & rsTemp("CODE") & "', '" & rsTemp("npwp") & "', " _
                                 & "" & rsTemp("PPN") & ", " & rsTemp("disc") & ", '" & rsTemp("KEYNO") & "', '" & rsTemp("USER") & "', '" & rsTemp("SPK") & "', '" & Format(rsTemp("TGL_REP"), "MM/DD/YYYY") & "', '" & rsTemp("JOB") & "', '" & rsTemp("JOBCD") & "', '" & rsTemp("FS_KM_AKHIR") & "', '" & rsTemp("USEJOBDSC") & "', " _
                                 & "'" & Format(rsTemp("RCDATE"), "MM/DD/YYYY") & "', '" & rsTemp("KD_K") & "', '" & rsTemp("CR") & "', '" & rsTemp("CDK") & "', '" & Format(rsTemp("TGL_D"), "MM/DD/YYYY") & "', '" & rsTemp("CDD") & "', '" & rsTemp("CRDTCD") & "', '" & rsTemp("FS_TH_MOBIL") & "', '" & rsTemp("USEFS_NM_HOBI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                 & "'" & Format(rsTemp("DOTAM"), "MM/DD/YYYY") & "', '" & rsTemp("USENAME") & "', '" & rsTemp("USEADDR") & "', '" & rsTemp("USECITY") & "', '" & rsTemp("USEPOS") & "', '" & rsTemp("USEPSTN") & "', '" & rsTemp("USEGSM") & "', '" & rsTemp("USECDMA") & "', '" & rsTemp("USET4LAHIR") & "', " _
                                 & "'" & Format(rsTemp("USETGLLAHIR"), "MM/DD/YYYY") & "', '" & rsTemp("ASURANSI") & "', '" & rsTemp("ASR_ADDR") & "', '" & rsTemp("ASR_CITY") & "', '" & rsTemp("ASR_TGLM") & "', '" & rsTemp("ASR_TGLS") & "' , '" & rsTemp("ASR_LIMIT") & "', '" & rsTemp("ASR_NOTE") & "', '" & rsTemp("NOTE") & "', '" & rsTemp("GSM") & "', '" & rsTemp("CDMA") & "', " _
                                 & "'" & rsTemp("USEAGAMA") & "', '" & rsTemp("USERADD") & "', '" & rsTemp("FLAGKM") & "', '" & rsTemp("HEAVYUSE") & "','" & Trim(rsTemp("PKP")) & "','" & Trim(rsTemp("KDPLG_ASR")) & "')")
172           End If
173           Call simpanHistPemilik(Trim(text(17)))
              
174           If Trim(text(34)) <> "" And tipeByr = "T" Then
175               Call objHT.InsertData_Ho("INSERT INTO " & HODb & "..MOUTRN VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & text(34) & "', '" & noWO & "', '" & text(20) & "', '" & text(17) & "', '" & text(16) & "', '" & gKdCbg & "', " & CDbl(text(29)) & ", 'U','','')")
176               Call objHT.UpdateData_Ho("UPDATE MOUHDR SET HUTANG = HUTANG + " & CDbl(text(29)) & " WHERE NO_MOU = '" & text(34) & "'")
                  
                  '// INPUT TEMPKIRIM
177               If Trim(text(34)) <> "" Then
178                   Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('MOU','" & text(34) & "','S')")
179               End If
180           End If
              
181           If erTranHO = False Then
182               Call objHT.CloseTransactionHO
183           Else
184               Call objHT.CloseTransactionHO_err
185           End If
186       Else
187           Set rsTemp = objO.FetchData("SELECT * FROM TAMPHISTORY WITH (NOLOCK) WHERE noinv = '" & text(20) & "' AND no_wo = '" & nWO & "' AND no_chassis = '" & text(17) & "'")
188           If Not rsTemp.EOF Then
189               PSQL = GensSql2("TAMPHISTORY", True) + "WHERE no_chassis = '" & text(17) & "' AND no_WO = '" & nWO & "' AND noInv = '" & text(20) & "'"
190           Else
191               PSQL = GensSql2("TAMPHISTORY", False)
192           End If
193           Call objO.UpdateData(PSQL)
194       End If
195       Exit Sub
          
vbError:
196       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanHO of Form frmBilling")

197       If erTranHO = False Then
198           Call objHT.CloseTransactionHO
199       Else
200           Call objHT.CloseTransactionHO_err
201       End If
End Sub

Private Sub simpanHistPemilik(noRangka As String)
      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset

1        On Error GoTo simpanHistPemilik_Error

2         Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & noRangka & "'")
3         If Not rsHlp.EOF Then
4             Set rsHlp = objO.FetchDataHO("SELECT * FROM " & HODb & "..HIST_PEMILIK WITH (NOLOCK) WHERE NO_CHASSIS = '" & noRangka & "' and FS_KD_PLG = '" & Trim(rsHlp("FS_KD_PLG")) & "' AND CUSTOMER = '" & Trim(rsHlp("CUSTOMER")) & "'")
5             If rsHlp2.EOF Then
6                 Call objHT.InsertData_Ho("INSERT INTO " & HODb & "..HIST_PEMILIK VALUES ('" & Trim(rsHlp("FS_KD_PLG")) & "', '" & Trim(rsHlp("FS_KD_WIL")) & "', '" & Trim(rsHlp("CUSTOMER")) & "', '" & Trim(rsHlp("CUST_ADRES")) & "', " _
                                 & "'" & Trim(rsHlp("KOTA")) & "','" & Trim(rsHlp("KODEPOS")) & "', '" & Trim(rsHlp("TELPON")) & "', '" & Trim(rsHlp("GSM")) & "', '" & Trim(rsHlp("CDMA")) & "', '" & Trim(rsHlp("USEAGAMA")) & "', " _
                                 & "'" & Trim(rsHlp("USEJOBDSC")) & "','" & Trim(rsHlp("USEFS_NM_HOBI")) & "', '" & Trim(rsHlp("NO_CHASSIS")) & "', '" & Trim(rsHlp("NO_POLISI")) & "', '" & Trim(rsHlp("NO_ENGINE")) & "', '" & Trim(rsHlp("MDL")) & "', " _
                                 & "'" & Trim(rsHlp("MODEL")) & "', '" & Trim(rsHlp("FMDL")) & "', '" & Trim(rsHlp("COLOR")) & "', '" & Trim(rsHlp("KEYNO")) & "', '" & Trim(rsHlp("FS_KM_AKHIR")) & "', '" & Trim(rsHlp("FS_TH_MOBIL")) & "', " _
                                 & "'" & Trim(rsHlp("FS_KD_DEALER")) & "', '" & Trim(rsHlp("FS_NM_DEALER")) & "', '" & Trim(rsHlp("DDATE")) & "','', '" & Format(rsHlp("DOTAM"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("NPWP")) & "', '" & Trim(rsHlp("PPN")) & "', " _
                                 & "'" & Trim(rsHlp("DISC")) & "', '" & Trim(rsHlp("USERADD")) & "', '" & Trim(rsHlp("USENAME")) & "', '" & Trim(rsHlp("USEADDR")) & "', '" & Trim(rsHlp("USECITY")) & "', '" & Trim(rsHlp("USEPOS")) & "', " _
                                 & "'" & Trim(rsHlp("USEPSTN")) & "', '" & Trim(rsHlp("USEGSM")) & "', '" & Trim(rsHlp("USECDMA")) & "','" & Trim(rsHlp("USET4LAHIR")) & "','" & Format(rsHlp("USETGLLAHIR"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("ASURANSI")) & "', '" & Trim(rsHlp("ASR_ADDR")) & "', " _
                                 & "'" & Trim(rsHlp("ASR_CITY")) & "', '" & Format(rsHlp("ASR_TGLM"), "MM/DD/YYYY") & "', '" & Format(rsHlp("ASR_TGLS"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("ASR_LIMIT")) & "', '" & Trim(rsHlp("ASR_NOTE")) & "', '" & Trim(rsHlp("NOTE")) & "','','', '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', '" & Trim(rsHlp("PKP")) & "','" & Trim(rsHlp("KDPLG_ASR")) & "')")
7             Else
8                 Call objHT.UpdateData_Ho("UPDATE " & HODb & "..HIST_PEMILIK SET FS_KD_WIL = '" & Trim(rsHlp("FS_KD_WIL")) & "', CUST_ADRES = '" & Trim(rsHlp("CUST_ADRES")) & "', KOTA = '" & Trim(rsHlp("KOTA")) & "',KODEPOS = '" & Trim(rsHlp("KODEPOS")) & "', " _
                                 & "TELPON = '" & Trim(rsHlp("TELPON")) & "',GSM = '" & Trim(rsHlp("GSM")) & "',CDMA = '" & Trim(rsHlp("CDMA")) & "',USEAGAMA = '" & Trim(rsHlp("USEAGAMA")) & "', " _
                                 & "USEJOBDSC = '" & Trim(rsHlp("USEJOBDSC")) & "',USEFS_NM_HOBI = '" & Trim(rsHlp("USEFS_NM_HOBI")) & "',NO_POLISI = '" & Trim(rsHlp("NO_POLISI")) & "',NO_ENGINE = '" & Trim(rsHlp("NO_ENGINE")) & "',MDL = '" & Trim(rsHlp("MDL")) & "', " _
                                 & "MODEL = '" & Trim(rsHlp("MODEL")) & "',FMDL = '" & Trim(rsHlp("FMDL")) & "',COLOR = '" & Trim(rsHlp("COLOR")) & "',KEYNO = '" & Trim(rsHlp("KEYNO")) & "',FS_KM_AKHIR = '" & CDbl(rsHlp("FS_KM_AKHIR")) & "',FS_TH_MOBIL = '" & Trim(rsHlp("FS_TH_MOBIL")) & "', " _
                                 & "FS_KD_DEALER = '" & Trim(rsHlp("FS_KD_DEALER")) & "',FS_NM_DEALER = '" & Trim(rsHlp("FS_NM_DEALER")) & "',DDATE = '" & Trim(rsHlp("DDATE")) & "','',DOTAM = '" & Format(rsHlp("DOTAM"), "MM/DD/YYYY") & "',NPWP = '" & Trim(rsHlp("NPWP")) & "',PPN = '" & Trim(rsHlp("PPN")) & "', " _
                                 & "DISC = '" & CDbl(rsHlp("DISC")) & "',USERADD = '" & Trim(rsHlp("USERADD")) & "',USENAME = '" & Trim(rsHlp("USENAME")) & "',USEADDR = '" & Trim(rsHlp("USEADDR")) & "',USECITY = '" & Trim(rsHlp("USECITY")) & "',USEPOS = '" & Trim(rsHlp("USEPOS")) & "', " _
                                 & "USEPSTN = '" & Trim(rsHlp("USEPSTN")) & "',USEGSM = '" & Trim(rsHlp("USEGSM")) & "',USECDMA = '" & Trim(rsHlp("USECDMA")) & "',USET4LAHIR = '" & Trim(rsHlp("USET4LAHIR")) & "',USETGLLAHIR = '" & Format(rsHlp("USETGLLAHIR"), "MM/DD/YYYY") & "',ASURANSI = '" & Trim(rsHlp("ASURANSI")) & "', " _
                                 & "ASR_ADDR = '" & Trim(rsHlp("ASR_ADDR")) & "',ASR_CITY = '" & Trim(rsHlp("ASR_CITY")) & "',ASR_TGLM = '" & Format(rsHlp("ASR_TGLM"), "MM/DD/YYYY") & "',ASR_TGLS = '" & Format(rsHlp("ASR_TGLS"), "MM/DD/YYYY") & "',ASR_LIMIT = '" & Trim(rsHlp("ASR_LIMIT")) & "',ASR_NOTE = '" & Trim(rsHlp("ASR_NOTE")) & "', " _
                                 & "NOTE = '" & Trim(rsHlp("NOTE")) & "', FS_LOG_DEL = '" & Trim(rsHlp("FS_LOG_DEL")) & "', FS_TGL_DEL = '" & Trim(rsHlp("FS_TGL_DEL")) & "', TGLENTRY = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "',PKP = '" & Trim(rsHlp("PKP")) & "',KDPLG_ASR = '" & Trim(rsHlp("KDPLG_ASR")) & "') " _
                                 & "WHERE NO_CHASSIS = '" & noRangka & "' and FS_KD_PLG = '" & Trim(rsHlp("FS_KD_PLG")) & "' AND CUSTOMER = '" & Trim(rsHlp("CUSTOMER")) & "'")
9             End If
10        End If

11       'On Error GoTo 0
12       Exit Sub

simpanHistPemilik_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanHistPemilik of Form frmBilling")
End Sub

Private Sub Text_Change(Index As Integer)
    If Index = 27 Then
        Select Case Index
            Case 27
                If Text42.text <> "" Then
                    text(29) = Format(Round(CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + IIf(nDuoInv, 0, CDbl(Text42.text)), 0), "#0,#")
                    txtTotalRp = Angka2Huruf(CDbl(text(29)), 1)
                End If
        End Select
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If

    Select Case Index
        Case 31
            VScroll1.value = 2500
        Case 32
            'VScroll1.value = 2700
        Case 28
            VScroll1.value = 2000
    End Select
        
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    Select Case KeyCode
        Case vbKeyF12
            Select Case Index
                Case 0
                    gMainKeyCari = text(Index)
                    'gSearchOption = "ORDER"
                    gSearchOption = "INV"
                    gBantuan = "AND RPR_COND <> 'K' AND NOINV = ''"
                    frmCari.Show (vbModal)
'                Case 42
'                    gMainKeyCari = text(Index)
'                    gSearchOption = "METERAI"
'                    frmCari.Show (vbModal)
            End Select
        Case vbKeyEscape
            If simpan = False Then
                If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
                    Call hapusForm
                    text(0).SelLength = Len(text(0))
                    text(0).SetFocus
                End If
            End If
    End Select
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii <> 13 Then
        Select Case Index
            Case 28
                KeyAscii = CekNumerik(KeyAscii)
                simpan = False
        End Select
    End If
End Sub

Private Function masihAdaLimitMOU(TipeMdl As String, noChassis As String, nilaiNominal As Double) As Boolean
      Dim noMou As String
      Dim rsSearch As ADODB.Recordset
          
1        On Error GoTo masihAdaLimitMOU_Error

2         masihAdaLimitMOU = False
3         noMou = Trim(cekMOU(Trim(TipeMdl), Trim(noChassis)))
4         If noMou <> "" Then
             ' If Log.LogHOConn = 0 Then
              'tidak online HO
              '    masihAdaLimitMOU = False
             ' Else
5                 Set rsSearch = objO.FetchData("SELECT LIMITAR-HUTANG+BAYAR AS SISA FROM MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & noMou & "'")
6                 If Not rsSearch.EOF Then
7                     If CDbl(rsSearch("SISA")) - CDbl(nilaiNominal) > 0 Then
8                         masihAdaLimitMOU = True
9                     End If
10                End If
             ' End If
              
11        End If

12       'On Error GoTo 0
13       Exit Function

masihAdaLimitMOU_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure masihAdaLimitMOU of Form frmBilling")
End Function

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset
          
1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 0
3                 If text(Index) <> "" Then
4                     boleh = True
5                     If simpan = False Then
6                         If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
7                             boleh = False
8                         End If
9                     End If
10                    If boleh Then
11                        n = Trim(text(0))
12                        Call hapusForm
13                        text(0) = n
14                        Set rsSearch = objO.FetchData("SELECT A.*, ISNULL(B.NO_WO,'notValid') AS VALID FROM jobDta A WITH (NOLOCK) LEFT JOIN VALIDORDER B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND USERDEL = '' WHERE A.no_WO = '" & text(Index) & "'")
15                        If Not rsSearch.EOF Then
16                            If Trim(rsSearch("VALID")) = "notValid" Then
17                                MsgBox ("Order belum di validasi SA !!!!")
18                                Exit Sub
19                            ElseIf (rsSearch("DELIV") = "PRC" Or rsSearch("DELIV") = "OKE") And rsSearch("RPR_COND") <> "K" And rsSearch("NOINV") = "" Then
20                                Call panggilWO(text(Index))
21                                limitMOU = False
22                                If Not opWO(2) Then
23                                    If Trim(cekMOU(Trim(text(12)), Trim(text(17)))) <> "" Then
24                                        If Not masihAdaLimitMOU(Trim(text(12)), Trim(text(17)), CDbl(text(22))) Then
25                                            limitMOU = True
26                                        End If
27                                    End If
28                                End If
29                                text(21).SelLength = Len(text(21))
30                                TblSimpan.Enabled = True
31                            ElseIf rsSearch("NOINV") <> "" Then
32                                MsgBox ("Telah dibuat invoice, Tanggal " & TglDDMMMMYYYY(Trim(rsSearch("INVDATE"))) & " No. Invoice = " & Trim(rsSearch("NOINV")))
33                                Exit Sub
34                            ElseIf rsSearch("DELIV") = "" And rsSearch("RPR_COND") <> "S" Then
                                  Dim rsHlp As ADODB.Recordset
35                                Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(rsSearch("NO_POLISI")) & " |" & Trim(rsSearch("NO_WO")) & "%'")
36                                If Not rsHlp.EOF Then
37                                    MsgBox ("Unit belum selesai dikerjakan !! " & vbCr & "Dikerjakan terakhir tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL"))) & " di " & Trim(rsHlp("NOSTALL")))
38                                End If
39                                Exit Sub
40                            Else
41                                gMainKeyCari = text(Index)
42                                gSearchOption = "INV"
43                                gBantuan = "AND RPR_COND <> 'K' AND NOINV = ''"
44                                frmCari.Show (vbModal)
45                            End If
46                        End If
47                        simpan = False
48                    End If
49                End If
50            Case 21, 23, 26
51                If text(Index) = "" Then
52                    text(Index) = 0
53                ElseIf text(Index) > gDiscMax Then '100 Then
54                    MsgBox "Discount melebihi batas !!"
55                    text(Index) = 0
56                    text(Index).SelLength = Len(text(Index))
57                    text(Index).SetFocus
58                    Exit Sub
59                End If
                  
60                Select Case Index
                      Case 21
61                        text(22) = subJob
62                        text(39) = subBody
63                        text(22) = CDbl(text(22)) * ((100 - CDbl(text(Index))) / 100)
64                        text(39) = CDbl(text(39)) * ((100 - CDbl(text(Index))) / 100)
65                    Case 23
66                        text(24) = subPart
67                        text(25) = subOli
68                        text(24) = CDbl(text(24)) * ((100 - CDbl(text(Index))) / 100)
69                        text(25) = CDbl(text(25)) * ((100 - CDbl(text(Index))) / 100)
70                    Case 26
71                        text(27) = subSublet
72                        text(27) = CDbl(text(27)) * ((100 - CDbl(text(Index))) / 100)
73                End Select
74                text(29) = Format(Round((CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + CDbl(Text42.text)), 0), "#0,#")
75                txtTotalRp = Angka2Huruf(CDbl(text(29)), 1)
      '        Case 42
      '            If text(Index) = "" Then
      '                text(Index) = 0
      '            Else
      '                Set rsSearch = objO.FetchData("SELECT FN_BATAS FROM T_METERAI WHERE FN_BATAS = '" & text(Index) & "'")
      '                If rsSearch.EOF Then
      '                    gMainKeyCari = text(Index)
      '                    gSearchOption = "METERAI"
      '                    frmCari.Show (vbModal)
      '                Else
      '                    text(29) = Format(Round((CDbl(text(22)) + CDbl(text(24)) + CDbl(text(25)) + CDbl(text(27)) + CDbl(text(28)) + CDbl(text(39)) + CDbl(text42.text)), 0), "#0,#")
      '                    txtTotalRp = Angka2Huruf(text(29), 1)
      '                End If
      '            End If
76        End Select

77       'On Error GoTo 0
78       Exit Sub

text_Validate_Error:

79        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmBilling")
End Sub

Private Sub TimerFiskal_Timer()
    TimerFiskal.Enabled = False
    If isFP_OK = False Then
        MsgBox "Anda tidak punya Nomor Faktur" & Chr(13) & "Hubungi Administrasi (ADH) - FP" & Chr(13) & "Invoice tidak Bisa Dibuka", vbCritical + vbOKOnly
        Call hapusForm
        Unload Me
    End If
    If isNoFP_OK = False Then
         MsgBox "Sisa Nomor Faktur : " & isNoFP_New & Chr(13) & "Segera Hubungi Administrasi (ADH) - FP", vbInformation + vbOKOnly
    End If
    If isNoFP_New < 2 Then
        MsgBox "Anda tidak punya Nomor Faktur" & Chr(13) & "Hubungi Administrasi (ADH) - FP" & Chr(13) & "Invoice tidak Bisa Dibuka", vbCritical + vbOKOnly
        Call hapusForm
        Unload Me
    End If
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

Sub updateBilling()
1        On Error GoTo updateBilling_Error

2         Call panggilSublet(Trim(text(0)), Trim(text(17)))

3        'On Error GoTo 0
4        Exit Sub

updateBilling_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure updateBilling of Form frmBilling")
End Sub

Private Sub panggilSublet(noOrder As String, noRangka As String)
          
1        On Error GoTo panggilSublet_Error

2         adaSublet = False
3         cmdSublet.Visible = False
4         text(27) = 0
          
5         Set rsSublet = objO.FetchData("SELECT NONOTA, KODEITEM, NAMABARANG, KODESATUAN, QTY, VALID, PRICE as HARGA,SUPPLIER FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noOrder & "' AND no_CHASSIS = '" & noRangka & "' ")
6         If Not rsSublet.EOF Then
7             Do While Not rsSublet.EOF
8                 text(27) = text(27) + (CDbl(rsSublet("HARGA")) * CDbl(rsSublet("QTY")))
9                 rsSublet.MoveNext
10            Loop
      '        Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
      '                     & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & noOrder & "' AND no_chassis = '" & noRangka & "' ) and no_wo = '" & noOrder & "' AND no_chassis = '" & noRangka & "' " _
      '                     & "ORDER BY nonota, kodeitem")
11            subSublet = CDbl(text(27))
12            adaSublet = True
13            cmdSublet.Visible = True
14            Set subletGrid.DataSource = rsSublet
15            Call doHeadSublet
16        End If

         'On Error GoTo 0
17       Exit Sub

panggilSublet_Error:

18        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilSublet of Form frmBilling")
End Sub

Private Function namaPembawa(noOrder As String) As String
1        On Error GoTo namaPembawa_Error

2         namaPembawa = objO.FetchData_String("SELECT ANCUST FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(noOrder) & "'")

3        'On Error GoTo 0
4        Exit Function

namaPembawa_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure namaPembawa of Form frmBilling")
End Function

Private Function cariTotalManual() As Double

    On Error GoTo vbError
    cariTotalManual = 0
    If rsJob.State <> adStateClosed Then
        rsJob.MoveFirst
        Do While Not rsJob.EOF
            cariTotalManual = cariTotalManual + IIf(rsJob("FreeC"), 0, (rsJob("HRG") - IIf(rsJob("DISC") <> 100, (rsJob("HRG") - rsJob("DISCRP")), 0)))
            rsJob.MoveNext
        Loop
    End If
    Exit Function
    
vbError:
    cariTotalManual = 0
End Function




