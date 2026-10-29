VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOrderDMSNonDSS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Order Pekerjaan"
   ClientHeight    =   9345
   ClientLeft      =   2100
   ClientTop       =   1830
   ClientWidth     =   11400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Frame FrAR 
      BackColor       =   &H000000FF&
      Caption         =   ". . . . . . . . . .::  CONTROL  AR  ::. . . . . . . . ."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1560
      TabIndex        =   287
      Top             =   1320
      Visible         =   0   'False
      Width           =   4650
      Begin VB.Timer TimerAR 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox PicARO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   800
         Left            =   270
         ScaleHeight     =   765
         ScaleWidth      =   4065
         TabIndex        =   288
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
      End
      Begin ucXPButton.XPButton XPButton1 
         Height          =   330
         Left            =   3720
         TabIndex        =   289
         Top             =   1110
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BTYPE           =   3
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmOrderDMSNonDSS.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   360
         ScaleHeight     =   735
         ScaleWidth      =   3930
         TabIndex        =   290
         Top             =   280
         Width           =   3935
         Begin TDBNumber6Ctl.TDBNumber txtRp 
            Height          =   315
            Index           =   200
            Left            =   1080
            TabIndex        =   291
            Top             =   30
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calculator      =   "frmOrderDMSNonDSS.frx":001C
            Caption         =   "frmOrderDMSNonDSS.frx":003C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmOrderDMSNonDSS.frx":00A8
            Keys            =   "frmOrderDMSNonDSS.frx":00C6
            Spin            =   "frmOrderDMSNonDSS.frx":0110
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
            Index           =   201
            Left            =   1080
            TabIndex        =   292
            Top             =   360
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calculator      =   "frmOrderDMSNonDSS.frx":0138
            Caption         =   "frmOrderDMSNonDSS.frx":0158
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmOrderDMSNonDSS.frx":01C4
            Keys            =   "frmOrderDMSNonDSS.frx":01E2
            Spin            =   "frmOrderDMSNonDSS.frx":022C
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
            Caption         =   "Current"
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
            Index           =   37
            Left            =   120
            TabIndex        =   295
            Top             =   60
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Batas AR"
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
            Index           =   38
            Left            =   120
            TabIndex        =   294
            Top             =   390
            Width           =   855
         End
         Begin VB.Label nAR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   525
            Left            =   2640
            TabIndex        =   293
            Top             =   120
            Width           =   1020
         End
      End
      Begin VB.Line Line28 
         X1              =   360
         X2              =   3600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblOto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OTORITAS  KACAB"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   765
         TabIndex        =   296
         Top             =   1110
         Width           =   2475
      End
      Begin VB.Image Image1 
         Height          =   1560
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   210
      End
      Begin VB.Image Image1 
         Height          =   1560
         Index           =   1
         Left            =   4440
         Top             =   0
         Width           =   210
      End
      Begin VB.Line Line34 
         BorderWidth     =   2
         X1              =   360
         X2              =   3600
         Y1              =   1130
         Y2              =   1130
      End
   End
   Begin VB.Frame frmEvent 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   14400
      TabIndex        =   266
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label sgnEvent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PESAN PART"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   267
         Top             =   240
         Width           =   1980
      End
   End
   Begin VB.Frame frmeGantiNoPol 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ooo000ooo - Ganti No Polisi Salah - ooo000ooo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   15360
      TabIndex        =   256
      Top             =   2640
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdBatal 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Ba&tal"
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
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   264
         Top             =   2160
         Width           =   1725
      End
      Begin VB.CommandButton cmdGanti 
         BackColor       =   &H00DCFCFB&
         Caption         =   "G&anti"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   262
         Top             =   2160
         Width           =   1725
      End
      Begin MSMask.MaskEdBox text 
         Height          =   525
         Index           =   110
         Left            =   1650
         TabIndex        =   257
         Top             =   675
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   525
         Index           =   111
         Left            =   1650
         TabIndex        =   259
         Top             =   1320
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "??"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   525
         Index           =   112
         Left            =   2280
         TabIndex        =   260
         Top             =   1320
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   525
         Index           =   113
         Left            =   3360
         TabIndex        =   261
         Top             =   1320
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label83 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi Sebelumnya"
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
         Left            =   570
         TabIndex        =   263
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi Sebelumnya"
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
         Left            =   570
         TabIndex        =   258
         Top             =   720
         Width           =   960
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
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   15000
      TabIndex        =   210
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
   Begin VB.Frame KabPass 
      BackColor       =   &H00C0FFFF&
      Height          =   1815
      Left            =   15000
      TabIndex        =   200
      Top             =   8280
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdBtlPass 
         Caption         =   "&Batal"
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
         Left            =   2040
         TabIndex        =   206
         Top             =   1440
         Width           =   1050
      End
      Begin VB.CommandButton tblOkPass 
         Caption         =   "&Ok"
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
         Left            =   480
         TabIndex        =   205
         Top             =   1440
         Width           =   1050
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   83
         Left            =   1200
         TabIndex        =   201
         Top             =   240
         Width           =   720
         _ExtentX        =   1270
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   84
         Left            =   1200
         TabIndex        =   203
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   85
         Left            =   1200
         TabIndex        =   207
         Top             =   960
         Width           =   2040
         _ExtentX        =   3598
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
         PromptChar      =   ">"
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   208
         Top             =   1005
         Width           =   690
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama "
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
         Left            =   615
         TabIndex        =   204
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Petugas"
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
         TabIndex        =   202
         Top             =   285
         Width           =   585
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   480
      TabIndex        =   78
      Top             =   3120
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "`"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(1)=   "Text1(4)"
      Tab(0).Control(2)=   "Text1(3)"
      Tab(0).Control(3)=   "Text1(2)"
      Tab(0).Control(4)=   "Text1(1)"
      Tab(0).Control(5)=   "Text1(0)"
      Tab(0).Control(6)=   "Frame11"
      Tab(0).Control(7)=   "Frame9"
      Tab(0).Control(8)=   "text(23)"
      Tab(0).Control(9)=   "text(24)"
      Tab(0).Control(10)=   "text(25)"
      Tab(0).Control(11)=   "text(26)"
      Tab(0).Control(12)=   "text(27)"
      Tab(0).Control(13)=   "text(74)"
      Tab(0).Control(14)=   "text(91)"
      Tab(0).Control(15)=   "text(92)"
      Tab(0).Control(16)=   "text(93)"
      Tab(0).Control(17)=   "text(94)"
      Tab(0).Control(18)=   "text(95)"
      Tab(0).Control(19)=   "text(97)"
      Tab(0).Control(20)=   "text(98)"
      Tab(0).Control(21)=   "text(99)"
      Tab(0).Control(22)=   "text(100)"
      Tab(0).Control(23)=   "text(101)"
      Tab(0).Control(24)=   "JualNRM"
      Tab(0).Control(25)=   "Image6"
      Tab(0).Control(26)=   "Image5"
      Tab(0).Control(27)=   "Shape1"
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "`"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(1)=   "Image1(0)"
      Tab(1).Control(2)=   "imgTanda(0)"
      Tab(1).Control(3)=   "imgTanda(1)"
      Tab(1).Control(4)=   "ImgCrop"
      Tab(1).Control(5)=   "imgTanda(2)"
      Tab(1).Control(6)=   "imgTanda(3)"
      Tab(1).Control(7)=   "imgTanda(4)"
      Tab(1).Control(8)=   "imgTanda(5)"
      Tab(1).Control(9)=   "imgTanda(6)"
      Tab(1).Control(10)=   "imgTanda(7)"
      Tab(1).Control(11)=   "imgTanda(8)"
      Tab(1).Control(12)=   "imgTanda(9)"
      Tab(1).Control(13)=   "imgTanda(10)"
      Tab(1).Control(14)=   "imgTanda(11)"
      Tab(1).Control(15)=   "imgTanda(12)"
      Tab(1).Control(16)=   "imgTanda(13)"
      Tab(1).Control(17)=   "imgTanda(14)"
      Tab(1).Control(18)=   "imgTanda(15)"
      Tab(1).Control(19)=   "imgTanda(16)"
      Tab(1).Control(20)=   "imgTanda(17)"
      Tab(1).Control(21)=   "imgTanda(18)"
      Tab(1).Control(22)=   "imgTanda(19)"
      Tab(1).Control(23)=   "imgTanda(20)"
      Tab(1).Control(24)=   "imgTanda(21)"
      Tab(1).Control(25)=   "imgTanda(22)"
      Tab(1).Control(26)=   "imgTanda(23)"
      Tab(1).Control(27)=   "imgTanda(24)"
      Tab(1).Control(28)=   "imgTanda(25)"
      Tab(1).Control(29)=   "imgTanda(26)"
      Tab(1).Control(30)=   "imgTanda(27)"
      Tab(1).Control(31)=   "imgTanda(28)"
      Tab(1).Control(32)=   "imgTanda(29)"
      Tab(1).Control(33)=   "imgTanda(30)"
      Tab(1).Control(34)=   "imgTanda(31)"
      Tab(1).Control(35)=   "imgTanda(32)"
      Tab(1).Control(36)=   "imgTanda(33)"
      Tab(1).Control(37)=   "imgTanda(34)"
      Tab(1).Control(38)=   "imgTanda(35)"
      Tab(1).Control(39)=   "imgTanda(36)"
      Tab(1).Control(40)=   "imgTanda(37)"
      Tab(1).Control(41)=   "imgTanda(38)"
      Tab(1).Control(42)=   "imgTanda(39)"
      Tab(1).Control(43)=   "imgTanda(40)"
      Tab(1).Control(44)=   "imgTanda(41)"
      Tab(1).Control(45)=   "imgTanda(42)"
      Tab(1).Control(46)=   "imgTanda(43)"
      Tab(1).Control(47)=   "imgTanda(44)"
      Tab(1).Control(48)=   "imgTanda(45)"
      Tab(1).Control(49)=   "imgTanda(46)"
      Tab(1).Control(50)=   "imgTanda(47)"
      Tab(1).Control(51)=   "imgTanda(48)"
      Tab(1).Control(52)=   "imgTanda(49)"
      Tab(1).Control(53)=   "imgTanda(50)"
      Tab(1).Control(54)=   "imgTanda(51)"
      Tab(1).Control(55)=   "imgTanda(52)"
      Tab(1).Control(56)=   "imgTanda(53)"
      Tab(1).Control(57)=   "Image9"
      Tab(1).Control(58)=   "Image10"
      Tab(1).Control(59)=   "text(38)"
      Tab(1).Control(60)=   "text(37)"
      Tab(1).Control(61)=   "text(36)"
      Tab(1).Control(62)=   "text(35)"
      Tab(1).Control(63)=   "text(34)"
      Tab(1).Control(64)=   "Frame6"
      Tab(1).Control(65)=   "Frame15"
      Tab(1).ControlCount=   66
      TabCaption(2)   =   "`"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape3"
      Tab(2).Control(1)=   "sgnPart"
      Tab(2).Control(2)=   "Label27"
      Tab(2).Control(3)=   "LbSupply"
      Tab(2).Control(4)=   "sgnStall"
      Tab(2).Control(5)=   "Frame10"
      Tab(2).Control(6)=   "FrameFill"
      Tab(2).Control(7)=   "Frame5"
      Tab(2).Control(8)=   "text(51)"
      Tab(2).Control(9)=   "Frame12"
      Tab(2).Control(10)=   "cmdSublet"
      Tab(2).Control(11)=   "cmdAddJob"
      Tab(2).Control(12)=   "cmdWaktu"
      Tab(2).Control(13)=   "frmeWaktu"
      Tab(2).Control(14)=   "cmdStall"
      Tab(2).Control(15)=   "tblCetakPO"
      Tab(2).Control(16)=   "FrSupply"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "`"
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Shape4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label10"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label52"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label45"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label36"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label35"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label34"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label63"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label64"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label67"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label16"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label53"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "text(87)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Txt(87)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Txt(53)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Frame4"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Frame14"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Frame2"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "text(114)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "text(109)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "MaskEdBox16"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "text(108)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "Frame13"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "MaskEdBox15"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "text(107)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "text(63)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "text(62)"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "text(105)"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "MaskEdBox14"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "text(103)"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "text(102)"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "text(90)"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "text(89)"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "text(88)"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "text(86)"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "text(75)"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "text(55)"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "text(54)"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "text(53)"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "text(52)"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Frame8"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "cmdFU"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Frame3"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "cboPelanggan"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "cmdGateS"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "chUse"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Txt(54)"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "Txt(55)"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "cmdLengkap"
      Tab(3).Control(48).Enabled=   0   'False
      Tab(3).Control(49)=   "frmLengkap"
      Tab(3).Control(49).Enabled=   0   'False
      Tab(3).Control(50)=   "framePembatalanWashing"
      Tab(3).Control(50).Enabled=   0   'False
      Tab(3).ControlCount=   51
      Begin VB.Frame framePembatalanWashing 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pembatalan Washing"
         Height          =   2295
         Left            =   3360
         TabIndex        =   323
         Top             =   840
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
            TabIndex        =   325
            Top             =   1800
            Width           =   1725
         End
         Begin VB.TextBox txtAlasanBtlWashing 
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   324
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame FrSupply 
         Height          =   3015
         Left            =   -75000
         TabIndex        =   319
         Top             =   2000
         Visible         =   0   'False
         Width           =   8895
         Begin TrueOleDBGrid70.TDBGrid PartGrid1 
            Height          =   3015
            Left            =   0
            TabIndex        =   321
            Top             =   120
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5318
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
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
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
            Caption         =   "PART ORDER - PART SUPPLY"
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
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(66)  =   "Named:id=33:Normal"
            _StyleDefs(67)  =   ":id=33,.parent=0"
            _StyleDefs(68)  =   "Named:id=34:Heading"
            _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   ":id=34,.wraptext=-1"
            _StyleDefs(71)  =   "Named:id=35:Footing"
            _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(73)  =   "Named:id=36:Selected"
            _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(75)  =   "Named:id=37:Caption"
            _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(77)  =   "Named:id=38:HighlightRow"
            _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(79)  =   "Named:id=39:EvenRow"
            _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(81)  =   "Named:id=40:OddRow"
            _StyleDefs(82)  =   ":id=40,.parent=33"
            _StyleDefs(83)  =   "Named:id=41:RecordSelector"
            _StyleDefs(84)  =   ":id=41,.parent=34"
            _StyleDefs(85)  =   "Named:id=42:FilterBar"
            _StyleDefs(86)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -70320
         ScaleHeight     =   1335
         ScaleWidth      =   3210
         TabIndex        =   298
         Top             =   480
         Width           =   3210
         Begin TDBNumber6Ctl.TDBNumber txtRp 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   299
            Top             =   30
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   556
            Calculator      =   "frmOrderDMSNonDSS.frx":0254
            Caption         =   "frmOrderDMSNonDSS.frx":0274
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmOrderDMSNonDSS.frx":02E0
            Keys            =   "frmOrderDMSNonDSS.frx":02FE
            Spin            =   "frmOrderDMSNonDSS.frx":0348
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1397882885
            MinValueVT      =   1112342533
         End
         Begin TDBNumber6Ctl.TDBNumber txtRp 
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   300
            Top             =   360
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   556
            Calculator      =   "frmOrderDMSNonDSS.frx":0370
            Caption         =   "frmOrderDMSNonDSS.frx":0390
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmOrderDMSNonDSS.frx":03FC
            Keys            =   "frmOrderDMSNonDSS.frx":041A
            Spin            =   "frmOrderDMSNonDSS.frx":0464
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1397882885
            MinValueVT      =   1112342533
         End
         Begin VB.Label nAR2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 %"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   525
            Left            =   1200
            TabIndex        =   303
            Top             =   720
            Width           =   1020
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Batas AR"
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
            TabIndex        =   302
            Top             =   360
            Width           =   855
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
            Index           =   0
            Left            =   0
            TabIndex        =   301
            Top             =   60
            Width           =   1050
         End
      End
      Begin VB.Frame frmLengkap 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "         ooo000ooo - Kelengkapan - ooo000ooo           "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   3240
         TabIndex        =   280
         Top             =   3240
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton btnCancel 
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
            Left            =   2520
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   284
            Top             =   1320
            Width           =   1245
         End
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
            Left            =   1200
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   283
            Top             =   1320
            Width           =   1245
         End
         Begin VB.CheckBox opLengkap 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "SPK"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   282
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox opLengkap 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Lainnya"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   281
            Top             =   960
            Width           =   975
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   115
            Left            =   1200
            TabIndex        =   285
            Top             =   720
            Visible         =   0   'False
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            TabIndex        =   286
            Top             =   960
            Visible         =   0   'False
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Image Image8 
            Height          =   1710
            Left            =   0
            Top             =   120
            Width           =   5025
         End
      End
      Begin AIFCmp1.asxPowerButton tblCetakPO 
         Height          =   360
         Left            =   -65160
         TabIndex        =   45
         Top             =   2640
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   635
         BackColor       =   14482683
         BorderStyle     =   3
         Caption         =   ""
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
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame15"
         Height          =   2295
         Left            =   -66960
         TabIndex        =   274
         Top             =   600
         Width           =   3975
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   39
            Left            =   0
            TabIndex        =   275
            Top             =   370
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
            Height          =   345
            Index           =   40
            Left            =   0
            TabIndex        =   276
            Top             =   705
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
            Height          =   345
            Index           =   41
            Left            =   0
            TabIndex        =   277
            Top             =   1030
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
            Height          =   345
            Index           =   42
            Left            =   0
            TabIndex        =   278
            Top             =   1365
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
            Height          =   345
            Index           =   43
            Left            =   0
            TabIndex        =   279
            Top             =   1690
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Image Image11 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   3870
         End
      End
      Begin VB.CommandButton cmdLengkap 
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   268
         Top             =   3360
         Visible         =   0   'False
         Width           =   2010
      End
      Begin VB.CommandButton cmdStall 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -65160
         Style           =   1  'Graphical
         TabIndex        =   246
         Top             =   1920
         Width           =   1965
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   55
         Left            =   1200
         MaxLength       =   500
         TabIndex        =   64
         Top             =   4155
         Width           =   8255
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   54
         Left            =   1200
         MaxLength       =   500
         TabIndex        =   63
         Top             =   3840
         Width           =   8255
      End
      Begin VB.Frame frmeWaktu 
         BackColor       =   &H00C0FFFF&
         Height          =   3375
         Left            =   -62160
         TabIndex        =   180
         Top             =   960
         Visible         =   0   'False
         Width           =   6135
         Begin TrueOleDBGrid70.TDBGrid GridWaktu1 
            Height          =   3855
            Left            =   8640
            TabIndex        =   181
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
            TabIndex        =   182
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
         Begin VB.Label Label57 
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
            Left            =   240
            TabIndex        =   192
            Top             =   2880
            Width           =   4425
         End
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
         Left            =   -62160
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   1440
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.CommandButton cmdAddJob 
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
         Left            =   -65160
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   3000
         Width           =   1965
      End
      Begin VB.CommandButton cmdSublet 
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
         Left            =   -65160
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   2280
         Width           =   1965
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
         Index           =   4
         Left            =   -74970
         MaxLength       =   500
         TabIndex        =   16
         Top             =   3170
         Width           =   12615
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
         Index           =   3
         Left            =   -74970
         MaxLength       =   500
         TabIndex        =   15
         Top             =   2900
         Width           =   12615
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
         Index           =   2
         Left            =   -74970
         MaxLength       =   500
         TabIndex        =   14
         Top             =   2630
         Width           =   12615
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
         Left            =   -74970
         MaxLength       =   500
         TabIndex        =   13
         Top             =   2365
         Width           =   12615
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
         Left            =   -74970
         MaxLength       =   500
         TabIndex        =   12
         Top             =   2100
         Width           =   12615
      End
      Begin VB.CheckBox chUse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pemakaian Berat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   55
         Top             =   1095
         Width           =   1575
      End
      Begin VB.CommandButton cmdGateS 
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   3360
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.ComboBox cboPelanggan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame12 
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
         Height          =   1900
         Left            =   -65880
         TabIndex        =   177
         Top             =   3720
         Width           =   3495
         Begin MSMask.MaskEdBox text 
            Height          =   300
            Index           =   69
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
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
            Height          =   300
            Index           =   70
            Left            =   120
            TabIndex        =   50
            Top             =   750
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
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
            Height          =   300
            Index           =   71
            Left            =   120
            TabIndex        =   51
            Top             =   1020
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
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
            Height          =   300
            Index           =   72
            Left            =   120
            TabIndex        =   52
            Top             =   1290
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
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
            Height          =   300
            Index           =   73
            Left            =   120
            TabIndex        =   53
            Top             =   1560
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Image Image13 
            Height          =   390
            Left            =   120
            Top             =   100
            Width           =   3270
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   -74970
         TabIndex        =   174
         Top             =   360
         Width           =   4335
         Begin VB.OptionButton opTpSvc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Regular"
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
            TabIndex        =   3
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opTpSvc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Program"
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
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton opTpSvc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Warranty"
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
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   1095
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   21
            Left            =   1320
            TabIndex        =   6
            Top             =   480
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   270
            Left            =   2640
            TabIndex        =   175
            Top             =   480
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
            Index           =   22
            Left            =   1320
            TabIndex        =   176
            Top             =   840
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   503
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
         Begin VB.Label sgnPreDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pre'Diagnosis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   2640
            TabIndex        =   297
            Top             =   0
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.Image Image4 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   4380
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   10680
         TabIndex        =   173
         Top             =   1560
         Width           =   1815
         Begin VB.OptionButton opByr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tagihan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   186
            Top             =   825
            Width           =   1335
         End
         Begin VB.OptionButton opByr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Card"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   185
            Top             =   615
            Width           =   1335
         End
         Begin VB.OptionButton opByr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   410
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Image Image18 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   1830
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -71640
         TabIndex        =   172
         Top             =   2670
         Width           =   3885
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Empty"
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
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1/4"
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
            Index           =   3
            Left            =   960
            TabIndex        =   82
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1/2"
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
            Left            =   1680
            TabIndex        =   81
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "3/4"
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
            Left            =   2400
            TabIndex        =   80
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Full"
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
            Left            =   3120
            TabIndex        =   79
            Top             =   480
            Width           =   615
         End
         Begin VB.Image Image12 
            Height          =   390
            Left            =   0
            Top             =   50
            Width           =   3870
         End
      End
      Begin VB.CommandButton cmdFU 
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3360
         Width           =   2010
      End
      Begin VB.Frame Frame8 
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
         Height          =   2175
         Left            =   6960
         TabIndex        =   137
         Top             =   480
         Width           =   2310
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   64
            Left            =   840
            TabIndex        =   138
            Top             =   510
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   65
            Left            =   840
            TabIndex        =   139
            Top             =   805
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   66
            Left            =   840
            TabIndex        =   140
            Top             =   1110
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   67
            Left            =   840
            TabIndex        =   141
            Top             =   1410
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   78
            Left            =   840
            TabIndex        =   247
            Top             =   1710
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Image Image14 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   2310
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Book"
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
            TabIndex        =   248
            Top             =   1710
            Width           =   645
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
            Left            =   120
            TabIndex        =   144
            Top             =   840
            Width           =   570
         End
         Begin VB.Label Label25 
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
            Left            =   120
            TabIndex        =   143
            Top             =   540
            Width           =   570
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inv no."
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
            TabIndex        =   142
            Top             =   1440
            Width           =   525
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Height          =   2170
         Left            =   -74970
         TabIndex        =   125
         Top             =   3460
         Width           =   9240
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   28
            Left            =   6330
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   285
            Left            =   9090
            TabIndex        =   126
            Top             =   480
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   29
            Left            =   6330
            TabIndex        =   24
            Top             =   750
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   285
            Left            =   9090
            TabIndex        =   127
            Top             =   750
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   30
            Left            =   6330
            TabIndex        =   26
            Top             =   1020
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   285
            Left            =   9090
            TabIndex        =   128
            Top             =   1020
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   31
            Left            =   6330
            TabIndex        =   28
            Top             =   1290
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   285
            Left            =   9090
            TabIndex        =   129
            Top             =   1290
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   32
            Left            =   6330
            TabIndex        =   30
            Top             =   1560
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   285
            Left            =   9090
            TabIndex        =   130
            Top             =   1560
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   33
            Left            =   6330
            TabIndex        =   32
            Top             =   1830
            Visible         =   0   'False
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   285
            Left            =   9090
            TabIndex        =   131
            Top             =   1830
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Image Image7 
            Height          =   390
            Left            =   0
            Top             =   40
            Width           =   9225
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  (Berhenti, Mesin dingin, Mesin Panas, Perpindahan Gigi, lainnya)"
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
            Left            =   885
            TabIndex        =   232
            Top             =   1860
            Width           =   4695
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  (Starter, Pemanasan, Jalan Konstan, Tambah/kurang Kecepatan, lainnya)"
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
            Left            =   885
            TabIndex        =   231
            Top             =   1590
            Width           =   5415
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  (Lampu menyala, Kedip sekali, Berkedip-kedip, lainnya)"
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
            Left            =   885
            TabIndex        =   230
            Top             =   1320
            Width           =   4020
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  (Jalan datar, Jalan Tol, Tanjakan Macet, Tanjakan Sepi, Bergelombang)"
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
            Left            =   885
            TabIndex        =   229
            Top             =   1050
            Width           =   5235
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  (Selalu, Kadang-kadang, Hanya sekali, lainnya)"
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
            Left            =   885
            TabIndex        =   228
            Top             =   780
            Width           =   3480
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "  (Akhir-akhir ini, Seminggu lalu, Lainnya)"
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
            Left            =   885
            TabIndex        =   227
            Top             =   525
            Width           =   2910
         End
         Begin VB.Label Label32 
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
            Left            =   240
            TabIndex        =   136
            Top             =   1590
            Width           =   555
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lampu"
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
            TabIndex        =   135
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label23 
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
            Left            =   360
            TabIndex        =   134
            Top             =   1050
            Width           =   435
         End
         Begin VB.Label Label22 
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
            Left            =   105
            TabIndex        =   133
            Top             =   780
            Width           =   690
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terjadi"
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
            Left            =   300
            TabIndex        =   132
            Top             =   525
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   270
         Index           =   23
         Left            =   -66720
         TabIndex        =   7
         Top             =   765
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   24
         Left            =   -66720
         TabIndex        =   8
         Top             =   1000
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   25
         Left            =   -66720
         TabIndex        =   9
         Top             =   1240
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   26
         Left            =   -66720
         TabIndex        =   10
         Top             =   1480
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   27
         Left            =   -66720
         TabIndex        =   11
         Top             =   1720
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   52
         Left            =   2760
         TabIndex        =   57
         Top             =   1380
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   53
         Left            =   1200
         TabIndex        =   69
         Top             =   4155
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   49
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   54
         Left            =   1200
         TabIndex        =   67
         Top             =   3840
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   55
         Left            =   1200
         TabIndex        =   68
         Top             =   3840
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   345
         Index           =   34
         Left            =   -71640
         TabIndex        =   34
         Top             =   970
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
         Height          =   345
         Index           =   35
         Left            =   -71640
         TabIndex        =   35
         Top             =   1310
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
         Height          =   345
         Index           =   36
         Left            =   -71640
         TabIndex        =   36
         Top             =   1640
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
         Height          =   345
         Index           =   37
         Left            =   -71640
         TabIndex        =   37
         Top             =   1965
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
         Height          =   345
         Index           =   38
         Left            =   -71640
         TabIndex        =   38
         Top             =   2300
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
         Index           =   74
         Left            =   -65520
         TabIndex        =   23
         Top             =   3600
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   75
         Left            =   1950
         TabIndex        =   59
         Top             =   1695
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Index           =   86
         Left            =   3240
         TabIndex        =   211
         Top             =   2280
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   88
         Left            =   1200
         TabIndex        =   60
         Top             =   2600
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   89
         Left            =   2430
         TabIndex        =   61
         Top             =   2600
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
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
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   90
         Left            =   1200
         TabIndex        =   213
         Top             =   480
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   91
         Left            =   -65520
         TabIndex        =   25
         Top             =   3915
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   92
         Left            =   -65520
         TabIndex        =   27
         Top             =   4215
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   93
         Left            =   -65520
         TabIndex        =   29
         Top             =   4515
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   94
         Left            =   -65520
         TabIndex        =   31
         Top             =   4815
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   95
         Left            =   -65520
         TabIndex        =   33
         Top             =   5115
         Visible         =   0   'False
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   240
         Index           =   97
         Left            =   -62280
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
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
         Height          =   240
         Index           =   98
         Left            =   -62280
         TabIndex        =   18
         Top             =   825
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
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
         Height          =   240
         Index           =   99
         Left            =   -62280
         TabIndex        =   19
         Top             =   1050
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
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
         Height          =   240
         Index           =   100
         Left            =   -62280
         TabIndex        =   20
         Top             =   1275
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
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
         Height          =   240
         Index           =   101
         Left            =   -62280
         TabIndex        =   21
         Top             =   1500
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
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
         Index           =   102
         Left            =   1200
         TabIndex        =   65
         Top             =   5055
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   103
         Left            =   1200
         TabIndex        =   54
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   270
         Left            =   2640
         TabIndex        =   224
         Top             =   1080
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
         Index           =   105
         Left            =   5175
         TabIndex        =   66
         Top             =   5055
         Width           =   4270
         _ExtentX        =   7541
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
         Height          =   285
         Index           =   62
         Left            =   1200
         TabIndex        =   240
         Top             =   2280
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   63
         Left            =   2430
         TabIndex        =   241
         Top             =   2280
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   107
         Left            =   1200
         TabIndex        =   56
         Top             =   1380
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox MaskEdBox15 
         Height          =   270
         Left            =   2640
         TabIndex        =   243
         Top             =   1380
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   10680
         TabIndex        =   178
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton opGanti 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Konfirmasi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opGanti 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tanpa Konfirmasi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   179
            Top             =   720
            Width           =   1575
         End
         Begin VB.Image Image16 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   1830
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   885
         Index           =   51
         Left            =   -65640
         TabIndex        =   249
         ToolTipText     =   "Jangan Sampai Keliru Mengisi KM !!!"
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1561
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   108
         Left            =   1200
         TabIndex        =   251
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox MaskEdBox16 
         Height          =   270
         Left            =   2640
         TabIndex        =   252
         Top             =   780
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
         Index           =   109
         Left            =   2760
         TabIndex        =   254
         Top             =   780
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   114
         Left            =   7080
         TabIndex        =   265
         Top             =   2760
         Visible         =   0   'False
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   3025
         Left            =   -74970
         TabIndex        =   166
         Top             =   360
         Width           =   8880
         Begin VB.CheckBox chSBNP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "SB NP"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7800
            TabIndex        =   245
            Top             =   630
            Width           =   855
         End
         Begin VB.CheckBox chEM 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "EM"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7800
            TabIndex        =   244
            Top             =   400
            Width           =   615
         End
         Begin VB.CheckBox chPDS 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "PDS"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7800
            TabIndex        =   40
            Top             =   180
            Width           =   735
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   46
            Left            =   600
            TabIndex        =   42
            Top             =   555
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   503
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
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   270
            Left            =   6120
            TabIndex        =   167
            Top             =   555
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
            Index           =   44
            Left            =   600
            TabIndex        =   41
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Left            =   1125
            TabIndex        =   168
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
            Index           =   45
            Left            =   1245
            TabIndex        =   169
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   47
            Left            =   4320
            TabIndex        =   83
            Top             =   1800
            Visible         =   0   'False
            Width           =   585
            _ExtentX        =   1032
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
            Index           =   50
            Left            =   5295
            TabIndex        =   43
            Top             =   240
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   270
            Left            =   6120
            TabIndex        =   190
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
            Index           =   82
            Left            =   6240
            TabIndex        =   199
            Top             =   240
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
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
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   96
            Left            =   6600
            TabIndex        =   215
            Top             =   1680
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
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
            Index           =   106
            Left            =   6600
            TabIndex        =   233
            Top             =   1560
            Visible         =   0   'False
            Width           =   675
            _ExtentX        =   1191
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
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   2025
            Left            =   15
            TabIndex        =   44
            Top             =   960
            Width           =   8820
            _ExtentX        =   15558
            _ExtentY        =   3572
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            HighLight       =   2
            ScrollBars      =   0
            SelectionMode   =   1
            BorderStyle     =   0
            Appearance      =   0
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mekanik"
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
            TabIndex        =   191
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order"
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
            TabIndex        =   171
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
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
            TabIndex        =   170
            Top             =   285
            Width           =   300
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   9360
         TabIndex        =   157
         Top             =   480
         Width           =   1355
         Begin VB.OptionButton opSvc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton opSvc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   84
            Top             =   720
            Width           =   975
         End
         Begin VB.Image Image15 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   1350
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   9360
         TabIndex        =   217
         Top             =   1560
         Width           =   1355
         Begin VB.OptionButton opWash 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tidak"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   218
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton opWash 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ya"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.Image Image17 
            Height          =   390
            Left            =   0
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2920
         Left            =   9720
         TabIndex        =   145
         Top             =   2560
         Width           =   2715
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   56
            Left            =   900
            TabIndex        =   146
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
            Index           =   57
            Left            =   900
            TabIndex        =   147
            Top             =   900
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
            Height          =   390
            Index           =   61
            Left            =   900
            TabIndex        =   148
            Top             =   2520
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   0
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
            Index           =   58
            Left            =   900
            TabIndex        =   76
            Top             =   1805
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   59
            Left            =   900
            TabIndex        =   149
            Top             =   1200
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
            Index           =   60
            Left            =   900
            TabIndex        =   150
            Top             =   1505
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
            Index           =   77
            Left            =   900
            TabIndex        =   77
            Top             =   2105
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
         Begin VB.Image Image19 
            Height          =   390
            Left            =   0
            Top             =   120
            Width           =   2685
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   0
            X2              =   2640
            Y1              =   2445
            Y2              =   2445
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Body Rpr"
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
            TabIndex        =   189
            Top             =   2205
            Width           =   660
         End
         Begin VB.Label Label41 
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
            Left            =   150
            TabIndex        =   156
            Top             =   1860
            Width           =   570
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jasa"
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
            Left            =   390
            TabIndex        =   155
            Top             =   645
            Width           =   330
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
            Left            =   420
            TabIndex        =   154
            Top             =   960
            Width           =   300
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            TabIndex        =   153
            Top             =   2565
            Width           =   645
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
            Left            =   270
            TabIndex        =   152
            Top             =   1560
            Width           =   450
         End
         Begin VB.Label Label46 
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
            Left            =   540
            TabIndex        =   151
            Top             =   1260
            Width           =   180
         End
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   53
         Left            =   1200
         MaxLength       =   249
         TabIndex        =   304
         Top             =   4460
         Width           =   8255
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   87
         Left            =   1200
         MaxLength       =   250
         TabIndex        =   305
         Top             =   4760
         Width           =   8255
      End
      Begin MSMask.MaskEdBox text 
         Height          =   270
         Index           =   87
         Left            =   1200
         TabIndex        =   70
         Top             =   4455
         Width           =   6360
         _ExtentX        =   11218
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Frame FrameFill 
         BackColor       =   &H00C0FFFF&
         Height          =   2400
         Left            =   -75000
         TabIndex        =   311
         Top             =   3360
         Visible         =   0   'False
         Width           =   8865
         Begin VB.Frame FrDiag 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "FrDiag"
            Height          =   2260
            Left            =   -8760
            TabIndex        =   312
            Top             =   120
            Width           =   8895
         End
         Begin TrueOleDBGrid70.TDBGrid partGridFill 
            Height          =   1785
            Left            =   0
            TabIndex        =   313
            Top             =   600
            Width           =   8850
            _ExtentX        =   15610
            _ExtentY        =   3149
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE PART"
            Columns(0).DataField=   "KODEITEM"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA PART"
            Columns(1).DataField=   "NAMAITEM"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "NAMA BY SA"
            Columns(2).DataField=   "NAMABYSA"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "INFO"
            Columns(3).DataField=   "SATITEM"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "QTY"
            Columns(4).DataField=   "QTY"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "OLI"
            Columns(5).DataField=   "OLI"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2910"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2831"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=4154"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4075"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(15)=   "Column(3).Width=2117"
            Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2037"
            Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(20)=   "Column(4).Width=1244"
            Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1164"
            Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
            Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(25)=   "Column(5).Width=873"
            Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=794"
            Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=8193"
            Splits(0)._ColumnProps(29)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2,.locked=-1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   119
            Left            =   960
            TabIndex        =   314
            Top             =   195
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox MaskEdBox17 
            Height          =   270
            Left            =   3900
            TabIndex        =   315
            Top             =   195
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
            Index           =   120
            Left            =   4080
            TabIndex        =   316
            Top             =   195
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin ucXPButton.XPButton cmdView 
            Height          =   460
            Left            =   6720
            TabIndex        =   318
            Top             =   120
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   820
            BTYPE           =   2
            TX              =   "&w"
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
            MICON           =   "frmOrderDMSNonDSS.frx":048C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label29 
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
            TabIndex        =   317
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   2400
         Left            =   -74970
         TabIndex        =   163
         Top             =   3360
         Width           =   8885
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   1780
            Left            =   0
            TabIndex        =   48
            Top             =   600
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   3149
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE PART"
            Columns(0).DataField=   "KODEITEM"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA PART"
            Columns(1).DataField=   "NAMAITEM"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "SATUAN"
            Columns(2).DataField=   "SATITEM"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "QTY"
            Columns(3).DataField=   "QTY"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   4
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "GANTI"
            Columns(4).DataField=   "GANTI"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "BAWA"
            Columns(5).DataField=   "BAWA"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "OLI"
            Columns(6).DataField=   "OLI"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2910"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2831"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=4154"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4075"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2117"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2037"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=1244"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1164"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=1402"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1323"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=1402"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1323"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=1667"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1588"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8196"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.locked=-1"
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   48
            Left            =   960
            TabIndex        =   46
            Top             =   195
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Left            =   3900
            TabIndex        =   164
            Top             =   195
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
            Left            =   4080
            TabIndex        =   47
            Top             =   195
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Label Label42 
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
            TabIndex        =   165
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Label sgnStall 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STALL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   -64650
         TabIndex        =   322
         Top             =   1560
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label LbSupply 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   -68520
         TabIndex        =   320
         Top             =   0
         Width           =   3015
      End
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
         Left            =   -72195
         TabIndex        =   306
         Top             =   1755
         Width           =   5490
      End
      Begin VB.Image Image10 
         Height          =   390
         Left            =   -71640
         Top             =   600
         Width           =   3870
      End
      Begin VB.Image Image9 
         Height          =   390
         Left            =   -74400
         Top             =   600
         Width           =   2040
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -74970
         Top             =   1725
         Width           =   12600
      End
      Begin VB.Image Image5 
         Height          =   390
         Left            =   -66720
         Top             =   360
         Width           =   4380
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Antrian"
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
         TabIndex        =   253
         Top             =   795
         Width           =   525
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "K. Meter :"
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
         Left            =   -64560
         TabIndex        =   250
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selesai"
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
         TabIndex        =   242
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label67 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A/N Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   223
         Top             =   5040
         Width           =   1005
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   53
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   52
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   51
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   50
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   49
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   48
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   47
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   46
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   45
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   44
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   43
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   42
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   41
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   40
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   39
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   38
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   37
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   36
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   35
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   34
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   33
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   32
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   31
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   30
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   29
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   28
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   27
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   26
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   25
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   24
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   23
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   22
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   21
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   20
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   19
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   18
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   17
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   16
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   15
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   14
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   13
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   12
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   11
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   10
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   9
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   8
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   7
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   6
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   5
         Left            =   -72840
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   4
         Left            =   -73080
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   3
         Left            =   -73320
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   2
         Left            =   -73560
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label sgnPart 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PESAN PART"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -65880
         TabIndex        =   216
         Top             =   3360
         Visible         =   0   'False
         Width           =   3420
      End
      Begin VB.Label Label64 
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
         Left            =   120
         TabIndex        =   214
         Top             =   525
         Width           =   345
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penyerahan"
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
         TabIndex        =   212
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label Label34 
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
         Left            =   120
         TabIndex        =   184
         Top             =   1720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image ImgCrop 
         Height          =   2100
         Left            =   -74040
         Top             =   1200
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   1
         Left            =   -73800
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   0
         Left            =   -74040
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   210
      End
      Begin VB.Image Image1 
         Height          =   2970
         Index           =   0
         Left            =   -74400
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesan"
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
         TabIndex        =   162
         Top             =   3840
         Width           =   435
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
         Left            =   120
         TabIndex        =   161
         Top             =   4170
         Width           =   420
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         TabIndex        =   160
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "Alasan Datang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   159
         Top             =   1400
         Width           =   1050
      End
      Begin VB.Label Label10 
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
         TabIndex        =   158
         Top             =   4500
         Width           =   585
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5325
         Left            =   -75000
         Top             =   315
         Width           =   12705
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5325
         Left            =   -74985
         Top             =   360
         Width           =   12705
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5325
         Left            =   -75000
         Top             =   300
         Width           =   12705
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5325
         Left            =   20
         Top             =   320
         Width           =   12675
      End
   End
   Begin VB.CheckBox chTHS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unit THS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9840
      TabIndex        =   219
      Top             =   120
      Width           =   1440
   End
   Begin VB.Frame Frame7 
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
      Height          =   2730
      Left            =   9840
      TabIndex        =   112
      Top             =   360
      Width           =   3375
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
         Left            =   -240
         Style           =   1  'Graphical
         TabIndex        =   307
         Top             =   2160
         Width           =   1935
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
         Height          =   375
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   308
         Top             =   2160
         Width           =   1935
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   113
         Top             =   795
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   340
         Index           =   20
         Left            =   1080
         TabIndex        =   114
         Top             =   1695
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   19
         Left            =   1080
         TabIndex        =   115
         Top             =   1395
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   18
         Left            =   1080
         TabIndex        =   119
         Top             =   1095
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   16
         Left            =   1080
         TabIndex        =   270
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label13 
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
         Left            =   360
         TabIndex        =   310
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label Label24 
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
         Left            =   2040
         TabIndex        =   309
         Top             =   2520
         Width           =   930
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   0
         Top             =   40
         Width           =   3375
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
         Left            =   120
         TabIndex        =   271
         Top             =   525
         Width           =   585
      End
      Begin VB.Label Label75 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "History"
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
         Left            =   4200
         TabIndex        =   234
         Top             =   0
         Width           =   195
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. RC."
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
         TabIndex        =   120
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Service"
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
         TabIndex        =   118
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   117
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label15 
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
         Left            =   0
         TabIndex        =   116
         Top             =   1395
         Width           =   585
      End
   End
   Begin VB.CheckBox ChkReturnOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11400
      TabIndex        =   220
      Top             =   120
      Width           =   1800
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   345
      Left            =   5280
      TabIndex        =   106
      Top             =   60
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   68
      Left            =   2760
      TabIndex        =   107
      Top             =   8940
      Width           =   620
      _ExtentX        =   1085
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
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   345
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   60
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
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
      Index           =   104
      Left            =   3390
      TabIndex        =   226
      Top             =   8940
      Width           =   720
      _ExtentX        =   1270
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
      PromptChar      =   ">"
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
      Height          =   2610
      Left            =   480
      TabIndex        =   98
      Top             =   480
      Width           =   9405
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   690
         TabIndex        =   92
         Top             =   750
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   4
         Left            =   690
         TabIndex        =   93
         Top             =   1045
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   690
         TabIndex        =   94
         Top             =   1650
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   6
         Left            =   690
         TabIndex        =   95
         Top             =   1945
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   6030
         TabIndex        =   86
         Top             =   450
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   10
         Left            =   6030
         TabIndex        =   87
         Top             =   750
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   6870
         TabIndex        =   88
         Top             =   1350
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   6030
         TabIndex        =   97
         Top             =   1050
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Left            =   8490
         TabIndex        =   96
         Top             =   750
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   2
         Left            =   2400
         TabIndex        =   91
         Top             =   450
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   1
         Left            =   690
         TabIndex        =   2
         Top             =   450
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Left            =   6870
         TabIndex        =   89
         Top             =   1650
         Width           =   1410
         _ExtentX        =   2487
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
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   6030
         TabIndex        =   90
         Top             =   1950
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   690
         TabIndex        =   123
         Top             =   2245
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   3210
         TabIndex        =   124
         Top             =   2245
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   2235
         TabIndex        =   183
         Top             =   450
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
         Index           =   76
         Left            =   6030
         TabIndex        =   187
         Top             =   2245
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   79
         Left            =   1095
         TabIndex        =   194
         Top             =   3120
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
         Index           =   80
         Left            =   5220
         TabIndex        =   196
         Top             =   3360
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
         Index           =   81
         Left            =   5220
         TabIndex        =   198
         Top             =   3000
         Width           =   1860
         _ExtentX        =   3281
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
         Index           =   117
         Left            =   690
         TabIndex        =   269
         Top             =   1350
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   118
         Left            =   8040
         TabIndex        =   272
         Top             =   450
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Image Image2 
         Height          =   390
         Left            =   0
         Top             =   0
         Width           =   9270
      End
      Begin VB.Label Label85 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   6960
         TabIndex        =   273
         Top             =   495
         Width           =   975
      End
      Begin VB.Label Label59 
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
         Left            =   4560
         TabIndex        =   197
         Top             =   3015
         Width           =   585
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDate"
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
         Left            =   495
         TabIndex        =   195
         Top             =   3165
         Width           =   450
      End
      Begin VB.Label SgnRekom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rekom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8520
         TabIndex        =   193
         Top             =   1687
         Visible         =   0   'False
         Width           =   630
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
         Left            =   5340
         TabIndex        =   188
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label9 
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
         Left            =   2700
         TabIndex        =   122
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label8 
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
         Left            =   105
         TabIndex        =   121
         Top             =   2280
         Width           =   315
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
         Left            =   5370
         TabIndex        =   111
         Top             =   1980
         Width           =   510
      End
      Begin VB.Label Label1 
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
         Left            =   6090
         TabIndex        =   110
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kd Lgn"
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
         TabIndex        =   109
         Top             =   480
         Width           =   480
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
         Left            =   6075
         TabIndex        =   105
         Top             =   1395
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
         Left            =   5490
         TabIndex        =   104
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
         Left            =   5265
         TabIndex        =   103
         Top             =   465
         Width           =   645
      End
      Begin VB.Label Label5 
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
         Left            =   90
         TabIndex        =   102
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label Label4 
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
         Left            =   105
         TabIndex        =   101
         Top             =   1695
         Width           =   330
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   100
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   99
         Top             =   795
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdPilih 
      Appearance      =   0  'Flat
      BackColor       =   &H00BFE8EA&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "No. Booking"
      Top             =   40
      Width           =   1965
   End
   Begin VB.CommandButton cmdWarranty 
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
      Height          =   405
      Left            =   5260
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   8880
      Width           =   2010
   End
   Begin VB.CommandButton cmdCetak 
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
      Left            =   7240
      Style           =   1  'Graphical
      TabIndex        =   239
      Top             =   8880
      Width           =   2010
   End
   Begin VB.CommandButton TblSimpan 
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   8880
      Width           =   2010
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
      Left            =   11210
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   8880
      Width           =   2010
   End
   Begin VB.CommandButton cmdKodePlg 
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
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   40
      Width           =   2010
   End
   Begin VB.CommandButton cmdGantiNoPol 
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   255
      Top             =   40
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton cmdCekHO 
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
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   40
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label49 
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
      Left            =   0
      TabIndex        =   326
      Top             =   60
      Width           =   300
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   128
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrderDMSNonDSS.frx":04A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrderDMSNonDSS.frx":0506
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrderDMSNonDSS.frx":0564
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   460
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Order Masuk Hari ini :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   615
      TabIndex        =   108
      Top             =   8940
      Width           =   2145
   End
End
Attribute VB_Name = "frmOrderDMSNonDSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsTmp As ADODB.Recordset
Dim rsHlp As ADODB.Recordset
Dim rsGridPart As New ADODB.Recordset
Dim rsOli As New ADODB.Recordset
Dim PSQL As String
Dim pesan As String
Dim hNamaPart As String
Dim hSatPart As String
Dim hOliPart As Boolean
Dim boleh As Boolean
Dim fOther As Boolean
Public baris As Integer
Dim barispart As Integer
Dim Jrate
Dim hHrgPart As Double
Public ntRate As Double
Dim nk As Integer
Dim nStok As Integer
Dim gSBE As Boolean
Dim bKondisi(6) As String
Dim KorX(100) As Integer
Dim KorY(100) As Integer
Dim fPassword As String
Dim nTipe As String
Dim nWKT As String
Dim gPO As String
'Dim bantuan As New BclsHelp
Dim simpan As Boolean
Dim custKetemu As Boolean
Dim onlySublet As Boolean
Dim newWo As Boolean
Dim newPre As Boolean
Public woOnProgres As Boolean
Dim bodyClear As Boolean
Dim Updated As Boolean
Dim noAntriLama As String
Dim lastService As Double
Dim support As New clsFormDealer

Const MaxBrs As Integer = 7
Private tipeAntrian As String
Private isVAR As Boolean
Private nPersenAR As Double
Private isSBI As Boolean
Private isUpdateFill As Boolean
Private NewId As Double
Private adjust As String
Private totalSublet As Double

Private isWashing As Boolean 'jika washing maka true

Private rsPart1 As New ADODB.Recordset
Private PassRSPart As New ADODB.Recordset

Public nilaiDPPart As Integer


Public Property Let passingType(VALUES As String)
    tipeAntrian = VALUES
    text(109) = tipeAntrian
End Property

Public Property Let getTotalSublet(VALUES As Double)
    totalSublet = VALUES
End Property

Public Property Let passingRS(VALUES As ADODB.Recordset)
    Set rsGridPart = VALUES
End Property

Public Property Let GetAdjustStall(VALUES As String)
     adjust = VALUES
End Property

Public Property Let isNowID(VALUES As Double)
    NewId = VALUES
End Property

Private Sub Command2_Click()
    If Trim(txtAlasanBtlWashing.text) <> "" Then
        If Len(txtAlasanBtlWashing.text) < 5 Then
            MsgBox "Alasan Pembatalan Washing minimal 5 karakter!", vbCritical, "Informasi Pembatalan Washing"
        Else
            objO.InsertData ("INSERT INTO WASHING VALUES ('" & Trim(text(64)) & "',GETDATE(),'" & gUser & "','" & txtAlasanBtlWashing.text & "')")
            MsgBox "Proses Pembatalan Washing berhasil dilakukan", vbInformation, "Informasi Pembatalan Washing"
            txtAlasanBtlWashing.text = ""
            framePembatalanWashing.Visible = False
            opWash(1).value = True
            
        End If
    Else
        MsgBox "Alasan Pembatalan Washing harus diisi!", vbCritical, "Informasi Pembatalan Washing"
    End If
End Sub

Private Sub Command3_Click()
    If Trim(text(123)) <> "" Then
        If CInt(Trim(text(123))) >= CInt(Trim(text(122))) Then
            nilaipart = CInt(Trim(text(123)))
        Else
            MsgBox "Nilai DP minimal " & Trim(text(122)), vbCritical
        End If
    Else
        MsgBox "Nilai Part Harus diisi!", vbCritical
    End If
End Sub

Private Sub Form_Activate()
    If validAktif Or subletAktif Or subRetAktif Then
        MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
        Unload Me
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
        cmdPilih.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    
    orderAktif = True
    'Timer1.Enabled = True
    text(104) = Trim(gSA)
    If Layar.Tag = "estOrder" Then
        baris = gBantuan
        barispart = gBantuan2
        ntRate = CDbl(gBantuan3)
        text(114) = ntRate
        If rsTmpPart.State <> adStateClosed Then
            If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                If rsGridPart.State Then rsGridPart.Close
                Set rsGridPart = rsTmpPart
                partGridFill.DataSource = rsGridPart
            Else
                If rsGridPart.State Then rsGridPart.Close
                Set rsGridPart = rsTmpPart
                partGrid.DataSource = rsGridPart
            End If

            If rsGridPart.State <> adStateClosed Then
                text(59) = 0
                text(57) = 0
                If Not rsGridPart.EOF Then
                rsGridPart.MoveFirst
                Do While Not rsGridPart.EOF
                    Set rsOli = objO.FetchDataACC("SELECT hargaJualA FROM inmst WITH (NOLOCK) WHERE kodeItem = '" & rsGridPart("KODEITEM") & "'")
                    If Not rsOli.EOF Then
                        If rsGridPart("OLI") Then
                            text(59) = text(59) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                        Else
                            text(57) = text(57) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                        End If
                    End If
                    rsGridPart.MoveNext
                Loop
                End If
            End If
            
        End If
        Call bukaForm
        
    End If
    nk = 0
    'cmdPilih.tag = "No. Booking"
    cmdPilih.Picture = ImageList1.ListImages(1).Picture
    SSTab1.Tab = 0
    SSTab1.Enabled = False
    text(68) = support.countUnitEntry
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    'asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Timer1.Enabled = False
    orderAktif = False
End Sub

Private Sub tutupForm()
    For i = 1 To 105
        text(i).BackColor = &HC0E0FF
    Next
    For i = 0 To 4
        Text1(i).BackColor = &HC0E0FF
    Next
End Sub

Private Sub bukaForm()
    For i = 1 To 105
        Select Case i
            Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 76, 45, 96, 82, 78, 90, 56, 57, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 86, 84, 104
                text(i).BackColor = &HC0E0FF
            Case Else
                text(i).BackColor = &HFFFFFF
        End Select
    Next
    For i = 0 To 4
        Text1(i).BackColor = &HFFFFFF
    Next
End Sub

Private Sub hapusForm()
    text(65) = TglDDMMYYYY(isnow)
    Dim i As Double
    If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
        FrameFill.Visible = True
        Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status = '" & gSA & "'")
    Else
        
    End If
    cmdHistory.Enabled = True
    ' HAPUS TEMP
    If Trim(Replace(text(15), ">", "")) <> "" Then
        Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_CHASSIS = '" & Trim(text(15)) & "'")    'NO_WO = '" &Trim(text(64)) & "' AND
        Call objO.DeleteData("DELETE TEMPFU WHERE NO_CHASSIS = '" & Trim(text(15)) & "'")   'NO_WO = '" &Trim(text(64)) & "' AND
    End If
    If Trim(Replace(text(14), ">", "")) <> "" Then
        Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN LIKE '" & Trim(text(14)) & " |%'")  ' AND MECH = '" & Trim(Text(50)) & "'
    End If
    If opTpSvc(2) Then
        If Trim(Replace(text(15), ">", "")) <> "" Then
            Call cekWarranty(Trim(text(15)), Trim(text(21)))
        End If
    End If
    
    simpan = True
    Updated = True
    For i = 1 To 118
        Select Case i
            Case 9, 20, 49, 51, 56, 57, 58, 59, 60, 61, 77, 96, 108, 114
                text(i) = 0
            Case 6, 7, 8
                text(i).Mask = ""
                text(i) = ">>>>->>>>>>>>>>"
                text(i).Mask = "9999-9999999999"
            Case 16, 17, 18, 62, 65, 79, 88, 118
                text(i).Mask = ""
                text(i) = ">>->>->>>>"
                text(i).Mask = "99-99-9999"
            Case 47
                text(i).Mask = ""
                text(i) = "00,00"
                text(i).Mask = "99,99"
            Case 63, 66, 86, 89
                text(i).Mask = ""
                text(i) = ">>:>>"
                text(i).Mask = "99:99"
            Case 111
                text(i).Mask = ""
                text(i) = ">>"
                text(i).Mask = "??"
            Case 112
                text(i).Mask = ""
                text(i) = ">>>>"
                text(i).Mask = "9999"
            Case 113
                text(i).Mask = ""
                text(i) = ">>>"
                text(i).Mask = "CCC"
            Case Else
                text(i) = ""
        End Select
    Next
'    For i = 1 To 105
'        Select Case i
'            Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 22, 76, 45, 96, 82, 78, 90, 56, 57, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 86, 84, 104
'                Text(i).BackColor = &HC0E0FF
'            Case Else
'                Text(i).BackColor = &HFFFFFF
'        End Select
'    Next
    Text1(0).text = ""
    Text1(1).text = ""
    Text1(2).text = ""
    Text1(3).text = ""
    Text1(4).text = ""
    For i = 0 To 100
        KorX(i) = 0
        KorY(i) = 0
    Next
'    For i = 0 To 4
'        Text1(i).BackColor = &HFFFFFF
'    Next
    For i = 0 To 53
        imgTanda(i).Visible = False
    Next

    SgnRekom.Visible = False
    chUse.value = Unchecked
    opBBM(4).value = True
    'CboStatus = "EXT"
    text(103) = "EXT"
    chPDS = Unchecked
    ChkReturnOrder.Enabled = False
    ntRate = 0
    text(44).BackColor = &H80000005
    text(46).BackColor = &H80000005
    SSTab1.Enabled = False
    cmdWarranty.Enabled = False
    
    Call aturGridJob
    For i = 0 To 7
        DGrid.TextMatrix(i, 1) = ""
        DGrid.TextMatrix(i, 2) = ""
        DGrid.TextMatrix(i, 3) = ""
        DGrid.TextMatrix(i, 4) = ""
        DGrid.TextMatrix(i, 5) = ""
        DGrid.TextMatrix(i, 6) = ""
    Next
    
    bKondisi(0) = ""
    bKondisi(1) = ""
    bKondisi(2) = ""
    bKondisi(3) = ""
    bKondisi(4) = ""
    bKondisi(5) = ""
    
    text(74).Visible = False
    text(91).Visible = False
    text(92).Visible = False
    text(93).Visible = False
    text(94).Visible = False
    text(95).Visible = False
    
    Set rsGridPart = Nothing
        
    ' ISI KARAKTER PELANGGAN
''    Set rsTemp = objO.FetchData("SELECT KODE FROM T_Karakter ORDER BY KODE")
''    If Not rsTemp.EOF Then
''        cboPelanggan.Clear
''        Do While Not rsTemp.EOF
''            cboPelanggan.AddItem (rsTemp("KODE"))
''            rsTemp.MoveNext
''        Loop
''    End If
    cboPelanggan.Clear
    cboPelanggan.AddItem "*"
    cboPelanggan.AddItem "**"
    cboPelanggan.AddItem "***"
    cboPelanggan.AddItem "****"
    cboPelanggan.AddItem "*****"

    cboPelanggan.ListIndex = 0
    nk = 0
    
    cmdStall.Tag = "belum"
    Set rsGridPart = Nothing
    opTpSvc(0).Enabled = True
    opTpSvc(1).Enabled = True
    opTpSvc(2).Enabled = True
    opTpSvc(0).value = True
    opGanti(0).value = True
    'opSvc(0).value = True
    '--- pergantian status tunggu ----
    opSvc(0).value = False
    opSvc(1).value = False
    
    opByr(0).value = True
    SSTab1.Tab = 0
    chPDS.Enabled = True
    
    woOnProgres = False
    text(68) = support.countUnitEntry
    noAntriLama = ""
    TblSimpan.Enabled = True
    updateHistory = False
    chEM.value = Unchecked
    chPDS.value = Unchecked
    chSBNP.value = Unchecked
    
    '' KELENGKAPAN
    opLengkap(0) = Unchecked
    opLengkap(1) = Unchecked
    lastService = 0
    isSBI = False
    isJualNRM = False
    FrDiag.Left = -9000
    NewId = 1
    
    'info stall
    sgnStall.Caption = ""
    sgnStall.Visible = False
End Sub

Private Sub CetakForm1()
    'Dim rpt As New RptOrder3            ' KGW
    Dim rpt As Object
    Select Case gPrintCetak
        Case "2"
            Set rpt = New RptOrder4            ' KGW + PEMUDA ' 2190
        Case "1"
            Set rpt = New RptOrder5            ' MLATI + ALL  ' 2180
        Case "3"
            Set rpt = New RptOrder6            ' SLAMET  ' LX 300+II
        Case "4"
            Set rpt = New RptOrder7            ' BLANKO UTK KERTAS KOSONG
            rpt.txtDealer.SetText support.kopOrder(0)
    End Select
    
    If MsgBox("Mencetak Order ?", vbYesNo, "Konfirmasi") = vbYes Then
        gNoOrder = Trim(text(64)) & "/" & Trim(text(103)) 'CboStatus.text
        gKetSA = gSA
        
        NamaSAOrder = objO.FetchData_String("Select FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS ='" & Trim(gKetSA) & "'")
        
        gCust = text(3)
        gAlmt = text(4) + text(117)
        gKota = text(5)
        gTelpon = text(6)
        gCatatan = Txt(53) & " " & Txt(87)
        gTahun = text(9)
        gNoPolisi = text(14)
        gModel = text(11)
        gWarna = text(13)
        gTglMasuk = text(65)
        gJamMasuk = text(66)
        gDealer = text(80)
        gDD = Replace(text(79), ">", "")
        gMDLOrder = text(10)
        gFRM = text(15)
        gEG = text(76)
        gKeluhan(1) = text(23)
        gKeluhan(2) = text(24)
        gKeluhan(3) = text(25)
        gKeluhan(4) = text(26)
        gKeluhan(5) = text(27)
        gOrder(1) = DGrid.TextMatrix(1, 1)
        gOrder(2) = DGrid.TextMatrix(2, 1)
        gOrder(3) = DGrid.TextMatrix(3, 1)
        gOrder(4) = DGrid.TextMatrix(4, 1)
        gOrder(5) = DGrid.TextMatrix(5, 1)
        gOrder(6) = DGrid.TextMatrix(6, 1)
        gOrder(7) = DGrid.TextMatrix(7, 1)
        
        Set rsHlp = objO.FetchData("SELECT NAMABARANG FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_CHASSIS = '" & text(15) & "'")
        If Not rsHlp.EOF Then
            i = 1
            Do While Not rsHlp.EOF
                If i < 10 Then
                    gSublet(i) = rsHlp("NAMABARANG")
                    i = i + 1
                End If
                rsHlp.MoveNext
            Loop
        End If
        
        gKM = Format(text(51), "#,#0")
        gPKhusus = Txt(54)
        gSAdvisor = NamaSAOrder
        gPSelesai = text(63)
        gKodePlg = text(1)
        
        Dim sHlp As Object  'New ADODB.Recordset
        Dim sHlp2 As Object
        Dim sHlp3 As Object
        Dim sHlp4 As Object
        Dim crxSubReportHistory As CRAXDRT.Report
        Dim crxSubReportBooking As CRAXDRT.Report
        Dim crxSubReportBooking2 As CRAXDRT.Report
        Dim crxSubReportBookingFU As CRAXDRT.Report
        If gUkuranOrder = False Then
        
            Dim rRs As ADODB.Recordset
            Set rRs = objO.FetchData("SELECT A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.NO_POLISI, A.NO_ENGINE, A.NO_CHASSIS, A.MDL, A.MODEL, A.FMDL, A.COLOR, " _
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
                                  & "D.TGL_D , D.CDD, D.USEFS_NM_HOBI, D.NOTE, E.FS_KD_PETUGAS, E.FS_NM_PETUGAS, E.FS_KD_PASS, E.FS_KD_GROUP, E.tipe, E.DIVISI, A.ANCUST, D.GSM, D.CDMA, A.TELPANCUST " _
                                  & "FROM JOBDTA AS A WITH (NOLOCK) INNER JOIN JOBDTL  AS B WITH (NOLOCK) ON A.NO_WO = B.NO_WO LEFT OUTER JOIN CUSTDAT AS D WITH (NOLOCK) ON A.NO_CHASSIS = D.NO_CHASSIS INNER JOIN " _
                                  & "T_PETUGAS AS E WITH (NOLOCK) ON A.FRONT = E.FS_KD_PETUGAS WHERE A.NO_WO = '" & Trim(text(64)) & "'")
            rpt.Database.SetDataSource rRs
                       
            gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
            
            sSQL = "SELECT * FROM historyUnitFunc('" & Trim(gFRM) & "')"
            Set sHlp = objO.FetchData(sSQL)
            Set crxSubReportHistory = rpt.Subreport2.OpenSubreport
            crxSubReportHistory.Database.SetDataSource sHlp
            
            sSQL = "select DATE_BOOK, TIME_D, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(78)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
            Set sHlp2 = objO.FetchData(sSQL)
            Set crxSubReportBooking = rpt.Subreport1.OpenSubreport
            crxSubReportBooking.Database.SetDataSource sHlp2
            
            sSQL = "select DATE, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(78)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
            Set sHlp4 = objO.FetchData(sSQL)
            Set crxSubReportBooking2 = rpt.Subreport4.OpenSubreport
            crxSubReportBooking2.Database.SetDataSource sHlp4
            
            sSQL = "select WAKTU FROM M_FUBOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(78)) & "' "
            Set sHlp3 = objO.FetchData(sSQL)
            Set crxSubReportBookingFU = rpt.Subreport3.OpenSubreport
            crxSubReportBookingFU.Database.SetDataSource sHlp3
            
            Set rsHlp = objO.FetchData("SELECT NOSTALL FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(text(64)) & "|%'") 'TANGGAL = '" & Format(Trim(text(65)), "MM/DD/YYYY") & "' AND
            If Not rsHlp.EOF Then
                rpt.txtStall.SetText Trim(rsHlp("NOSTALL"))
            End If
    
            rpt.mechPlan.SetText ""
            
            If bodyClear Then
                rpt.Text93.TextColor = &HFF&
                rpt.Text94.TextColor = &HFF&
                rpt.Text95.TextColor = &HFF&
                rpt.Text96.TextColor = &HFF&
                rpt.Text97.TextColor = &HFF&
            End If
            
            Dim n As Integer
            Set rsHlp = objO.FetchData("SELECT KOORDY,KOORDX FROM JOBKOORD WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "'")
            If Not rsHlp.EOF Then
                n = 1
                Do While Not rsHlp.EOF
                    Call pasangTanda(rpt, n, CDbl(rsHlp("KOORDY")), CDbl(rsHlp("KOORDX")))
                    n = n + 1
                    rsHlp.MoveNext
                Loop
            End If
            
            Set rsHlp = objO.FetchData("SELECT BBM FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "'")
            If Not rsHlp.EOF Then
                Select Case rsHlp("BBM")
                    Case 0
                        ' FULL
                        rpt.Text90.TextColor = &H0&
                    Case 1
                        ' 3/4
                        rpt.Text62.TextColor = &H0&
                    Case 2
                        ' 1/2
                        rpt.Text64.TextColor = &H0&
                    Case 3
                        ' 1/4
                        rpt.Text88.TextColor = &H0&
                    Case 4
                        ' EMPTY
                        rpt.Text89.TextColor = &H0&
                End Select
            End If
            
            With FrmctkOrder.CRViewer1
                .ReportSource = rpt
                .ViewReport
                .Zoom 100
                .Visible = True
            End With
            FrmctkOrder.cmdTutup.Visible = True
            Screen.MousePointer = vbDefault
            FrmctkOrder.WindowState = vbMaximized
            FrmctkOrder.Show vbModal
        Else
            Call CetakFormKecil
        End If
    End If
End Sub

Private Sub pasangTanda(rpt As Report, Urut As Integer, nilX As Integer, nilY As Integer)
Dim nTop As Integer
Dim nLeft As Integer
    
    nTop = 6750
    nLeft = 4300
    
    Select Case Urut
        Case 1
            rpt.imgTnd1.TextColor = &H0&
            rpt.imgTnd1.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd1.Left = nLeft + nilY - (210 / 2)
        Case 2
            rpt.imgTnd2.TextColor = &H0&
            rpt.imgTnd2.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd2.Left = nLeft + nilY - (210 / 2)
        Case 3
            rpt.imgTnd3.TextColor = &H0&
            rpt.imgTnd3.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd3.Left = nLeft + nilY - (210 / 2)
        Case 4
            rpt.imgTnd4.TextColor = &H0&
            rpt.imgTnd4.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd4.Left = nLeft + nilY - (210 / 2)
        Case 5
            rpt.imgTnd5.TextColor = &H0&
            rpt.imgTnd5.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd5.Left = nLeft + nilY - (210 / 2)
        Case 6
            rpt.imgTnd6.TextColor = &H0&
            rpt.imgTnd6.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd6.Left = nLeft + nilY - (210 / 2)
        Case 7
            rpt.imgTnd7.TextColor = &H0&
            rpt.imgTnd7.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd7.Left = nLeft + nilY - (210 / 2)
        Case 8
            rpt.imgTnd8.TextColor = &H0&
            rpt.imgTnd8.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd8.Left = nLeft + nilY - (210 / 2)
        Case 9
            rpt.imgTnd9.TextColor = &H0&
            rpt.imgTnd9.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd9.Left = nLeft + nilY - (210 / 2)
        Case 10
            rpt.imgTnd10.TextColor = &H0&
            rpt.imgTnd10.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd10.Left = nLeft + nilY - (210 / 2)
        Case 11
            rpt.imgTnd11.TextColor = &H0&
            rpt.imgTnd11.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd11.Left = nLeft + nilY - (210 / 2)
        Case 12
            rpt.imgTnd12.TextColor = &H0&
            rpt.imgTnd12.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd12.Left = nLeft + nilY - (210 / 2)
        Case 13
            rpt.imgTnd13.TextColor = &H0&
            rpt.imgTnd13.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd13.Left = nLeft + nilY - (210 / 2)
        Case 14
            rpt.imgTnd14.TextColor = &H0&
            rpt.imgTnd14.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd14.Left = nLeft + nilY - (210 / 2)
        Case 15
            rpt.imgTnd15.TextColor = &H0&
            rpt.imgTnd15.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd15.Left = nLeft + nilY - (210 / 2)
        Case 16
            rpt.imgTnd16.TextColor = &H0&
            rpt.imgTnd16.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd16.Left = nLeft + nilY - (210 / 2)
        Case 17
            rpt.imgTnd17.TextColor = &H0&
            rpt.imgTnd17.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd17.Left = nLeft + nilY - (210 / 2)
        Case 18
            rpt.imgTnd18.TextColor = &H0&
            rpt.imgTnd18.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd18.Left = nLeft + nilY - (210 / 2)
        Case 19
            rpt.imgTnd19.TextColor = &H0&
            rpt.imgTnd19.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd19.Left = nLeft + nilY - (210 / 2)
        Case 20
            rpt.imgTnd20.TextColor = &H0&
            rpt.imgTnd20.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd20.Left = nLeft + nilY - (210 / 2)
        Case 21
            rpt.imgTnd21.TextColor = &H0&
            rpt.imgTnd21.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd21.Left = nLeft + nilY - (210 / 2)
        Case 22
            rpt.imgTnd22.TextColor = &H0&
            rpt.imgTnd22.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd22.Left = nLeft + nilY - (210 / 2)
        Case 23
            rpt.imgTnd23.TextColor = &H0&
            rpt.imgTnd23.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd23.Left = nLeft + nilY - (210 / 2)
        Case 24
            rpt.imgTnd24.TextColor = &H0&
            rpt.imgTnd24.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd24.Left = nLeft + nilY - (210 / 2)
        Case 25
            rpt.imgTnd25.TextColor = &H0&
            rpt.imgTnd25.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd25.Left = nLeft + nilY - (210 / 2)
        Case 26
            rpt.imgTnd26.TextColor = &H0&
            rpt.imgTnd26.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd26.Left = nLeft + nilY - (210 / 2)
        Case 27
            rpt.imgTnd27.TextColor = &H0&
            rpt.imgTnd27.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd27.Left = nLeft + nilY - (210 / 2)
        Case 28
            rpt.imgTnd28.TextColor = &H0&
            rpt.imgTnd28.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd28.Left = nLeft + nilY - (210 / 2)
        Case 29
            rpt.imgTnd29.TextColor = &H0&
            rpt.imgTnd29.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd29.Left = nLeft + nilY - (210 / 2)
        Case 30
            rpt.imgTnd30.TextColor = &H0&
            rpt.imgTnd30.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd30.Left = nLeft + nilY - (210 / 2)
        Case 31
            rpt.imgTnd31.TextColor = &H0&
            rpt.imgTnd31.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd31.Left = nLeft + nilY - (210 / 2)
        Case 32
            rpt.imgTnd32.TextColor = &H0&
            rpt.imgTnd32.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd32.Left = nLeft + nilY - (210 / 2)
        Case 33
            rpt.imgTnd33.TextColor = &H0&
            rpt.imgTnd33.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd33.Left = nLeft + nilY - (210 / 2)
        Case 34
            rpt.imgTnd34.TextColor = &H0&
            rpt.imgTnd34.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd34.Left = nLeft + nilY - (210 / 2)
        Case 35
            rpt.imgTnd35.TextColor = &H0&
            rpt.imgTnd35.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd35.Left = nLeft + nilY - (210 / 2)
        Case 36
            rpt.imgTnd36.TextColor = &H0&
            rpt.imgTnd36.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd36.Left = nLeft + nilY - (210 / 2)
        Case 37
            rpt.imgTnd37.TextColor = &H0&
            rpt.imgTnd37.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd37.Left = nLeft + nilY - (210 / 2)
        Case 38
            rpt.imgTnd38.TextColor = &H0&
            rpt.imgTnd38.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd38.Left = nLeft + nilY - (210 / 2)
        Case 39
            rpt.imgTnd39.TextColor = &H0&
            rpt.imgTnd39.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd39.Left = nLeft + nilY - (210 / 2)
        Case 40
            rpt.imgTnd40.TextColor = &H0&
            rpt.imgTnd40.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd40.Left = nLeft + nilY - (210 / 2)
        Case 41
            rpt.imgTnd41.TextColor = &H0&
            rpt.imgTnd41.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd41.Left = nLeft + nilY - (210 / 2)
        Case 42
            rpt.imgTnd42.TextColor = &H0&
            rpt.imgTnd42.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd42.Left = nLeft + nilY - (210 / 2)
        Case 43
            rpt.imgTnd43.TextColor = &H0&
            rpt.imgTnd43.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd43.Left = nLeft + nilY - (210 / 2)
        Case 44
            rpt.imgTnd44.TextColor = &H0&
            rpt.imgTnd44.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd44.Left = nLeft + nilY - (210 / 2)
        Case 45
            rpt.imgTnd45.TextColor = &H0&
            rpt.imgTnd45.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd45.Left = nLeft + nilY - (210 / 2)
        Case 46
            rpt.imgTnd46.TextColor = &H0&
            rpt.imgTnd46.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd46.Left = nLeft + nilY - (210 / 2)
        Case 47
            rpt.imgTnd47.TextColor = &H0&
            rpt.imgTnd47.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd47.Left = nLeft + nilY - (210 / 2)
        Case 48
            rpt.imgTnd48.TextColor = &H0&
            rpt.imgTnd48.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd48.Left = nLeft + nilY - (210 / 2)
        Case 49
            rpt.imgTnd49.TextColor = &H0&
            rpt.imgTnd49.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd49.Left = nLeft + nilY - (210 / 2)
        Case 50
            rpt.imgTnd50.TextColor = &H0&
            rpt.imgTnd50.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd50.Left = nLeft + nilY - (210 / 2)
        Case 51
            rpt.imgTnd51.TextColor = &H0&
            rpt.imgTnd51.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd51.Left = nLeft + nilY - (210 / 2)
        Case 52
            rpt.imgTnd52.TextColor = &H0&
            rpt.imgTnd52.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd52.Left = nLeft + nilY - (210 / 2)
        Case 53
            rpt.imgTnd53.TextColor = &H0&
            rpt.imgTnd53.Top = nTop + nilX - (195 / 2)
            rpt.imgTnd53.Left = nLeft + nilY - (210 / 2)
    End Select
End Sub

Private Sub CetakFormKecil()
    Dim rpt As New RptOrderKecil
   
    rpt.Label2 = "NPWP : " & gNpwp
    rpt.Label3 = "PKP  : " & gPkp
    rpt.Label44 = "TELP : " & gTelp
    
    rpt.Label8 = "FRAME : " & gFRM
    rpt.Label9 = "E/G : " & gEG
    rpt.Label10 = "DEALER : " & gDealer
    
    rpt.Label11 = "NAMA : " & gCust
    rpt.Label12 = "ALAMAT : " & gAlmt
    rpt.Label13 = "KOTA : " & gKota & " (" & gTelpon & ")"
    rpt.Label14 = gMDLOrder
    rpt.Label15 = gWarna
    rpt.Label16 = "Tahun : " & gTahun
    rpt.Label52 = "Kode Plg   : " & gKodePlg
    rpt.Label4 = "Order No.  : " & gNoOrder
    rpt.Label5 = "Tanggal    : " & gTglMasuk
    rpt.Label6 = "Jam        : " & gJamMasuk
    rpt.Label7 = "No. Polisi : " & gNoPolisi
    
    rpt.Label18 = gKeluhan(1)
    rpt.Label19 = gKeluhan(2)
    rpt.Label20 = gKeluhan(3)
    rpt.Label21 = gKeluhan(4)
    rpt.Label22 = gKeluhan(5)
    rpt.Label27 = gOrder(1)
    rpt.Label28 = gOrder(2)
    rpt.Label29 = gOrder(3)
    rpt.Label30 = gOrder(4)
    rpt.Label31 = gOrder(5)
    rpt.Label32 = gOrder(6)
    rpt.Label33 = gOrder(7)
    
    rpt.Label46 = "Kilometer     : " & gKM
    rpt.Label47 = "Serv. Advisor : " & gSAdvisor
    rpt.Label40 = gSAdvisor
    
    rpt.Printer.Orientation = 1
    rpt.Printer.PaperSize = 256
    rpt.Printer.PaperWidth = 1440 * 8.5
    rpt.Printer.PaperHeight = 1440 * 5.5
    
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show
    End If
End Sub

Private Sub panggilWarranty(noRangka As String, modL As String)
Dim rsRec As ADODB.Recordset
Dim terlibat As Boolean

    Set rsTemp = objO.FetchData("SELECT A.ID_WARRANTY FROM ho_warranty A WITH (NOLOCK) INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
               & "WHERE MDL = '" & modL & "' AND no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
               & "ORDER BY A.id_warranty,mdl,no_Rangka1")
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            Set rsRec = objO.FetchData("SELECT NO_RANGKA FROM TEMPWAR WITH (NOLOCK) WHERE NO_RANGKA = '" & noRangka & "' AND ID_WARRANTY = '" & rsTemp("ID_WARRANTY") & "'")
            If Not rsRec.EOF Then
                terlibat = False
                Exit Do
            Else
                terlibat = True
            End If
            
            Set rsRec = objO.FetchData("SELECT no_chassis FROM History WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
            If Not rsRec.EOF Then
                terlibat = False
                Exit Do
            Else
                terlibat = True
            End If
            
            If Log.LogHOConn = 1 Then
                If ConHO = True Then
                    
                    Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
                    If Not rsRec.EOF Then
                        terlibat = False
                        Exit Do
                    Else
                        terlibat = True
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
        
        If terlibat Then
            MsgBox "UNIT TERLIBAT WARRANTY !!!", vbCritical, "UNIT SSC"
            cmdWarranty.BackColor = &HFF&
            '
            cmdWarranty.Enabled = True
            opTpSvc(2).Enabled = True
        End If
    Else
        cmdWarranty.BackColor = &HC0C0C0
        cmdWarranty.Enabled = False
    End If
End Sub

Private Sub panggilHistory(noRgk As String) ', noPols As String
    On Error Resume Next
    'On Error GoTo ErrHist
    Dim rsHist As ADODB.Recordset
        
    ' AMBIL HISTORY DARI HO
''    Dim rsHlp2 As ADODB.Recordset
''    Dim dBaseHO As String
''
''    dBaseHO = HODb
''    If Log.LogHOConn = 1 Then
''    If ConHO = True Then
''        Set rsHlp = objO.FetchDataHO("SELECT * FROM HISTORYNEW WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
''        If Not rsHlp.EOF Then
''            Do While Not rsHlp.EOF
''                If rsHlp("NO_CHASSIS") <> "" Then
''                    ' INPUT HISTORY
''                    Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "' AND KM= '" & rsHlp("KM") & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC, NO_WO DESC")
''                    If rsHlp2.EOF Then
''                        Dim nSQL As String
''
''                        nSQL = "INSERT INTO HISTORY VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & Format(rsHlp("DATE"), "mm/dd/yyyy") & "'," & rsHlp("PART_G") & "," & rsHlp("PART_NG") & "," & rsHlp("OIL_G") & "," & rsHlp("OIL_NG") & "," & rsHlp("MTRL") & "," & rsHlp("LABOR") & "," & rsHlp("BODY") & "," & rsHlp("SBLT") & "," & rsHlp("MET") & "," & rsHlp("TOTAL") & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "'," & _
''                               "'" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "','" & rsHlp("MEKANIK4") & "','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "'," & _
''                               "'" & rsHlp("KLPKERJA7") & "','" & rsHlp("KLPKERJA8") & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & rsHlp("FN_RETURN") & "','" & rsHlp("FS_RC") & "','" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "','" & rsHlp("FS_KD_WIL") & "','" & rsHlp("THS") & "','" & rsHlp("MOU") & "','" & rsHlp("PDS") & "','" & rsHlp("FS_LOG_DEL") & "','" & rsHlp("FD_TGL_DEL") & "', " & _
''                               "'" & rsHlp("ALASAN_DEL") & "', '','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')"
''
''                        Call objO.InsertData(nSQL)
''                    End If
''
''                    ' INPUT PART
''                    Dim stringAwal As String
''                    Dim stringPotong1 As String
''                    Dim stringPotong2 As String
''                    Dim nNext As Integer
''                    Dim hasil(15) As String
''                    'If rsHlp("MEMOPART") <> "" Then
''                        Set rsHlp2 = objO.FetchDataHO("SELECT * FROM T_PARTnew WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "'")
''                        If rsHlp2.EOF Then
''                            stringAwal = Trim(rsHlp("MEMOPART"))
''                            Do While Len(stringAwal) > 0
''                                stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
''                                nNext = 0
''                                Do While Len(stringPotong1) > 0
''                                    stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
''                                    hasil(nNext) = Trim(stringPotong2)
''                                    nNext = nNext + 1
''                                    stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
''                                Loop
''                                stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
''
''                                ' SIMPAN
''                                Call objO.InsertData("INSERT INTO T_PART VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "'," & Trim(hasil(3)) & "," & setKoma(Trim(hasil(4))) & ", " & IIf(Trim(hasil(5)), 1, 0) & ",'" & Trim(hasil(6)) & "', " _
''                                             & " " & setKoma(Trim(hasil(7))) & ",'" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "')")
''                            Loop
''                        End If
''                    'End If
''
''                    ' INPUT SUBLET
''                    If rsHlp("MEMOSUBLET") <> "" Then
''                        Set rsHlp2 = objO.FetchData("SELECT * FROM HIST_SUBLET WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "' AND NO_CHASSIS = '" & noRgk & "'")
''                        If rsHlp2.EOF Then
''                            stringAwal = Trim(rsHlp("MEMOSUBLET"))
''                            Do While Len(stringAwal) > 0
''                                stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
''                                nNext = 0
''                                Do While Len(stringPotong1) > 0
''                                    stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
''                                    hasil(nNext) = Trim(stringPotong2)
''                                    nNext = nNext + 1
''                                    stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
''                                Loop
''                                stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
''
''                                ' SIMPAN
''                                Call objO.InsertData("INSERT INTO HIST_SUBLET VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & Trim(hasil(5)) & "','" & Trim(hasil(6)) & "', " _
''                                             & "'" & Trim(hasil(7)) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "','" & Trim(hasil(10)) & "','" & Trim(hasil(11)) & "','" & Trim(hasil(12)) & "','" & Trim(hasil(13)) & "','" & Trim(hasil(14)) & "')")
''                            Loop
''                        End If
''                    End If
''                End If
''                rsHlp.MoveNext
''            Loop
''            If rsHlp.State <> adStateClosed Then rsHlp.MoveLast
''            Text(17) = Format(Trim(rsHlp("DATE")), "dd-mm-yyyy")
''            'Text(18) = Format(Trim(rsHlp("RCDATE")), "dd-mm-yyyy")
''            Text(19) = Nul(Trim(rsHlp("FS_NM_DEALER")))
''            Text(20) = Format(CDbl(rsHlp("KM")), "0#,#")
''            cmdHistory.Enabled = True
''        End If
''        rsHlp.Close
''    End If
''    End If
    
    On Error GoTo vbError
    
    'Set rsHist = objO.FetchData("SELECT a.*, C.INVDATE, rcdate FROM History a INNER JOIN custdat b ON a.no_chassis = b.no_chassis AND A.FS_LOG_DEL = '' LEFT JOIN T_INVOICE C ON A.NO_WO = C.NO_WO AND A.NOINV = C.NOINV AND A.NO_CHASSIS = C.NO_CHASSIS AND C.FS_LOG_DEL = '' WHERE a.no_chassis = '" & Trim(noRgk) & "' ORDER BY A.date DESC, A.NO_WO DESC")
    Set rsHist = objO.FetchData("SELECT DISTINCT a.*, C.INVDATE, rcdate FROM History a WITH (NOLOCK) INNER JOIN custdat b WITH (NOLOCK) ON a.no_chassis = b.no_chassis AND A.FS_LOG_DEL = '' LEFT JOIN T_INVOICE C WITH (NOLOCK) ON A.NO_WO = C.NO_WO AND A.NOINV = REPLACE(REPLACE(REPLACE(C.NOINV,'-',''),'B',''),'A','') AND A.NO_CHASSIS = C.NO_CHASSIS AND C.FS_LOG_DEL = '' WHERE a.no_chassis = '" & Trim(noRgk) & "' ORDER BY A.date DESC, A.NO_WO DESC")
    
    If Not rsHist.EOF Then
        text(17) = Format(Trim(rsHist("DATE")), "dd-mm-yyyy")
        text(18) = Format(Trim(rsHist("RCDATE")), "dd-mm-yyyy")
        text(19) = Nul(Trim(rsHist("FS_NM_DEALER")))
        text(20) = Format(CDbl(IIf(rsHist("KM") = "", 0, rsHist("KM"))), "#,#")
        If Trim(text(20)) = "" Then text(20) = 0
        
        ChkReturnOrder.Enabled = False
        ChkReturnOrder.value = Unchecked
        If NulTgl(rsHist("INVDATE")) <> "01-01-1900" Then
            If DateDiff("d", Format(rsHist("INVDATE"), "dd/mm/yyyy"), Format(text(65), "dd/mm/yyyy")) <= 15 Then
                ChkReturnOrder.Enabled = True
            End If
        End If
        cmdHistory.Enabled = True
    Else
        ChkReturnOrder.Enabled = False
        ChkReturnOrder.value = Unchecked
        cmdHistory.Enabled = True
    End If
vbError:
    Exit Sub
ErrHist:
    Exit Sub
    'MsgBox err.Description
End Sub


Private Sub aturGridJob()
    baris = 1
    DGrid.Rows = 12
    DGrid.cols = 7
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 6000
    DGrid.ColWidth(2) = 800
    DGrid.ColWidth(3) = 0
    DGrid.ColWidth(4) = 0
    DGrid.ColWidth(5) = 0
    DGrid.ColWidth(6) = 0
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.TextMatrix(0, 1) = "Keluhan & Order Booking"
    
    ' WAKTU
    Dim gTglTampung As Date
    gTglTampung = gTanggal
    text(65) = TglDDMMYYYY(gTglTampung)
    'text(66) = Format(gTglTampung, "HH:NN")
    text(66) = Format(gTglServer, "HH:NN")
    text(62) = text(65) 'TglDDMMYYYY(gTglTampung)
    text(63) = text(66) 'Format(gTglTampung, "HH:NN")
    text(88) = text(65) 'TglDDMMYYYY(gTglTampung)
    text(89) = text(66) 'Format(gTglTampung, "HH:NN")
    text(86) = text(66) 'Format(gTglTampung, "HH:NN")
End Sub

Private Sub refrPart(noFkt As String)
    Dim rsHlp As ADODB.Recordset
    If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
        Set rsHlp = objO.FetchDataACC("select KodeItem, Keterangan as namaBarang, Qty, '' as KDOLI, BO as KodeSatuan, NamaBYSA, Nomor as nID from SOORD WITH (NOLOCK) Where NoSPK = '" & noFkt & "' and BO <> 'CANCEL'")
    'End If
   ' Set rsHlp = objO.FetchDataACC("select KodeItem, Keterangan as namaBarang, Qty, '' as KDOLI, BO as KodeSatuan, NamaBYSA, Nomor as nID from SOORD Where NoSPK = '" & noFkt & "' and BO <> 'CANCEL'")
    'If rsHlp.RecordCount = 0 Then
    Else
        Set rsHlp = objO.FetchData("SELECT kodeItem, namaBarang, qty, ISNULL(FS_KD_OLI,'') AS KDOLI, KODESATUAN, NAMABYSA, nID FROM book_part A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & noFkt & "' ORDER BY nID ASC")
    End If
    If Not rsHlp.EOF Then
        If createTblPart Then
            Do While Not rsHlp.EOF
                rsGridPart.AddNew
                rsGridPart!KodeItem = rsHlp("KODEITEM")
                rsGridPart!NAMAITEM = rsHlp("NAMABARANG")
                rsGridPart!SATITEM = rsHlp("KODESATUAN")
                rsGridPart!qty = CDbl(rsHlp("QTY"))
                rsGridPart!oli = IIf(rsHlp("KDOLI") <> "", True, False)
                rsGridPart!ganti = False
                rsGridPart!BAWA = False
                rsGridPart!namabySA = rsHlp("NAMABYSA")
                rsGridPart!nID = Val(NulA(rsHlp("nID")))
                rsGridPart.Update
                NewId = Val(NulA(rsHlp("nID")))
                rsHlp.MoveNext
            Loop
        End If
        Set partGrid.DataSource = rsGridPart
        Set partGridFill.DataSource = rsGridPart
        Call HeadColumn
        NewId = NewId + 1
    Else
        Set rsGridPart = Nothing
    End If
End Sub

Public Sub isiPart(iKdItem As String, iNmItem As String, iSatItem As String, iQty As Integer, iOli As Boolean, iNmBySA As String, iID As Double)
    If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
        If rsGridPart.State = adStateClosed Then createTblPart
        
        If Not rsGridPart.BOF Then rsGridPart.MoveFirst
        If iNmBySA <> "" Then
            rsGridPart.Find "NAMABYSA = '" & Trim(iNmBySA) & "'", adSearchForward, 1
        Else
            rsGridPart.Find "KODEITEM = '" & Trim(iKdItem) & "'", adSearchForward, 1
        End If
        If rsGridPart.EOF Then
            rsGridPart.AddNew
            rsGridPart!KodeItem = iKdItem
            rsGridPart!NAMAITEM = iNmItem
            rsGridPart!SATITEM = iSatItem
            rsGridPart!qty = CDbl(iQty)
            rsGridPart!oli = iOli
            rsGridPart!namabySA = iNmBySA
            rsGridPart!nID = iID
            rsGridPart.Update
        End If
        Set partGrid.DataSource = rsGridPart
        Set partGridFill.DataSource = rsGridPart
        Call HeadColumn
    Else
        If rsGridPart.State = adStateClosed Then createTblPart
        
        If Not rsGridPart.BOF Then rsGridPart.MoveFirst
        rsGridPart.Find "KODEITEM = '" & Trim(iKdItem) & "'", adSearchForward, 1
        If rsGridPart.EOF Then
            rsGridPart.AddNew
            rsGridPart!KodeItem = iKdItem
            rsGridPart!NAMAITEM = iNmItem
            rsGridPart!SATITEM = iSatItem
            rsGridPart!qty = CDbl(iQty)
            rsGridPart!oli = iOli
            rsGridPart!namabySA = iNmBySA
            rsGridPart!nID = iID
            rsGridPart.Update
        End If
        Set partGrid.DataSource = rsGridPart
        Call HeadColumn
    End If
    NewId = NewId + 1
    
    If rsGridPart.State <> adStateClosed Then
        text(59) = 0
        text(57) = 0
        If Not rsGridPart.EOF Then
        rsGridPart.MoveFirst
        Do While Not rsGridPart.EOF
            Set rsOli = objO.FetchDataACC("SELECT hargaJualA FROM inmst WITH (NOLOCK) WHERE kodeItem = '" & rsGridPart("KODEITEM") & "'")
            If Not rsOli.EOF Then
                If rsGridPart("OLI") Then
                    text(59) = text(59) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                Else
                    text(57) = text(57) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                End If
            End If
            rsGridPart.MoveNext
        Loop
        End If
    End If

End Sub

Function createTblPart() As Boolean
    On Error GoTo Vb_Error
    
    If rsGridPart.State Then rsGridPart.Close
    With rsGridPart
        .Fields.Append "KODEITEM", adVarChar, 30
        .Fields.Append "NAMAITEM", adVarChar, 200
        .Fields.Append "SATITEM", adVarChar, 30
        .Fields.Append "QTY", adInteger
        .Fields.Append "GANTI", adBoolean
        .Fields.Append "BAWA", adBoolean
        .Fields.Append "OLI", adBoolean
        .Fields.Append "NAMABYSA", adVarChar, 100
        .Fields.Append "nID", adDouble
        .Open
    End With
    If rsGridPart.State Then createTblPart = True
    Exit Function
    
Vb_Error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub HeadColumn()
    partGrid.Columns(0).Caption = "KODE PART"
    partGrid.Columns(0).DataField = "KODEITEM"
    partGrid.Columns(0).Width = 1300
    partGrid.Columns(1).Caption = "NAMA PART"
    partGrid.Columns(1).DataField = "NAMAITEM"
    partGrid.Columns(1).Width = 2500
    partGrid.Columns(2).Caption = "SATUAN PART"
    partGrid.Columns(2).DataField = "SATITEM"
    partGrid.Columns(2).Width = 800
    partGrid.Columns(3).Caption = "QTY"
    partGrid.Columns(3).DataField = "QTY"
    partGrid.Columns(3).Width = 600
    partGrid.Columns(4).Caption = "RPL"
    partGrid.Columns(4).DataField = "GANTI"
    partGrid.Columns(4).Alignment = dbgCenter
    partGrid.Columns(4).Width = 500
    partGrid.Columns(5).Caption = "PART SENDIRI"
    partGrid.Columns(5).DataField = "BAWA"
    partGrid.Columns(5).Alignment = dbgCenter
    partGrid.Columns(5).Width = 1300
    partGrid.Columns(6).Caption = "OLI"
    partGrid.Columns(6).DataField = "OLI"
    partGrid.Columns(6).Visible = False
    'partGrid.Columns(7).Caption = "NAMA BY SA"
    'partGrid.Columns(7).DataField = "NAMABYSA"
    'partGrid.Columns(7).Width = 2500
End Sub

Private Sub TampilWaktu(Tggl As Date)
Dim rsWaktu As ADODB.Recordset
Dim nStall As String
Dim i As Integer

    GridWaktu.TextMatrix(0, 0) = "..."
    GridWaktu.TextMatrix(0, 1) = "STALL"
    
    'FUNGSI MENGHASILKAN WAKTU TERSEDIA BERDASARKAN TGL TTN (HARI H)
    If DateDiff("d", isnow, Tggl) <= 0 Then
        Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 2, "mm/dd/yyyy") & "'")
        GridWaktu.TextMatrix(0, 2) = "  H  "
        GridWaktu.TextMatrix(0, 3) = "H + 1"
        GridWaktu.TextMatrix(0, 4) = "H + 2"
        GridWaktu.TextMatrix(0, 5) = "H + 3"
    End If
    
    'FUNSGI UNTUK MENAMPILKAN WAKTU TERSEDIA H-2 SAMPAI H+1
    Set rsWaktu = objO.FetchData("SELECT NOSTALL FROM STALLWKF WITH (NOLOCK) ORDER BY LEFT(NOSTALL,1),CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
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

Private Function PanggilDataBooking(noFkt As String) As String
    If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
        Dim rsDel As ADODB.Recordset
        Set rsDel = objO.FetchData("Select NoFAK FROm DelBook WITH (NOLOCK) Where Hapus = 0 and NoFak = '" & noFkt & "'")
        If rsDel.RecordCount > 0 Then
            MsgBox "Data Booking " & noFkt & ", sudah dalam Log Auto Delete !" & Chr(13) & "Input data by No Rangka", vbInformation + vbOKOnly, "WO dari Booking"
            PanggilDataBooking = ""
            Exit Function
        End If
    End If
    
    Dim xDate As Date
    xDate = isnow
    Set rsTemp = objO.FetchData("SELECT distinct a.*, b.no_engine, b.keyno, b.no_chassis, b.fs_kd_dealer, b.fs_nm_dealer, ISNULL(b.ddate,'01/01/1900') AS DDATE, ISNULL(b.dotam,'01/01/1900') AS DOTAM, A.fs_kd_plg, A.fs_kd_wil, b.heavyuse " _
                             & " FROM book a WITH (NOLOCK) LEFT JOIN custdat b WITH (NOLOCK) ON (a.no_chassis = b.no_chassis OR a.no_polisi = b.no_polisi) AND b.no_chassis <> '' WHERE noFak = '" & noFkt & "'")   ' AND deliv = 'BOK'
    If Not rsTemp.EOF Then
        If rsTemp("DELIV") = "BOK" Or Left(rsTemp("DELIV"), 3) = "RES" Then
            SSTab1.Enabled = True
            text(1) = Nul(rsTemp("FS_KD_PLG"))
            text(2) = Nul(rsTemp("FS_KD_WIL"))
            text(3) = Nul(rsTemp("CUSTOMER"))
            If Len(Trim(Nul(rsTemp("CUST_ADRES")))) > 60 Then
                text(4) = Left(Nul(rsTemp("CUST_ADRES")), 60)
                text(117) = Mid(Nul(rsTemp("CUST_ADRES")), 61, 60)
            Else
                text(4) = Left(Nul(rsTemp("CUST_ADRES")), 60)
            End If
            text(5) = Nul(rsTemp("KOTA"))
            text(6) = Nul(rsTemp("TELPON"))
            text(7) = Nul(rsTemp("GSM"))
            text(8) = Nul(rsTemp("CDMA"))
            text(9) = Nul(rsTemp("FS_TH_MOBIL"))
            text(10) = rsTemp("MDL")
            text(11) = rsTemp("MODEL")
            text(12) = rsTemp("FMDL")
            text(13) = rsTemp("COLOR")
            text(14).Mask = ""
            text(14) = Nul(rsTemp("NO_POLISI")) 'Replace(rsTemp("NO_POLISI"), "-", "")
            text(15) = Nul(rsTemp("NO_CHASSIS"))
            text(16).Mask = ""
            text(16) = TglDDMMYYYY(rsTemp("DOTAM"))
            text(16).Mask = "99-99-9999"
            text(118).Mask = ""
            text(118) = TglDDMMYYYY(rsTemp("DDATE"))
            text(118).Mask = "99-99-9999"
            text(90) = cekMOU(rsTemp("MDL"), Nul(rsTemp("NO_CHASSIS")))
            cmdLengkap.Visible = False
            If Trim(text(90)) <> "" Then
                cmdLengkap.Visible = True
            End If
        
            text(79).Mask = ""
            text(79) = TglDDMMYYYY(rsTemp("DDATE"))
            text(118).Mask = ""
            text(118) = TglDDMMYYYY(rsTemp("DDATE"))
            text(118).Mask = "99-99-9999"
            text(81) = Nul(rsTemp("FS_KD_DEALER"))
            text(80) = Nul(rsTemp("FS_NM_DEALER"))
            text(76) = Nul(rsTemp("NO_ENGINE"))
            text(23) = rsTemp("REQ1")
            text(24) = rsTemp("REQ2")
            text(25) = rsTemp("REQ3")
            text(26) = rsTemp("REQ4")
            text(27) = rsTemp("REQ5")
            
            text(88).Mask = ""
            text(88) = TglDDMMYYYY(rsTemp("DATESERAH"))
            text(88).Mask = "99-99-9999"
            
            text(89).Mask = ""
            text(89) = Trim(rsTemp("TIMESERAH"))
            text(89).Mask = "99:99"
            
            ' ISI KELUHAN
            Dim hlp As String
            For i = 1 To 5
                hlp = "CMPL" & i
                'text(96 + i) = rsTemp(hlp)
                Text1(-1 + i) = rsTemp(hlp)
            Next
            
            text(50) = rsTemp("MECH")
            Txt(54) = rsTemp("REMARK1")
            Txt(55) = rsTemp("SUGGEST")
            text(56) = Format(CDbl(rsTemp("TOTLABOUR")), "0#,#")
'            Text(57) = Format(CDbl(rsTemp("TOTPART")), "0#,#")
            text(62) = TglDDMMYYYY(xDate)
            text(63).Mask = ""
            text(63) = rsTemp("ESTTIME")
            text(86) = rsTemp("ESTTIME")
            SgnRekom.Visible = IIf(rsTemp("REKOM") = "Y", True, False)
            chTHS = IIf(rsTemp("THS") = 0, Unchecked, Checked)
            ChkReturnOrder = IIf(rsTemp("FN_RETURN") = 0, Unchecked, Checked)
            ChkReturnOrder.Enabled = True
            chEM = IIf(rsTemp("CODE") = 1, Checked, Unchecked)
            chSBNP = IIf(rsTemp("KD_K") = 1, Checked, Unchecked)
            
            Txt(53).text = rsTemp("Catatan1")
            Txt(87).text = rsTemp("Catatan2")
            
            ' ISI PEKERJAAN
            Dim rsHlp As ADODB.Recordset
            Dim hlp2 As String
            Dim nMech As String
            baris = 1
            nMech = Trim(rsTemp("MECH"))
            For i = 1 To 8
                hlp = "REP" & i
                hlp2 = "RTE" & i
ulang:
                If Trim(rsTemp(hlp)) <> "" Then
                    '
                    Set rsHlp = objO.FetchData("SELECT * FROM T_MEKANIK WITH (NOLOCK) WHERE TANGGAL = '" & Format(xDate, "DD/MM/YYYY") & "' AND MECH = '" & Trim(nMech) & "' AND JAMMULAI <= '" & Format(isnow, "HH:NN") & "' AND JAMSELESAI > '" & Format(xDate, "HH:NN") & "'")
                    If Not rsHlp.EOF Then
                        DGrid.TextMatrix(i, 1) = Trim(rsTemp(hlp))
                        DGrid.TextMatrix(i, 2) = CDbl(rsTemp(hlp2))
                        DGrid.TextMatrix(i, 6) = Trim(nMech)
                        text(50) = Trim(nMech)
                    Else
                        MsgBox "Mekanik " & nMech & " Tidak tersedia. " & vbCr & "Jam Mekanik belum diset. "
REPEET:
                        nMech = UCase(InputBox("Mekanik Pengganti : ", "Penggantian Mekanik"))
                        If nMech = "" Then
                            If MsgBox("MEMBATALKAN ORDER? ", vbYesNo, "PEMBATALAN ORDER") = vbYes Then
                                Call batalMekanik(Trim(text(64)))
                                Call hapusForm
                                Exit Function
                            Else
                                MsgBox "Isi inisial mekanik !!"
                                GoTo REPEET
                            End If
                        Else
                            GoTo ulang
                        End If
                    End If
                End If
            Next
            
            chUse = IIf(rsTemp("HEAVYUSE") = 1, Checked, Unchecked)
            
            ' TAMPIL CONDITION
            
            Dim kCond As String
            Dim kBant As String
                For i = 1 To 6
                    hlp = "COND" & i
                    If rsTemp(hlp) <> "" Then
                        If Left(rsTemp(hlp), 3) = "L| " Then
                            kBant = Mid(rsTemp(hlp), 4, Len(rsTemp(hlp)) - 3)
                            kCond = "LAINNYA"
                        Else
                            kBant = ""
                            kCond = Trim(rsTemp(hlp))
                        End If
                    Else
                        kBant = ""
                        kCond = ""
                    End If
                    text(27 + i) = kCond
                    bKondisi(i - 1) = kBant
                Next
            
            ' ISI JOB ORDER
            myChassis = Trim(text(15))
            ntRate = 0
            text(114) = 0
            For j = 1 To 7
                If Trim(DGrid.TextMatrix(j, 1)) <> "" Then
                    Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                    If Not rsHlp.EOF Then
                        DGrid.TextMatrix(j, 2) = rsHlp("RATE")
                        DGrid.TextMatrix(j, 3) = False
                        DGrid.TextMatrix(j, 4) = rsHlp("JOB")
                        DGrid.TextMatrix(j, 5) = hitungLabour(rsHlp("RATE"), text(11))
                        If isJualNRM = True Then
                            Me.JualNRM.Caption = nJualNRM
                        Else
                            Me.JualNRM.Caption = ""
                        End If
                    Else
                        DGrid.TextMatrix(j, 3) = True
                        DGrid.TextMatrix(j, 5) = 0
                    End If
                    
                    baris = baris + 1
                End If
                ntRate = ntRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(DGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                text(114) = ntRate
            Next
            
            ' ISI SUBLET
            Call objO.InsertData("INSERT INTO TEMPJOBSUBLET (NO_CHASSIS, NOFAK, NAMABARANG, KODESATUAN, QTY, FLAG, NO_POLISI) SELECT '" & Trim(text(15)) & "', NOFAK, NAMABARANG, KODESATUAN, QTY, FLAG, '" & Trim(text(14)) & "' FROM BOOK_SUBLET WITH (NOLOCK) WHERE NOFAK = '" & Trim(noFkt) & "' ")
            
            ' CHECK STALL
            Set rsTmp = objO.FetchData("SELECT * FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & noFkt & "|%'")
            If Not rsTmp.EOF Then
                cmdStall.Tag = ""
                sgnStall.Visible = True
                sgnStall.Caption = Trim(rsTmp("NOSTALL")) '& " : " & Trim(rsTmp("JAMMULAI")) & " s/d " & Trim(rsTmp("JAMSELESAI"))
                    
                Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN like '%|" & noFkt & "|%' AND Tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'")
                Call objO.DeleteData("INSERT INTO tempStall SELECT noStall, '" & Format(xDate, "MM/DD/YYYY") & "', JamMulai, JamSelesai, rate, [Group], Mech, Flag, Keterangan,adjust FROM STALL WITH (NOLOCK) " _
                             & "WHERE KETERANGAN like '%|" & noFkt & "|%' AND NOT EXISTS " _
                             & "(SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE KETERANGAN like '%|" & noFkt & "|%' AND Tanggal = '" & Format(text(65), "mm/dd/yyyy") & "')")
                             
                nWKT = hitungWaktu(ntRate, IIf(Trim(rsTmp("JAMMULAI")) > Format(xDate, "HH:NN"), Trim(rsTmp("JAMMULAI")), Format(xDate, "HH:NN")), Format(Trim(text(65)), "MM/DD/YYYY"))
                text(86) = IIf(Trim(rsTmp("JAMMULAI")) > Format(xDate, "HH:NN"), Trim(rsTmp("JAMMULAI")), Format(xDate, "HH:NN"))
                text(63) = Left(nWKT, InStr(nWKT, "|") - 1)
                text(62) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
            Else
                cmdStall.Tag = "belum"
            End If
            
            ' ISI PART
            Call refrPart(noFkt)
            
            If Trim(rsTemp("FLAGSVC")) = "TINGGAL" Then
                opSvc(0).value = True
            ElseIf Trim(rsTemp("FLAGSVC")) = "TUNGGU" Then
                opSvc(1).value = True
            End If
                       
            text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
        Else
            text(78) = ""
            If rsTemp("DELIV") = "OKE" Then
                If Trim(text(15)) <> "" Then
                    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(15)) & "' AND DELIV <> 'BTL' AND RPR_COND <> 'K' ORDER BY [DATE] DESC")
                    If Not rsHlp.EOF Then
                        MsgBox "Transaksi Lama Belum Closing !!"
                        hlp = Trim(text(0))
                        Call hapusForm
                        text(0) = hlp
                        Exit Function
                    End If
                End If
            End If
            SSTab1.Enabled = False
        End If
        simpan = False
        'lakukan proses cek prediagnose
                
        'rsTemp = Nothing
        Dim rsTempDiagnose As ADODB.Recordset
        
        Set rsTempDiagnose = objO.FetchData("SELECT * FROM diagapproval WITH (NOLOCK) WHERE NOFAK = '" & noFkt & "'")
        If Not rsTempDiagnose.EOF Then
            'jika booking tersebut termasuk dalam diagnose
            opTpSvc(2).Enabled = True
            opTpSvc(2).value = True
            text(22) = rsTempDiagnose.Fields(0)
            opTpSvc(0).Enabled = False
            opTpSvc(1).Enabled = False
            
        End If
        Call hitungEstimasiPenyerahan
        PanggilDataBooking = Nul(rsTemp("NO_CHASSIS"))
    End If
End Function

Private Function PanggilDataCustomer(nRgk As String, noWO As String) As Boolean
Dim Ket As String
    Ket = ""
    If Trim(noWO) <> "" Then
        Ket = " AND B.NO_WO <= '" & Trim(noWO) & "' "
    End If
    Set rsTemp = objO.FetchData("SELECT a.*, rpr_cond, TPSVC FROM custdat a WITH (NOLOCK) LEFT JOIN jobdta b WITH (NOLOCK) ON a.no_chassis = b.no_chassis AND b.deliv <> 'BTL' " _
                             & "WHERE a.no_chassis = '" & nRgk & "' " & Ket & " ORDER BY no_wo DESC")
    If Not rsTemp.EOF Then
        If cmdPilih.Tag = "No. Rangka" Then
            text(0) = Trim(nRgk)
        End If
        Select Case Trim(text(1))
            Case ""
                text(1) = rsTemp("FS_KD_PLG")
                text(2) = rsTemp("FS_KD_WIL")
                text(3) = rsTemp("CUSTOMER")
                Select Case Len(Trim(Nul(rsTemp("CUST_ADRES"))))
                    Case Is > 60
                        text(4) = Left(rsTemp("CUST_ADRES"), 60)
                        text(117) = Mid(rsTemp("CUST_ADRES"), 61, 60)
                    Case Else
                        text(4) = Left(rsTemp("CUST_ADRES"), 60)
                End Select
                text(5) = rsTemp("KOTA")
                text(6) = Replace(rsTemp("TELPON"), " ", "")
                text(7) = rsTemp("GSM")
                text(8) = rsTemp("CDMA")
            Case Trim(rsTemp("FS_KD_PLG"))
                Select Case Len(Trim(Nul(rsTemp("CUST_ADRES"))))
                    Case Is > 60
                        text(4) = Left(rsTemp("CUST_ADRES"), 60)
                        text(117) = Mid(rsTemp("CUST_ADRES"), 61, 60)
                    Case Else
                        text(4) = Left(rsTemp("CUST_ADRES"), 60)
                End Select
        End Select
        text(9) = rsTemp("FS_TH_MOBIL")
        text(10) = rsTemp("MDL")
        text(11) = rsTemp("MODEL")
                Call CexChassis_Discon(text(11))
                If isJualNRM = True Then
                    If tglJualNRM = "01/01/1900" Then
                        tglJualNRM = FormatDDMMYYYYtoMMDDYYYY(rsTemp("DDATE"))
                    End If
                Else
                    tglJualNRM = FormatDDMMYYYYtoMMDDYYYY(rsTemp("DDATE"))
                End If
                If isOnCarClass = False Then
                    MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                End If
                
                If DateAdd("m", Log.warranty, tglJualNRM) >= Format(isnow, "MM/DD/YYYY") Then
                    MsgBox "Unit Masih dalam Masa Warranty", vbInformation, "Warranty Nasmoco Club"
                Else
                    MsgBox "Unit sudah tidak dalam Masa Warranty", vbInformation, "Warranty Nasmoco Club"
                End If
        text(12) = rsTemp("FMDL")
        text(13) = rsTemp("COLOR")
        text(14).Mask = ""
        text(14) = rsTemp("NO_POLISI")
        text(15) = rsTemp("NO_CHASSIS")
        text(16).Mask = ""
        text(16) = TglDDMMYYYY(rsTemp("DOTAM"))
        text(16).Mask = "99-99-9999"
        text(20) = Format(CDbl(rsTemp("FS_KM_AKHIR")), "#,#")
        If Trim(text(20)) = "" Then text(20) = 0
        chUse = IIf(rsTemp("HEAVYUSE") = 1, Checked, Unchecked)
        text(76) = rsTemp("NO_ENGINE")
        text(79).Mask = ""
        text(79) = rsTemp("DDATE")
        text(118).Mask = ""
        text(118) = TglNul(rsTemp("DDATE"))
        text(118).Mask = "99-99-9999"
        
        text(81) = rsTemp("FS_KD_DEALER")
        text(80) = rsTemp("FS_NM_DEALER")
        text(90) = cekMOU(rsTemp("MDL"), rsTemp("NO_CHASSIS"))
        cmdLengkap.Visible = False
        If Trim(text(90)) <> "" Then
            cmdLengkap.Visible = True
        End If
        
        If rsTemp("RPR_COND") = "K" Or IsNull(rsTemp("RPR_COND")) Then
            opTpSvc(0).Enabled = True
            opTpSvc(1).Enabled = True
            opTpSvc(2).Enabled = True
        Else
            If cmdPilih.Tag <> "No. WO" Then
                If rsTemp("TPSVC") <> "W" Then
                    If opTpSvc(0).Enabled Or Not SSTab1.Enabled Then
                        If adaBelumClosing(Trim(rsTemp("NO_CHASSIS"))) Then
                            If Trim(cekMOU(Trim(rsTemp("MDL")), Trim(rsTemp("NO_CHASSIS")))) = "" Then
                                If cekWoInvoice(Trim(rsTemp("NO_CHASSIS"))) Then
                                    MsgBox "Transaksi lama belum Invoice "
                                    opTpSvc(0).Enabled = False
                                    opTpSvc(1).Enabled = False
                                    opTpSvc(2).Enabled = True
                                    opTpSvc(2).value = True
                                Else
                                    If Not cekLastIsSBIDealer(Trim(rsTemp("NO_CHASSIS"))) Then
                                        If Not cekLastIsPDS(Trim(rsTemp("NO_CHASSIS"))) Then
                                            MsgBox "Transaksi lama belum Closing "
                                            opTpSvc(0).Enabled = False
                                            opTpSvc(1).Enabled = False
                                            opTpSvc(2).Enabled = True
                                            opTpSvc(2).value = True
                                        Else
                                            If Not cekSBIDealerInvoice(Trim(rsTemp("NO_CHASSIS"))) Then
                                                MsgBox "Transaksi lama belum Closing "
                                                opTpSvc(0).Enabled = False
                                                opTpSvc(1).Enabled = False
                                                opTpSvc(2).Enabled = True
                                                opTpSvc(2).value = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If rsTemp("TPSVC") = "W" Then
                    opTpSvc(2).Enabled = True
                    If adaBelumClosing(Trim(rsTemp("NO_CHASSIS"))) Then
                        If Trim(cekMOU(Trim(rsTemp("MDL")), Trim(rsTemp("NO_CHASSIS")))) = "" Then
                            If cekWoInvoice(Trim(rsTemp("NO_CHASSIS"))) Then
                                opTpSvc(0).Enabled = False
                                opTpSvc(1).Enabled = False
                                opTpSvc(2).Enabled = True
                                opTpSvc(2).value = True
                            Else
                                If Not cekLastIsSBIDealer(Trim(rsTemp("NO_CHASSIS"))) Then
                                    If Not cekLastIsPDS(Trim(rsTemp("NO_CHASSIS"))) Then
                                        opTpSvc(0).Enabled = False
                                        opTpSvc(1).Enabled = False
                                        opTpSvc(2).Enabled = True
                                        opTpSvc(2).value = True
                                    Else
                                        If cekSBIDealerInvoice(Trim(rsTemp("NO_CHASSIS"))) Then
                                            opTpSvc(0).Enabled = False
                                            opTpSvc(1).Enabled = False
                                            opTpSvc(2).Enabled = True
                                            opTpSvc(2).value = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        simpan = False
        PanggilDataCustomer = True
    Else
        PanggilDataCustomer = False
    End If
End Function

Private Sub PanggilDataOrder(nRgk As String, nWO As String)
Dim rsTmp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
Dim rsOli As ADODB.Recordset
Dim n As Integer
        
    Set rsHlp2 = objO.FetchData("SELECT NO_WO FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(nWO) & "' AND NO_CHASSIS = '" & Trim(nRgk) & "' AND USERDEL = ''")
    If rsHlp2.EOF Then
        Set rsTemp = objO.FetchData("SELECT a.*,b.gsm,b.cdma,b.tgl_rep,b.rcdate,b.fs_km_akhir, b.Ddate, b.FS_kd_dealer, b.Fs_nm_dealer " _
                                 & "FROM jobdta a WITH (NOLOCK) INNER JOIN custdat b WITH (NOLOCK) ON a.no_chassis = b.no_chassis " _
                                 & "WHERE a.no_chassis = '" & nRgk & "' AND a.NO_WO = '" & nWO & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' ") 'AND TIPEWO <> 'W'")
                                 
        If Not rsTemp.EOF Then
            If rsTemp("RPR_COND") <> "K" Then
                text(1) = rsTemp("FS_KD_PLG")
                text(2) = rsTemp("FS_KD_WIL")
                text(3) = rsTemp("CUSTOMER")
                If Len(Trim(Nul(rsTemp("CUST_ADRES")))) > 60 Then
                    text(4) = Left(rsTemp("CUST_ADRES"), 60)
                    text(117) = Mid(rsTemp("CUST_ADRES"), 61, 60)
                Else
                    text(4) = Left(rsTemp("CUST_ADRES"), 60)
                End If
                text(5) = rsTemp("KOTA")
                text(6) = rsTemp("TELPON")
                text(7) = rsTemp("GSM")
                text(8) = rsTemp("CDMA")
                text(9) = rsTemp("FS_TH_MOBIL")
                text(10) = rsTemp("MDL")
                text(11) = rsTemp("MODEL")
                text(12) = rsTemp("FMDL")
                text(13) = rsTemp("COLOR")
                text(14).Mask = ""
                text(14) = rsTemp("NO_POLISI")
                cmdGantiNoPol.Visible = True
                text(15) = Nul(rsTemp("NO_CHASSIS"))
                text(16).Mask = ""
                On Error GoTo NextLine
                text(16) = TglDDMMYYYY(rsTemp("DOTAM"))
NextLine:
                text(16).Mask = "99-99-9999"
                text(79) = TglDDMMYYYY(rsTemp("DDATE"))
                text(118).Mask = ""
                text(118) = TglDDMMYYYY(rsTemp("DDATE"))
                text(118).Mask = "99-99-9999"
                text(17).Mask = ""
                text(17) = TglDDMMYYYY(rsTemp("TGL_REP"))
                text(18).Mask = ""
                text(18) = TglDDMMYYYY(rsTemp("RCDATE"))
                text(19) = "" 'rsTemp("")
                text(20) = Format(CDbl(rsTemp("FS_KM_AKHIR")), "#,#")
                If Trim(text(20)) = "" Then text(20) = 0
                text(23) = rsTemp("REQ1")
                text(24) = rsTemp("REQ2")
                text(25) = rsTemp("REQ3")
                text(26) = rsTemp("REQ4")
                text(27) = rsTemp("REQ5")
                text(90) = cekMOU(rsTemp("MDL"), Nul(rsTemp("NO_CHASSIS")))
                cmdLengkap.Visible = False
                If Trim(text(90)) <> "" Then
                    cmdLengkap.Visible = True
                End If
                chPDS = IIf(rsTemp("PDS") = 1, Checked, Unchecked)
                chEM = IIf(rsTemp("CODE") = 1, Checked, Unchecked)
                chSBNP = IIf(rsTemp("KD_K") = 1, Checked, Unchecked)
                text(76) = rsTemp("NO_ENGINE")
                text(104) = rsTemp("FRONT")
                text(105) = rsTemp("TELPANCUST")
                text(106) = rsTemp("GROUP")
                
                nilaiDPPart = CInt(rsTemp("NILAIDPPART"))
                                
                Dim hlp As String
                Set rsTmp = objO.FetchData("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    ' WALK AROUND CHECK
                    ' %%%% KONDISI KENDARAAN %%%%
                    text(34) = Nul(rsTmp("WAK1"))
                    text(35) = Nul(rsTmp("WAK2"))
                    text(36) = Nul(rsTmp("WAK3"))
                    text(37) = Nul(rsTmp("WAK4"))
                    text(38) = Nul(rsTmp("WAK5"))
                    
                    ' %%%% BARANG BERHARGA %%%%
                    text(39) = Nul(rsTmp("WAE1"))
                    text(40) = Nul(rsTmp("WAE2"))
                    text(41) = Nul(rsTmp("WAE3"))
                    text(42) = Nul(rsTmp("WAE4"))
                    text(43) = Nul(rsTmp("WAE5"))
                    
                    text(107) = Trim(rsTmp("DATANG"))
                    text(52) = Trim(rsTmp("DTLDATANG"))
                    cboPelanggan = Trim(rsTmp("KARAKTER"))
                    text(75) = Trim(rsTmp("DTLKARAKTER"))
                    
                    ' ISI KELUHAN
                    For i = 1 To 5
                        hlp = "CMPL" & i
                        Text1(-1 + i) = Trim(rsTmp(hlp))
                    Next
                    
                    chUse.value = IIf(rsTmp("HEAVYUSE") = 1, Checked, Unchecked)
                    
                    ' BBM
                    n = IIf(Nul(rsTmp("BBM")) = "", 4, Nul(rsTmp("BBM")))
                    opBBM(n).value = True
                    
                    ' FIELD _ FIELD LAIN
                    n = IIf(Nul(rsTmp("GANTIPART")) = "", 0, Nul(rsTmp("GANTIPART")))
                    opGanti(n).value = True
                    
                    n = IIf(Nul(rsTmp("TUNGGUSVC")) = "", 0, Nul(rsTmp("TUNGGUSVC")))
                    opSvc(n).value = True
                    
                    n = IIf(Nul(rsTmp("WASHING")) = 0, 1, 0)
                    opWash(n).value = True
                    
                    If Nul(rsTmp("WASHING")) = 0 Then
                        isWashing = False
                    Else
                        isWashing = True
                    End If
                    
                    Dim nBYR As Integer
                    nBYR = IIf(Nul(rsTmp("PEMBAYARAN")) = "CS", 0, IIf(rsTmp("PEMBAYARAN") = "CR", 1, 2))
                    If nBYR = 2 Then Updated = False
                End If
                
                ' %%%% KOORDINAT %%%%
                nk = 0
                bodyClear = True
                Set rsTmp = objO.FetchData("SELECT * FROM jobKoord WITH (NOLOCK) WHERE no_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        imgTanda(nk).Top = 600 + rsTmp("KOORDY") - (imgTanda(nk).Height / 2)
                        imgTanda(nk).Left = 600 + rsTmp("KOORDX") - (imgTanda(nk).Width / 2)
                        imgTanda(nk).Visible = True
                        KorX(nk) = rsTmp("KOORDX")
                        KorY(nk) = rsTmp("KOORDY")
                        nk = nk + 1
                        rsTmp.MoveNext
                    Loop
                    bodyClear = False
                End If
        
                ' TAMPIL CONDITION
                Dim kCond As String
                Dim kBant As String
                Set rsTmp = objO.FetchData("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    For i = 1 To 6
                        hlp = "COND" & i
                        kBant = ""
                        kCond = ""
                        If rsTmp(hlp) <> "" Then
                            If Left(rsTmp(hlp), 3) = "L| " Then
                                kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
                                kCond = "LAINNYA"
                            Else
                                kCond = Trim(rsTmp(hlp))
                                kBant = ""
                            End If
                        End If
                        text(27 + i) = kCond
                        bKondisi(i - 1) = kBant
                    Next
                End If
                
                ' ISI JOBORDER
                myChassis = Trim(text(15))
                ntRate = 0
                text(114) = 0
                For i = 1 To 8
                    tmp1 = "REP" & i
                    tmp2 = "RTE" & i
                    tmp3 = "MEKANIK" & i
                    DGrid.TextMatrix(i, 1) = rsTemp(tmp1)
                    If rsTemp(tmp1) <> "" Then
                        baris = baris + 1
                        Set rsTmp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(tmp1) & "'")
                        If Not rsTmp.EOF Then
                            DGrid.TextMatrix(i, 3) = False
                            DGrid.TextMatrix(i, 4) = rsTmp("JOB")
                        Else
                            DGrid.TextMatrix(i, 3) = True
                        End If
                        DGrid.TextMatrix(i, 2) = rsTemp(tmp2)
                        DGrid.TextMatrix(i, 5) = hitungLabour(CDbl(rsTemp(tmp2)), text(11))
                        DGrid.TextMatrix(i, 6) = rsTemp(tmp3)
                        ntRate = ntRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(rsTemp(tmp1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                        text(114) = ntRate
                        If isJualNRM = True Then
                            Me.JualNRM.Caption = nJualNRM
                        Else
                            Me.JualNRM.Caption = ""
                        End If
                    End If
                Next
                
                ' ISI SUBLET
                Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & nWO & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND KODEITEM <> 'A'")
                Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, NO_WO, NO_INV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, HARGA, NONOTA, KETERANGAN, FLAG, NO_POLISI,PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,estwaktu FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & nWO & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND KODEITEM <> 'A'")
                gPO = objO.FetchData_String("SELECT NOPO FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & nWO & "' AND NO_CHASSIS = '" & Trim(text(15)) & "'")
                
                ' ISI FU
                Set rsTmp = objO.FetchData("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    If Trim(rsTmp("TELPFU")) <> "" And Not IsNull(rsTmp("TELPFU")) Then
                        Call objO.InsertData("INSERT INTO TEMPFU SELECT NO_WO, NO_CHASSIS, TGLFU, JAMFU, TELPFU, KETFU FROM jobDtl WITH (NOLOCK) " _
                                     & "WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' AND " _
                                     & "NOT EXISTS (SELECT * FROM TEMPFU WITH (NOLOCK) WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "')")
                    End If
                    If Trim(rsTmp("TIMESERAH")) <> "" Then
                        text(88) = TglDDMMYYYY(rsTmp("DATESERAH"))
                        text(89) = Trim(rsTmp("TIMESERAH"))
                    End If
                End If
                
                ' ISI PART
                If rsGridPart.State <> adStateClosed Then rsGridPart.Close
                Set rsGridPart = Nothing
                If rsGridPart.State = adStateClosed Then createTblPart
                
                If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
                    Set rsTmp = objO.FetchDataACC("select KodeItem as KODE_ITEM, Keterangan as NAMA_ITEM, Qty, 'N' as GANTI, 'N' as Bawa, '' as KDOLI, BO as SAT_ITEM, NamaBYSA, Nomor as nID, '' as PARTBAWA from SOORD WITH (NOLOCK) Where NoSO = '" & rsTemp("NO_WO") & "' and AddJob = 0 and (isnull(BO,'null')='null' or BO <> 'CANCEL')")
                Else
                    'Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, SAT_ITEM FROM jobPart A LEFT JOIN T_OLI B ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & rsTemp("NO_CHASSIS") & "' AND no_wo = '" & rsTemp("NO_WO") & "' AND ADDJOB = 0 Order By nID asc")
                    Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, SAT_ITEM FROM jobPart A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & rsTemp("NO_CHASSIS") & "' AND no_wo = '" & rsTemp("NO_WO") & "' AND ADDJOB = 0")
                End If
                
                'Set rsTmp = objO.FetchDataACC("select KodeItem as KODE_ITEM, Keterangan as NAMA_ITEM, Qty, 'N' as GANTI, 'N' as Bawa, '' as KDOLI, BO as SAT_ITEM, NamaBYSA, Nomor as nID, '' as PARTBAWA from SOORD Where NoSO = '" & rsTemp("NO_WO") & "' and BO <> 'CANCEL' and AddJob = 0")
                'If rsTmp.RecordCount = 0 Then
                '    Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, SAT_ITEM FROM jobPart A LEFT JOIN T_OLI B ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & rsTemp("NO_CHASSIS") & "' AND no_wo = '" & rsTemp("NO_WO") & "' AND ADDJOB = 0 Order By nID asc")
                'End If
                            
                
                
                'Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, SAT_ITEM FROM jobPart A LEFT JOIN T_OLI B ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & rsTemp("NO_CHASSIS") & "' AND no_wo = '" & rsTemp("NO_WO") & "' AND ADDJOB = 0 Order By nID asc")
                If Not rsTmp.EOF Then
                    n = 69
                    Do While Not rsTmp.EOF
                        With rsGridPart
                            .AddNew
                            !KodeItem = rsTmp("KODE_ITEM")
                            !NAMAITEM = rsTmp("NAMA_ITEM")
                            !SATITEM = IIf(IsNull(rsTmp("SAT_ITEM")), "", rsTmp("SAT_ITEM"))
                            !qty = CDbl(rsTmp("QTY"))
                            !ganti = IIf(rsTmp("GANTI") <> "N", True, False)
                            !BAWA = IIf(rsTmp("BAWA") <> "N", True, False)
                            !oli = IIf(rsTmp("KDOLI") <> "", True, False)
                            !namabySA = rsTmp("NAMABYSA")
                            !nID = Val(NulA(rsTmp("nID")))
                            .Update
                        End With
                        If rsTmp("PARTBAWA") <> "" Then
                            text(n) = rsTmp("PARTBAWA")
                            n = n + 1
                        End If
                    rsTmp.MoveNext
                    Loop
                    rsTmp.MoveLast
                    NewId = Val(NulA(rsTmp("nID"))) + 1
                    Call HeadColumn
                    Set partGrid.DataSource = rsGridPart
                    Set partGridFill.DataSource = rsGridPart
                End If
                
                ' ISI KELENGKAPAN
                Set rsTmp = objO.FetchData("SELECT SPKKet,LainKet FROM JOBLENGKAP WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "'")
                If Not rsTmp.EOF Then
                    If Trim(rsTmp("SPKKet")) <> "" Then
                        opLengkap(0) = Checked
                        text(115) = Trim(rsTmp("SPKKet"))
                    End If
                    If Trim(rsTmp("LainKet")) <> "" Then
                        opLengkap(1) = Checked
                        text(116) = Trim(rsTmp("LainKet"))
                    End If
                End If
                
                cmdWarranty.Enabled = False
                If Trim(rsTemp("TPSVC")) = "W" Then
                    opTpSvc(0).Enabled = False
                    opTpSvc(1).Enabled = False
                    opTpSvc(2).value = True
                    opTpSvc(2).Enabled = True
                    text(21) = rsTemp("KODETIPE")
                    text(22) = objO.FetchData_String("SELECT WARRANTY FROM HO_WARRANTY WITH (NOLOCK) WHERE ID_WARRANTY = '" & Trim(text(21)) & "'")
                ElseIf Trim(rsTemp("TPSVC")) = "P" Then
                    opTpSvc(1).value = True
                    text(21) = rsTemp("KODETIPE")
                    text(22) = objO.FetchData_String("SELECT PROGRAM FROM HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & Trim(text(21)) & "'")
                ElseIf Trim(rsTemp("TPSVC")) = "" Then
                    opTpSvc(0).value = True
                    text(21) = ""
                    opTpSvc(2).Enabled = False
                End If
                
                text(102) = rsTemp("ANCUST")
                text(50) = rsTemp("MECH")
                text(51) = rsTemp("KM")
                Txt(53) = Left(rsTemp("CATAT"), 250)
                Txt(87) = Right(rsTemp("CATAT"), Len(rsTemp("CATAT")) - IIf(Len(rsTemp("CATAT")) > 250, 250, Len(rsTemp("CATAT"))))
                Txt(54) = rsTemp("REMARK1")
                Txt(55) = rsTemp("SUGGEST")
                text(56) = Format(CDbl(rsTemp("LABOR")), "0#,#")
                text(57) = Format(CDbl(rsTemp("PART")), "0#,#")
                text(58) = Format(CDbl(rsTemp("MTRL")), "0#,#")
                text(59) = Format(CDbl(rsTemp("OIL")), "0#,#")
                text(60) = Format(CDbl(rsTemp("SBLT")), "0#,#")
                text(61) = Format(CDbl(rsTemp("TOTAL")), "0#,#")
                text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                
                '' NO ANTRIAN
                Set rsTmp = objO.FetchData("SELECT QUEUE, TIPE FROM LISTANTRIAN WITH (NOLOCK) WHERE WO = '" & Trim(text(0)) & "'")
                If Not rsTmp.EOF Then
                    text(108) = Trim(rsTmp("QUEUE"))
                    text(109) = Trim(rsTmp("TIPE"))
                    noAntriLama = Trim(text(108))
                End If
                
                ' HITUNG NILAI PART & OLI
                If rsGridPart.State <> adStateClosed Then
                    text(59) = 0
                    text(57) = 0
                    If Not rsGridPart.EOF Then
                    rsGridPart.MoveFirst
                    Do While Not rsGridPart.EOF
                        Set rsOli = objO.FetchDataACC("SELECT hargaJualA FROM inmst WITH (NOLOCK) WHERE kodeItem = '" & rsGridPart("KODEITEM") & "'")
                        If Not rsOli.EOF Then
                            If rsGridPart("OLI") Then
                                text(59) = text(59) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                            Else
                                text(57) = text(57) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                            End If
                        End If
                        rsGridPart.MoveNext
                    Loop
                    End If
                End If
                
                text(103) = rsTemp("STATUS")
                
                If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                    If Trim(text(64)) <> "" Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where NOSO = '" & Trim(text(64)) & "'")
                    text(64) = Nul(rsTemp("NO_WO"))
                    If Trim(text(64)) <> "" Then Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NOSO = '" & Trim(text(64)) & "'")
                Else
                    text(64) = Nul(rsTemp("NO_WO"))
                End If
                
                text(65) = TglDDMMYYYY(rsTemp("DATE"))
                text(66) = rsTemp("TIME")
                text(67) = rsTemp("NOINV")
                'If Nul(rsTemp("NO_WO")) <> "" And Text(64) <> "" Then
                If Nul(rsTemp("NO_WO")) <> "" And text(64) = "" Then
                    text(78) = ""
                Else
                    text(78) = Nul(rsTemp("NOFAK"))
                End If
                ChkReturnOrder = IIf(Trim(rsTemp("FN_RETURN")) = 1, Checked, Unchecked)
                ChkReturnOrder.Enabled = True
                chTHS = IIf(rsTemp("THS") = 1, Checked, Unchecked)
                simpan = False
                opByr(nBYR).value = True
                
                text(44).BackColor = &H80000005
                text(46).BackColor = &H80000005
                
                ' CHECK STALL
                Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & nWO & "|%'")
                If Not rsTmp.EOF Then
                    ' TUTUP TAMBAH PEKERJAAN
                    If rsTmp("FLAG") <> "N" Then
                        text(44).BackColor = &HC0E0FF
                        text(46).BackColor = &HC0E0FF
                        woOnProgres = True
                    End If
                    If UCase(rsTmp("FLAG")) = "G" Then
                        MsgBox "Unit sedang menunggu proses cuci", vbInformation, "Informasi Unit"
                    ElseIf UCase(rsTmp("FLAG")) = "A" Then
                        MsgBox "Unit sedang proses cuci", vbInformation, "Informasi Unit"
                    ElseIf UCase(rsTmp("FLAG")) = "F" Then
                        MsgBox "Unit telah selesai proses cuci", vbInformation, "Informasi Unit"
                    End If
                    
                    cmdStall.Tag = ""
                    sgnStall.Visible = True
                    sgnStall.Caption = Trim(rsTmp("NOSTALL")) '& " : " & Trim(rsTmp("JAMMULAI")) & " s/d " & Trim(rsTmp("JAMSELESAI"))
                    
                    Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN like '%|" & nWO & "|%'")
                    Call objO.InsertData("INSERT INTO tempStall SELECT noStall, tanggal, JamMulai, JamSelesai, rate, [Group], Mech, Flag, Keterangan,adjust FROM STALL WITH (NOLOCK) " _
                                 & "WHERE KETERANGAN like '%|" & nWO & "|%' AND NOT EXISTS " _
                                 & "(SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE KETERANGAN like '%|" & nWO & "|%' )") 'AND Tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'
                                 
                    text(86) = Trim(rsTmp("JAMMULAI"))
                    
                    Select Case rsTmp("FLAG")
                        Case "I"
                            If rsTemp("NOINV") <> "" Then
                                pesan = "Sudah diinvoice dengan No Invoice : " & Trim(rsTemp("NOINV")) & "  ,Tanggal : " & TglDDMMMMYYYY(rsTemp("INVDATE"))
                                Call tutupForm
                            Else
                                pesan = IIf(Nul(rsTemp("WASHING")) = 0, "Kendaraan sudah selesai dikerjakan", "Kendaraan dalam proses washing")
                            End If
                            MsgBox pesan
                        Case "S"
                            MsgBox "Kendaraan dalam proses pengerjaan"
                        Case "P"
                            MsgBox "Kendaraan dalam proses dihentikan sementara"
                        Case "H"
                            MsgBox "Kendaraan menunggu proses Final Inspection"
                        Case "Q"
                            MsgBox "Kendaraan dalam proses Final Inspection"
                        Case "G"
                            MsgBox "Kendaraan menunggu proses Washing"
                        Case "A"
                            MsgBox "Kendaraan dalam proses Washing"
                        Case "F"
                            MsgBox "Kendaraan sudah selesai dikerjakan"
                        Case "K"
                            MsgBox "Kendaraan sudah Closing"
                            Call tutupForm
                        Case ""
                            MsgBox "Kendaraan dalam proses menunggu perbaikan"
                    End Select
                Else
                    cmdStall.Tag = "belum"
                    If chPDS.value = Checked Then
                        woOnProgres = True
                    End If
                End If
                TblSimpan.Enabled = True
                cmdAddJob.Enabled = True
                If rsTemp("NOINV") <> "" Then
                    'MsgBox "WO ini sudah di Invoice !!"
                    TblSimpan.Enabled = False
                    cmdAddJob.Enabled = False
                End If
                Dim rsTempDiagnose As ADODB.Recordset
        
                Set rsTempDiagnose = objO.FetchData("SELECT * FROM diagapproval WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(0)) & "'")
                If Not rsTempDiagnose.EOF Then
                    'jika booking tersebut termasuk dalam diagnose
                    opTpSvc(2).Enabled = True
                    opTpSvc(2).value = True
                    text(22).Visible = True
                    text(22) = rsTempDiagnose.Fields(0)
                    opTpSvc(0).Enabled = False
                    opTpSvc(1).Enabled = False
                    
                End If
            Else
                MsgBox "WO ini sudah di Closing "
                Exit Sub
            End If
        End If
    Else
        MsgBox "WO ini sudah divalidasi"
        TblSimpan.Enabled = False
    End If
End Sub

Private Function isiNOWO() As String
    'isiNOWO = objO.FetchData_String("SELECT FN_NO_WO FROM T_PROFIL WITH (NOLOCK)")
    isiNOWO = objO.FetchData_String("SELECT FN_NO_WO FROM T_PROFIL")
End Function

Private Function isiNoPRE() As String
    Dim xDate As Date
    xDate = isnow
    'Set rsHlp = objO.FetchData("SELECT DISTINCT NO_DIAG FROM DIAGDTA WITH (NOLOCK) WHERE SUBSTRING(NO_DIAG,2,4) = '" & Format(Right(Year(isnow), 2), "00") & Format(Month(isnow), "00") & "' ORDER BY NO_DIAG DESC")
    Set rsHlp = objO.FetchData("SELECT DISTINCT NO_DIAG FROM DIAGDTA WHERE SUBSTRING(NO_DIAG,2,4) = '" & Format(Right(Year(isnow), 2), "00") & Format(Month(isnow), "00") & "' ORDER BY NO_DIAG DESC")
    If Not rsHlp.EOF Then
        isiNoPRE = "9" & Format(Right(Year(xDate), 2), "00") & Format(Month(xDate), "00") & Format(CDbl(Right(rsHlp("NO_DIAG"), 3)) + 1, "000")
    Else
        isiNoPRE = "9" & Format(Right(Year(xDate), 2), "00") & Format(Month(xDate), "00") & Format(1, "000")
    End If
End Function


Private Sub Label13_Click()
    Call cmdHistory_Click
End Sub


Private Sub LbSupply_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrSupply.Visible = False
End Sub

Private Sub cmdView_Click()
    If FrSupply.Visible = True Then
        FrSupply.Visible = False
        Exit Sub
    End If
    Dim rsHlp3 As ADODB.Recordset
    With rsPart1
        Dim strSQL As String
        If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
            strSQL = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
                 "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
                 "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
                 "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(text(0)) & "'"

        Else
            strSQL = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QTYCANCEL FROM INMST " & _
                "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                "Where Z.Norack = '" & Trim(text(0)) & "'"
        End If
        Set rsHlp3 = objO.FetchDataACC(strSQL)

        Call doCreatePart
        If Not rsHlp3.EOF Then
            Do While Not rsHlp3.EOF
                .AddNew
                !KodeItem = Trim(rsHlp3("KODEITEM"))
                !NAMAITEM = Trim(rsHlp3("NAMABARANG"))
                !QtyOrder = CDbl(rsHlp3("QtyOrder"))
                !qty = CDbl(rsHlp3("JML"))
                If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                    !QtyBO = objO.FetchData_DoubleACC("select Qty FROM BOFILE WITH (NOLOCK) Where TutupOrderFlag = 0 and (NoSO = '" & Trim(text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD WITH (NOLOCK) Where NOSO ='" & Trim(text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'")
                Else
                    !QtyBO = 0
                End If
                '!QtyBO = IIf(UCase(isBO) = "CANCEL", 0, objO.FetchData_DoubleACC("select Qty FROM BOFILE Where TutupOrderFlag = 0 and (NoSO = '" & Trim(Text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD Where NOSO ='" & Trim(Text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'"))
                !QtyCancel = CDbl(rsHlp3("QtyCANCEL"))
                !Harga = CDbl(rsHlp3("HARGAJUALA"))
                Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
                If Not rsHlp.EOF Then
                    If rsHlp("TIPE") = "O" Then
                        !oli = True
                    ElseIf rsHlp("TIPE") = "P" Then
                        !oli = False
                    End If
                End If
                .Update
                rsHlp3.MoveNext
            Loop
            Call doHeadPart1
            FrSupply.Visible = True
        End If
    End With
End Sub


Private Sub doCreatePart()
    With rsPart1
        If .State Then .Close
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NAMAITEM", adVarChar, 200
        .Fields.Append "QTYOrder", adInteger
        .Fields.Append "QTY", adInteger
        .Fields.Append "QTYBO", adInteger
        .Fields.Append "QTYCANCEL", adInteger
        .Fields.Append "HARGA", adDouble
        .Fields.Append "OLI", adBoolean
        .Fields.Append "EXTDISC", adDouble
        .Open
    End With
    Set PartGrid1.DataSource = rsPart1
End Sub

Private Sub doHeadPart1()
    With PartGrid1
        .Columns(0).Width = 2000
        .Columns(0).Caption = "KODE PART"
        .Columns(0).DataField = "KODEITEM"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 2500
        .Columns(1).Caption = "NAMA PART"
        .Columns(1).DataField = "NAMAITEM"
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 1000
        .Columns(2).Caption = "QTY Order"
        .Columns(2).DataField = "QTYOrder"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 900
        .Columns(3).Caption = "QTY Supply"
        .Columns(3).DataField = "QTY"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(4).Width = 900
        .Columns(4).Caption = "QTY BO"
        .Columns(4).DataField = "QTYBO"
        .Splits(0).Columns(2).Locked = True
        On Error GoTo Column_Ini_Untuk_Grid_FillRate
        
        .Columns(6).Visible = False
        .Columns(6).Width = 1000
        .Columns(6).Caption = "HARGA"
        .Columns(6).DataField = "Harga"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(5).Visible = True
        .Columns(5).Width = 1000
        .Columns(5).Caption = "QTY CANCEL"
        .Columns(5).DataField = "QtyCANCEL"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(7).Visible = False
        .Columns(7).Caption = "OLI"
        .Columns(7).DataField = "OLI"
        .Splits(0).Columns(5).Locked = True
Column_Ini_Untuk_Grid_FillRate:
    End With
End Sub

Private Sub opWash_Click(Index As Integer)
    If isWashing = True Then
        If opWash(1).value = True Then
            'jika pembatalan washing
            If Trim(text(64)) <> "" Then
                Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & Trim(text(64)) & "|%'")
                If Not rsTmp.EOF Then
                    ' TUTUP TAMBAH PEKERJAAN
                    
                    If UCase(rsTmp("FLAG")) = "G" Then
                        framePembatalanWashing.Visible = True
                        opWash(0).value = True
                    Else
                        opWash(0).value = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub TblSimpan_Click()
     If Trim(text(64)) <> "" Then
        isUpdateFill = True
    Else
        isUpdateFill = False
    End If
    Dim ARBOLK As Boolean
    isVAR = False
    ARBOLK = False
    Me.PicARO.Visible = False
    Dim is80 As Boolean
    is80 = False
        Dim nIsAR As Boolean
        nIsAR = True
        If Me.opTpSvc(2).value = True Then nIsAR = False  'warranty
        If isSBI = True Then nIsAR = False                'SBI
        If chPDS.value = 1 Then nIsAR = False             'PDS
        If ChkReturnOrder.value = 1 Then nIsAR = False    'RTJ
        
            If Log.LockAR = True And woOnProgres = False And Trim(text(64).text) = "" Then
                If nIsAR = True Then
                    If isOverdueAR = True Then
                        Me.PicARO.Visible = True
                        If objO.FetchIsValid_AR_Kabeng(text(1).text) = False Then
                            Me.lblOto.Caption = "OTORITAS KACAB"
                            Me.PicARO.Visible = True
                            MsgBox "Ada AR Overdue !" & Chr(13) & Chr(13) & "WORK ORDER TIDAK DAPAT DIPROSES !" & Chr(13) & "OTORITAS KEPALA CABANG", vbInformation + vbOKOnly
                            Exit Sub
                        Else
                            MsgBox "AR OVERDUE" & Chr(13) & "TELAH DIVALIDASI OLEH KACAB", vbInformation + vbOKOnly
                            FrAR.Visible = False
                        End If
                        ARBOLK = True
                    Else
                        If isCurrentAR = True Then
                            If nPersenAR >= 100 Then
                                If objO.FetchIsValid_AR_Kabeng(text(1).text) = False Then
                                    Me.lblOto.Caption = "OTORITAS KACAB"
                                    Me.PicARO.Visible = False
                                    MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "Work Order Tidak Dapat Di Proses !" & Chr(13) & Chr(13) & "OTORITAS KEPALA CABANG", vbInformation + vbOKOnly
                                    Exit Sub
                                Else
                                    MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "TELAH DIVALIDASI OLEH KACAB", vbInformation + vbOKOnly
                                End If
                                ARBOLK = True
                            ElseIf nPersenAR >= 80 Then
                                Me.lblOto.Caption = "W A R N I N G"
                                Me.PicARO.Visible = False
                                If MsgBox("AR sudah >= 80% dari Batas AR !" & Chr(13) & "LANJUTKAN PROSES ?", vbQuestion + vbYesNo) = vbNo Then
                                    Exit Sub
                                End If
                                is80 = True
                            End If
                        End If
                    End If
                End If
            End If


    If text(0) <> "" Then
        pesan = "Lengkapi data berikut !" & vbCr
        
        boleh = True
        onlySublet = False
        If text(51) = 0 Then
            pesan = pesan + "KM. "
            boleh = False
        End If
        Set rsTemp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE KETERANGAN LIKE '" & Trim(text(14)) & " |%'")
        If rsTemp.EOF Then
            cmdStall.Tag = "belum"
        End If
                
        If text(1) = "" Then
            pesan = pesan + "Kode Pelanggan. "
            boleh = False
        End If
        If text(15) = "" Then
            pesan = pesan + "No. Rangka. "
            boleh = False
        End If
        If text(105) = "" Then
            pesan = pesan + "No. Telpon Pembawa. "
            boleh = False
        End If
        If Not woOnProgres Then
            If text(65) = text(88) Then
                If DateDiff("d", text(89), text(66)) > 0 Then
                    pesan = pesan & vbCr & "Waktu penyerahan harus lebih besar dari waktu sekarang. "
                    boleh = False
                End If
            ElseIf DateDiff("d", text(88), text(65)) > 0 Then
                pesan = pesan & vbCr & "Waktu penyerahan harus lebih besar dari waktu sekarang. "
                boleh = False
            End If
            
            If DateDiff("d", Format(text(88), "DD/MM/YYYY"), Format(text(65), "DD/MM/YYYY")) = 0 Then
                If DateDiff("s", Format(text(63), "HH:NN"), Format(text(89), "HH:NN")) < 0 Then
                    boleh = False
                    pesan = pesan + "Waktu Penyerahan tidak boleh lebih kecil dari Waktu Selesai ! "
                    text(89) = Format(text(63), "HH:NN")
                End If
            End If
        End If
        
        If opSvc(0).value = False And opSvc(1).value = False Then
            pesan = pesan + "Waiting Customer."
            boleh = False
        End If
        
        If text(88) = ">>->>->>>>" Or text(89) = ">>:>>" Then
            pesan = pesan + "Waktu Penyerahan. "
            boleh = False
        End If
        If text(23) = "" And text(24) = "" And text(25) = "" And text(26) = "" And text(27) = "" Then
            pesan = pesan + "Permintaan Customer"
            boleh = False
        End If
        Set rsTemp = objO.FetchData("SELECT FS_KD_PLG FROM CUSTDAT WITH (NOLOCK) WHERE no_Chassis = '" & text(15) & "'") 'AND fs_kd_plg = '" & text(1) & "'
        If Not rsTemp.EOF Then
            If Trim(rsTemp("FS_KD_PLG")) <> Trim(text(1)) Then
                If MsgBox("Kode Pelanggan berbeda, tetap akan melanjutkan ?", vbYesNo, "Pelanggan") = vbNo Then
                    pesan = pesan & vbCr & "Data Customer tidak sesuai. Isi Master customer. "
                    boleh = False
                End If
            End If
        Else
            pesan = pesan & vbCr & "Data Customer belum ada.Isi Master customer. "
            boleh = False
        End If
        
        If chPDS = Unchecked And Not sgnPreDiag.Visible Then
            Set rsTemp = objO.FetchData("SELECT No_Wo FROM TEMPFU WITH (NOLOCK) WHERE no_Chassis = '" & text(15) & "' AND NO_WO = '" & Trim(text(64)) & "'")
            If rsTemp.EOF Then
                pesan = pesan & vbCr & "3 Days Follow Up belum ada. "
                boleh = False
            End If
        End If
                
        If text(50) <> "" And cmdPilih.Tag <> "No. Booking" And cmdPilih.Tag <> "No. WO" Then
            nTgl = isnow
            gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
            gBantuan3 = Format(nTgl, "DD/MM/YYYY")
            Set rsTemp = objO.FetchData("select top 1 A.MECH from mekanik A WITH (NOLOCK)  LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH AND hari = '" & gBantuan2 & "' AND TANGGAL = '" & gBantuan3 & "' LEFT JOIN T_STALL C WITH (NOLOCK) ON C.HARI = '" & gBantuan2 & "' WHERE (MENITKERJA-ALOKASI) > 0 AND A.MECH = '" & text(50) & "'")
            If rsTemp.EOF Then
                boleh = False
                pesan = pesan + "Waktu Mekanik tidak mencukupi"
            End If
        End If
        
        If text(103) = "" Then
            boleh = False
            pesan = pesan + "Tipe Customer. "
        ElseIf Trim(text(103)) = "EXT" Then
            If MsgBox("Yakin tipe customer EXT ?", vbYesNo, "TIPE CUSTOMER") = vbNo Then
                boleh = False
                pesan = pesan + "Tipe Customer. "
            End If
        End If
        
        If text(107) = "" Then
            boleh = False
            pesan = pesan + "Alasan Kedatangan. "
        End If
        
        If text(102) = "" Then
            boleh = False
            pesan = pesan + "Nama Pembawa Kendaraan. "
            text(102).SelLength = Len(text(102))
            'Text(102).SetFocus
        End If
        
        bodyClear = False
        If nk = 0 Then
            If MsgBox("Kondisi body kendaraan baik? ", vbYesNo, "Walk Around Check") = vbYes Then
                bodyClear = True
            Else
                boleh = False
                pesan = pesan + "Walk Around Check"
            End If
        End If
        
        If Trim(text(20)) = "" Then text(20) = 0
        If CDbl(text(51)) < CDbl(text(20)) Then
            boleh = False
            pesan = pesan + "KM sekarang harus lebih besar dari KM terakhir !!"
        End If
        
        If opTpSvc(1).value Then
            If Trim(text(21)) = "" Then
                boleh = False
                pesan = pesan + "Tipe Program. "
            End If
        End If
        Dim rsQu As ADODB.Recordset
        If gQueue Then
            If Val(Replace(Replace(Trim(text(108)), "C", ""), ">", "")) < 1 Or Trim(text(109)) = "" Then
            'If (Trim(text(108)) = "" Or Trim(text(108)) = "0") Or Trim(text(109)) = "" Then
                If chTHS.value = 1 Or chPDS.value = 1 Or opTpSvc(2).value = True Or ChkReturnOrder.value = 1 Then
                    ' boleh tanpa antrian
                    text(108) = "0"
                    text(109) = ""
                Else
                    'GANTI LEN TIPE MENJADI <=3
                    
'                    Set rsQu = objO.FetchData("SELECT DISTINCT QUEUE, TIPE FROM " & Log.LogQueue & "..PANGGILAN WHERE SA = '" & gSA & "' AND NOPOL = '' AND " & _
'                                              "ID NOT IN (SELECT IDCALL FROM LISTANTRIAN) AND LEN(TIPE) = 3 AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & Format(TglYYYYMMDD(Text(65)), "MM/DD/YYYY") & "' " & _
'                                              "AND CONVERT(varchar(8), TGL_PANGGIL, 8) <= '" & Format(isnow, "HH:NN") & "' " & _
'                                              "AND QUEUE = " & Val(Replace(Replace(Trim(Text(108)), "C", ""), ">", "")) & " " & _
'                                              "AND TIPE = '" & Trim(Text(109)) & "'")

                    Set rsQu = objO.FetchData("SELECT DISTINCT QUEUE, TIPE FROM " & Log.LogQueue & "..PANGGILAN WITH (NOLOCK) WHERE SA = '" & gSA & "' AND NOPOL = '' AND " & _
                                              "ID NOT IN (SELECT IDCALL FROM LISTANTRIAN WITH (NOLOCK)) AND LEN(TIPE) <= 3 AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & Format(TglYYYYMMDD(text(65)), "MM/DD/YYYY") & "' " & _
                                              "AND CONVERT(varchar(8), TGL_PANGGIL, 8) <= '" & Format(isnow, "HH:NN") & "' " & _
                                              "AND QUEUE = " & Val(Replace(Replace(Trim(text(108)), "C", ""), ">", "")) & " " & _
                                              "AND TIPE = '" & Trim(text(109)) & "'")
                                              
                    If rsQu.RecordCount = 0 Then
                        If Trim(text(64)) = "" Then
                            MsgBox "Pembuatan WO harus dengan ANTRIAN !!!" & Chr(13) & "Sesuai Login SA", vbCritical + vbOKOnly
                            Exit Sub
                        End If
                    End If
                    'If gQueue Then
                    '    pesan = pesan + "No Antrian. "
                    '    boleh = False
                    'Else
                    '    If MsgBox("Tidak dengan antrian ?", vbYesNo, "Isian Antrian") = vbNo Then
                    '        pesan = pesan + "No Antrian. "
                    '        boleh = False
                    '    End If
                    'End If
                End If
            'End If
            Else
                If chTHS.value = 1 Or chPDS.value = 1 Or opTpSvc(2).value = True Or ChkReturnOrder.value = 1 Then
                    ' boleh tanpa antrian
                    text(108) = "0"
                    text(109) = ""
                Else
                    'UPDATE ANTRIAN EM
                    Set rsQu = objO.FetchData("SELECT DISTINCT QUEUE, TIPE FROM " & Log.LogQueue & "..PANGGILAN WITH (NOLOCK) WHERE SA = '" & gSA & "' AND NOPOL = '' AND " & _
                                              "ID NOT IN (SELECT IDCALL FROM LISTANTRIAN WITH (NOLOCK)) AND LEN(TIPE) <= 3 AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & Format(TglYYYYMMDD(text(65)), "MM/DD/YYYY") & "' " & _
                                              "AND CONVERT(varchar(8), TGL_PANGGIL, 8) <= '" & Format(isnow, "HH:NN") & "' " & _
                                              "AND QUEUE = " & Val(Replace(Replace(Trim(text(108)), "C", ""), ">", "")) & " " & _
                                              "AND TIPE = '" & Trim(text(109)) & "'")
                    If rsQu.RecordCount = 0 Then
                        If Trim(text(64)) = "" Then
                            MsgBox "Pembuatan WO harus dengan ANTRIAN !!!" & Chr(13) & "Sesuai Login SA", vbCritical + vbOKOnly
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        If chPDS.value = 0 And chTHS.value = 0 Then 'UNCHECKED
            If baris = 1 Then
                Set rsTemp = objO.FetchData("SELECT Top 1 No_Wo FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(15)) & "'")
                If rsTemp.EOF Then
                    pesan = pesan + "Job Order. "
                    boleh = False
                    
                    If cmdStall.Tag = "belum" Then
                        pesan = pesan + "Stall. "
                        boleh = False
                    End If
                Else
                    ' HANYA INPUT JOB ORDER
                    If MsgBox("Belum input Job Order !!, Tetap akan melanjutkan ?", vbYesNo) = vbNo Then
                        pesan = pesan + "Job Order. "
                        boleh = False
                        If cmdStall.Tag = "belum" Then
                            pesan = pesan + "Stall. "
                            boleh = False
                        End If
                    Else
                        onlySublet = True
                    End If
                End If
            Else
                If cmdStall.Tag = "belum" Then
                    pesan = pesan + "Stall. "
                    boleh = False
                End If
            End If
        Else
            If baris = 1 Then
                pesan = pesan + "Job Order. "
                boleh = False
            End If
        End If
        
        '20/12/2012  Cust Non Axtif
        Dim rsNon As ADODB.Recordset
        Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & text(1) & "' and nonaktifflag = 0")
        If rsNon.RecordCount = 0 Then
            boleh = False
            pesan = "Kode Langgan : " & text(1).text & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Rubah Kode Langgan"
        End If
        
        If Log.DmsNonDss = 1 Then
            'lakukan pengecekan apakah harus menginputkan DP part
            
            If ((Log.persentaseDP / 100) * CDbl(Trim(text(0)))) > Log.plafonDP Then
                'lakukan pengisian DP
                MsgBox "Nilai part membutuhkan DP!", vbInformation, "Informasi DP Part"
                boleh = False
            End If
        End If
        If boleh Then
            Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderDMSNonDSS", "tblSimpan", 0)
            erTranAll = False
            onTR = True
            
            If newWo Then
                Dim gTglTampung As Date
                gTglTampung = gTglServer
                text(65) = TglDDMMYYYY(gTglTampung)
                text(66) = Format(gTglTampung, "HH:NN")
            End If
            
            Call objHT.CloseTransactionAll_Err
            Call objHT.OpenTransactionAll
            
            If Not sgnPreDiag.Visible Then
                'SIMPAN ORDER
                ' %%%% INPUT JOBDTA %%%%
                Call simpanJOBDTA
                If newWo Then
                    Call objHT.UpdateData_Bengkel("UPDATE t_profil SET FN_NO_WO = FN_NO_WO + 1")
                End If
                
                If opTpSvc(2).value = True And Trim(text(22)) <> "" Then
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGDTL SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGKOORD SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGLENGKAP SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGPART SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGSUBLET SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGAPPROVAL SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                    Call objHT.UpdateData_Bengkel("UPDATE DIAGADD2 SET NO_WO = '" & Trim(text(64)) & "' WHERE NO_DIAG = '" & Trim(text(22)) & "' ")
                End If
            
                ' %%%% INPUT JOBDTL %%%%
                Call simpanJOBDTL
                
                ' %%%% INPUT KOORDINAT & STALL %%%%
                Call simpanJOBKoord
                
                ' %%%% INPUT PART & SUBLET %%%%
                Call simpanJOBPART
                
                ' %%%% ANTRIAN %%%%
                Call simpanAntrian(IIf(Trim(noAntriLama) = "", 1, 3), Trim(text(109)), Val(Replace(Replace(Trim(text(108)), "C", ""), ">", "")), text(65), Trim(text(64)), Trim(text(14)))
                
                If Trim(text(78)) <> "" Then Call removeNextDayBooking(Trim(text(78)), True)
                
                ' JIKA DARI BOOKING, UPDATE STATUS DELIV BOOK = OKE
                If cmdPilih.Tag = "No. Booking" Then
                    Call objHT.UpdateData_Bengkel("UPDATE BOOK SET deliv = 'OKE' WHERE nofak = '" & Trim(text(0)) & "'")
                End If
                
            Else
                '' SIMPAN PREDIAG
                Call simpanPrediag
                Call simpanPreDtl
                Call simpanPreKoord
                Call simpanPrePart
                'Call cetakFormPrediag
            End If
        
                If is80 = True Then Call SimpanAR80
                If ARBOLK = True Then Call objHT.Update_LogAR(Trim(text(1).text), Trim(text(64).text), nAR2.Caption)
            
            If Not sgnPreDiag.Visible Then
                    If Log.LogHOConn = 1 Then
                        If ConHO = False Then
                            Call Saving_Log_HO("Insert", sSQLHO)
                        Else
                            If Get_ConnectionProfil_Bengkel99 = False Then
                                Call Saving_Log_HO("Insert", sSQLHO)
                            Else
                                Call objHO.OpenTransactionLog_NRM
                                Call objHO.AddNew_WO(sSQLHO)
                                Call objHO.CloseTransactionLog_NRM
                                If errHO = True Then Call Saving_Log_HO("Insert", sSQLHO)
                            End If
                        End If
                    Else
                        Call Saving_Log_HO("Insert", sSQLHO)
                    End If
            End If
                        
            If erTranAll = False Then
                Call objHT.CloseTransactionAll
                onTR = False
                Dim nRP As Double
                If txtRp(1) > 0 Then
                    nRP = Round((txtRp(0) / txtRp(1)) * 100, 2)
                Else
                    nRP = 0
                End If
                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderDMSNonDSS", "tblSimpan", 1)
                If Not sgnPreDiag.Visible Then
                    Call CetakForm1
                Else
                    Call cetakFormPrediag
                End If
                Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_CHASSIS = '" & Trim(text(15)) & "'")    'NO_WO = '" &Trim(text(64)) & "' AND
                Call objO.DeleteData("DELETE TEMPFU WHERE NO_CHASSIS = '" & Trim(text(15)) & "'")   'NO_WO = '" &Trim(text(64)) & "' AND
                Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN LIKE '" & Trim(text(14)) & " |%'") 'AND MECH = '" & Trim(Text(50)) & "'
                'Call CetakForm1
                If Trim(text(21)) <> "PRE" Then
                    If opTpSvc(2) Then
                        Call cekWarranty(Trim(text(15)), Trim(text(21)))
                    End If
                End If
                If updateHistory Then
                    If Get_Connectiontemp("", "", "", "") Then
                        Set rsHlp = objO.FetchDataAAS("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(15)) & "'")
                        If Not rsHlp.EOF Then
                            Call transferHistoryLama(Trim(text(15)), rsHlp)
                        End If
                    End If
                End If
                Call hapusForm

                Layar.Tag = ""
                simpan = True
                Exit Sub
            Else
                Call objHT.CloseTransactionAll_Err
                onTR = False
                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderDMSNonDSS", "tblSimpan", 2)
                MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
                Exit Sub
            End If
        Else
            MsgBox pesan, vbCritical, "LENGKAPI DATA"
            Exit Sub
        End If
    Else
        MsgBox "Input data yg akan diproses", vbCritical, "DATA KOSONG"
        Exit Sub
    End If
    Exit Sub

vbError:
    If erTranAll = True Then
        Call objHT.CloseTransactionAll_Err
        Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderDMSNonDSS", "tblSimpan", 2)
        MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
    End If
End Sub

Private Sub SimpanAR80()
On Error GoTo LangsungInsert
    Dim xBatasKr As Double
    Dim xCurrent As Double
    xBatasKr = objO.Fetch_BatasKredit_Ar(Trim(Replace(text(1).text, ">", "")))
    txtRp(1) = xBatasKr
    xCurrent = objO.Fetch_Current_Ar(Trim(Replace(text(1).text, ">", "")))
    txtRp(0) = xCurrent
LangsungInsert:
    If txtRp(1) = 0 Then
        Call objHT.AddNew_Valid_AR_80(Trim(text(1).text), txtRp(0), txtRp(1), 0, Trim(text(64).text))
    Else
        Call objHT.AddNew_Valid_AR_80(Trim(text(1).text), txtRp(0), txtRp(1), Trim(RepRp(Round((txtRp(0) / txtRp(1)) * 100, 2))), Trim(text(64).text))
    End If
End Sub

Private Sub Saving_Log_HO(nStatus As String, isSQl As String)
    Call objHO.AddNew_LogHO(Trim(text(64)), nStatus, "WO", isSQl)
End Sub

Private Sub simpanJOBDTA()
    Dim nilaiKonterAkhir As Integer
    'ISI FIELD
    Call ClearsSql
    newWo = False
    If Trim(text(64)) = "" Then
        text(64) = isiNOWO
        newWo = True
    End If
            
    'If cmdPilih.tag = "No. Booking" Then
        gsSql1(0) = "NOFAK"
        gsSql2(0) = text(78)   'Text(0)
    'End If
    
    gsSql1(1) = "FS_KD_PLG"
    gsSql2(1) = text(1)
    gsSql1(2) = "FS_KD_WIL"
    gsSql2(2) = text(2)
    gsSql1(3) = "CUSTOMER"
    gsSql2(3) = text(3)
    gsSql1(4) = "CUST_ADRES"
    gsSql2(4) = text(4) & " " & text(117)
    gsSql1(5) = "KOTA"
    gsSql2(5) = text(5)
    gsSql1(6) = "TELPON"
    gsSql2(6) = text(6)
    gsSql1(7) = "FS_TH_MOBIL"
    gsSql2(7) = text(9)
    gsSql1(8) = "MDL"
    gsSql2(8) = text(10)
    gsSql1(9) = "MODEL"
    gsSql2(9) = text(11)
    gsSql1(10) = "FMDL"
    gsSql2(10) = text(12)
    gsSql1(11) = "COLOR"
    gsSql2(11) = text(13)
    gsSql1(12) = "NO_POLISI"
    gsSql2(12) = text(14)
    gsSql1(13) = "NO_CHASSIS"
    gsSql2(13) = text(15)
    gsSql1(14) = "DOTAM"
    gsSql2(14) = Replace(text(16), ">", "")
    gsSql1(15) = "NO_ENGINE"
    gsSql2(15) = text(76)
    gsSql1(16) = "FS_KD_DEALER"
    gsSql2(16) = Trim(gKdCbg)
    gsSql1(17) = "FS_NM_DEALER"
    gsSql2(17) = Trim(gNmCbg)
    gsSql1(18) = "FRONT"
    gsSql2(18) = gSA
    gsSql1(19) = "MECH"
    gsSql2(19) = text(50)
    
    gsSql1(99) = "[GROUP]"
    gsSql2(99) = Trim(text(106))
    
    gsSql1(20) = "KM"
    gsSql2(20) = CDbl(text(51))
    gsSql1(21) = "DATE"
    gsSql2(21) = "getdate()"
    gsSql1(22) = "TIME"
    gsSql2(22) = text(66)
    gsSql1(23) = "PART"
    gsSql2(23) = CDbl(text(57))
    gsSql1(24) = "MTRL"
    gsSql2(24) = CDbl(text(58))
    gsSql1(25) = "OIL"
    gsSql2(25) = CDbl(text(59))
    gsSql1(26) = "SBLT"
    gsSql2(26) = CDbl(text(60))
    gsSql1(27) = "LABOR"
    gsSql2(27) = CDbl(text(56))
    gsSql1(28) = "BODY"
    gsSql2(28) = CDbl(text(77))
    gsSql1(29) = "TOTAL"
    gsSql2(29) = CDbl(text(61))
    gsSql1(30) = "REQ1"
    gsSql2(30) = text(23)
    gsSql1(31) = "REQ2"
    gsSql2(31) = text(24)
    gsSql1(32) = "REQ3"
    gsSql2(32) = text(25)
    gsSql1(33) = "REQ4"
    gsSql2(33) = text(26)
    gsSql1(34) = "REQ5"
    gsSql2(34) = text(27)
    For i = 1 To 7
        gsSql1(34 + i) = "REP" & i
        gsSql2(34 + i) = DGrid.TextMatrix(i, 1)
        gsSql1(41 + i) = "RTE" & i
        gsSql2(41 + i) = Replace(IIf(DGrid.TextMatrix(i, 2) = "", 0, DGrid.TextMatrix(i, 2)), ",", ".")
        
        ' Actual Rate
        Dim nActRate As Double
        Set rsTemp = objO.FetchData("SELECT * FROM actrate WITH (NOLOCK) " _
                                 & "WHERE FS_ITEM = '" & DGrid.TextMatrix(i, 1) & "' AND MDL = '" & Trim(text(10)) & "' AND FS_LOG_DEL = '' " _
                                 & "ORDER BY levelMech")
        If Not rsTemp.EOF Then
            gsSql1(48 + i) = "ACTRATE" & i
            gsSql2(48 + i) = Replace(CDbl(rsTemp("RATE")), ",", ".")
        Else
            gsSql1(48 + i) = "ACTRATE" & i
            gsSql2(48 + i) = 0
            gSBE = False
        End If
        nActRate = nActRate + gsSql2(48 + i)
        
        ' GROUP JOB
        gsSql1(55 + i) = "KLPKERJA" & i
        gsSql2(55 + i) = DGrid.TextMatrix(i, 4)
        
        ' NILAI HARGA JOB
        gsSql1(62 + i) = "OFF" & i
        gsSql2(62 + i) = IIf(DGrid.TextMatrix(i, 5) <> "", DGrid.TextMatrix(i, 5), 0)
        
        ' MECH YG MENGERJAKAN PER JOB
        gsSql1(69 + i) = "MEKANIK" & i
        gsSql2(69 + i) = DGrid.TextMatrix(i, 6)
        
        Set rsHlp = objO.FetchData("SELECT [GROUP] FROM MEKANIK WITH (NOLOCK) WHERE MECH = '" & Trim(gsSql2(69 + i)) & "'")
        If Not rsHlp.EOF Then
            gsSql1(76 + i) = "GRP" & i
            gsSql2(76 + i) = rsHlp("GROUP")
        Else
            gsSql1(76 + i) = "GRP" & i
            gsSql2(76 + i) = ""
        End If
    Next
    gsSql1(84) = "NO_WO"
    gsSql2(84) = Trim(text(64))
    gsSql1(85) = "TPSVC"
    gsSql2(85) = IIf(opTpSvc(2).value, "W", IIf(opTpSvc(1).value, "P", ""))
    gsSql1(86) = "KODETIPE"
    gsSql2(86) = IIf(opTpSvc(2).value, text(21), IIf(opTpSvc(1).value, text(21), ""))
    gsSql1(87) = "NOINV"
    gsSql2(87) = text(67)
    gsSql1(88) = "STATUS"
    gsSql2(88) = Trim(text(103)) 'CboStatus
    gsSql1(89) = "PAID"
    gsSql2(89) = IIf(opByr(0).value, "CS", IIf(opByr(1).value, "CR", "T"))
    gsSql1(90) = "REMARK1"
    gsSql2(90) = Txt(54)
    gsSql1(91) = "SUGGEST"
    gsSql2(91) = Txt(55)
    gsSql1(92) = "FS_LOG_ADD"
    gsSql2(92) = gSA
    gsSql1(93) = "CATAT"
    gsSql2(93) = Replace(Txt(53), "'", "`") & " " & Replace(Txt(87), "'", "`")
    gsSql1(94) = "FN_RETURN"
    gsSql2(94) = IIf(ChkReturnOrder.value = Checked, 1, 0)
    gsSql1(95) = "THS"
    gsSql2(95) = IIf(chTHS.value = Checked, "1", "0")
    'gsSql1(82) = "SBE"
    'gsSql2(82) = IIf(gSBE, "SBE", "")
    
    gsSql1(96) = "ANCUST"
    gsSql2(96) = Replace(Trim(text(102)), "'", "`")
    gsSql1(97) = "MOU"
    gsSql2(97) = Trim(text(90))
    gsSql1(98) = "TELPANCUST"
    gsSql2(98) = Trim(text(105))
    gsSql1(100) = "[CODE]"
    gsSql2(100) = IIf(chEM.value = Checked, "1", "0")
    gsSql1(101) = "KD_K"
    gsSql2(101) = IIf(chSBNP.value = Checked, "1", "0")
    
    
    
    If Not woOnProgres Then
        If chPDS = Checked Then
            gsSql1(102) = "RPR_COND"
            gsSql2(102) = "S"
            gsSql1(103) = "DELIV"
            gsSql2(103) = "OKE"
            gsSql1(104) = "PDS"
            gsSql2(104) = "1"
        Else
            gsSql1(102) = "RPR_COND"
            gsSql2(102) = IIf(onlySublet, "S", "")
            gsSql1(103) = "DELIV"
            gsSql2(103) = IIf(onlySublet, "OKE", "")
            gsSql1(104) = "PDS"
            gsSql2(104) = "0"
        End If
        nilaiKonterAkhir = 104
    Else
        If chPDS = Checked Then
            gsSql1(102) = "PDS"
            gsSql2(102) = "1"
        Else
            gsSql1(102) = "PDS"
            gsSql2(102) = "0"
        End If
        nilaiKonterAkhir = 102
    End If
    
    If Log.DmsNonDss = 1 Then
        gsSql1(nilaiKonterAkhir + 1) = "nilaiDPPart"
        gsSql2(nilaiKonterAkhir + 1) = nilaiDPPart
    End If
    ' KURANGI JAM ALOKASI LAMA
    Dim rsLama As ADODB.Recordset
    Dim rsAct As ADODB.Recordset
    Dim hlp As String
    Dim nTotAct As Integer
    
    Call createMekanik
    nMK = 0
    If cmdPilih.Tag = "No. Booking" Then
        Set rsLama = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(text(78)) & "'")
        If Not rsLama.EOF Then
            nTotAct = 0
            For i = 1 To 7
                hlp = "REP" & i
                If rsLama(hlp) <> "" Then
                    nTotAct = nTotAct + (objO.FetchData_Double("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'"))
                End If
            Next
            Call createMekanik
            'Call decresRestTime(rsLama("MECH"), CDbl(nTotAct), 1, Format(rsLama("DATE"), "MM/DD/YYYY"))
        End If
        rsLama.Close
    ElseIf cmdPilih.Tag = "No. WO" Then
        Set rsLama = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "'")
        If Not rsLama.EOF Then
            nTotAct = 0
            For i = 1 To 7
                hlp = "REP" & i
                If rsLama(hlp) <> "" Then
                    nTotAct = nTotAct + (objO.FetchData_Double("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'"))
                End If
            Next
            Call createMekanik
            'Call decresRestTime(rsLama("MECH"), CDbl(nTotAct), 1, Format(rsLama("DATE"), "MM/DD/YYYY"))
        End If
    End If
    
    Set rsTemp = objO.FetchData("SELECT * FROM jobdta WITH (NOLOCK) WHERE no_chassis = '" & Trim(text(15)) & "' AND no_WO = '" & Trim(text(64)) & "'")
    If Not rsTemp.EOF Then
        sSQL = GensSql2("JOBDTA", True) + "WHERE no_chassis = '" & text(15) & "' AND no_WO = '" & Trim(text(64)) & "'"
    Else
        sSQL = GensSql2("JOBDTA", False)
    End If
    Call objHT.UpdateData_Bengkel(sSQL)
    sSQLHO = sSQL
'    UPDATE JAM TERPAKAI MEKANIK
    nTotAct = 0
    For i = 1 To 7
        If DGrid.TextMatrix(i, 1) <> "" Then
            nTotAct = nTotAct + (objO.FetchData_Double("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(text(10)) & "'"))
        End If
    Next
    LoopMk = False
    Call countRestTime(Trim(text(50)), CDbl(nTotAct), 1, Format(text(65), "MM/DD/YYYY"))
    For i = 0 To nMK - 1
        If onTR = True Then
            Call objHT.UpdateData_Bengkel(UpdateMK(i))
        Else
            Call objO.UpdateData(UpdateMK(i))
        End If
    Next i
    
    ' UPDATE CLOCK TUNGGU DILAYANI & CLOCK SA
    Set rsTemp = objO.FetchData("SELECT jam, jamPanggil, DATEDIFF(MINUTE, jam, jamPanggil) as waktu, DATEDIFF(MINUTE, jamPanggil, GETDATE()) as waktuSA FROM T_ANTRIAN WITH (NOLOCK) WHERE loket = '" & gLoket & "' AND noAntrian = '" & gNoAntri & "' AND statusCust = '" & gTpAntri & "'" _
                             & "AND jam >= '" & Format(isnow, "mm/dd/yyyy 00:00:00") & "' AND jam < '" & Format(isnow, "mm/dd/yyyy 23:59:59") & "'")
    If Not rsTemp.EOF Then
        Call objHT.UpdateData_Bengkel("UPDATE JOBDTA SET FS_RC = '', TIME_AT = '" & hitungJam(rsTemp("WAKTU")) & "', TIME_SA = '" & hitungJam(rsTemp("WAKTUSA")) & "' where NO_WO ='" & Trim(text(64)) & "'")
        Call objHT.UpdateData_Bengkel("UPDATE T_ANTRIAN SET NoPol = '" & text(14) & "', NO_WO = '" & Trim(text(64)) & "' WHERE noAntrian = '" & gNoAntri & "' AND StatusCust = '" & gTpAntri & "' AND JAM BETWEEN '" & Format(isnow, "mm/dd/yyyy 00:00:00") & "' AND '" & Format(isnow, "mm/dd/yyyy 23:59:59") & "'")
    End If
    
    ' UPDATE NILAI REGULAR CHECK
    Call objHT.UpdateData_Bengkel("UPDATE JOBDTA SET FS_RC ='', SBE = '' where NO_WO ='" & Trim(text(64)) & "' ")
    Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC <>''")
    If Not rsTemp.BOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Call objHT.UpdateData_Bengkel("UPDATE JOBDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_WO ='" & Trim(text(64)) & "' " _
                 & " AND REP1+ REP2+ REP3+ REP4+ REP5+ REP6+ REP7 like '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%'")
            rsTemp.MoveNext
        Loop
    End If
    
    ' UPDATE NILAI REGULAR CHECK
    Dim rsTemp2 As ADODB.Recordset
    Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC <>''")
    If Not rsTemp.BOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Set rsTemp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND REP LIKE '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%'")
            If Not rsTemp2.EOF Then
                Call objHT.UpdateData_Bengkel("UPDATE JOBDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_WO ='" & Trim(text(64)) & "' ")
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    ' SIMPAN KELENGKAPAN
    For i = 0 To 1
        Select Case i
            Case 0
                '' SPK
                If opLengkap(i) = Checked Then
                    Set rsTemp2 = objO.FetchData("SELECT * FROM JOBLENGKAP WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "'")
                    If rsTemp2.EOF Then
                        Call objHT.InsertData_Bengkel("INSERT INTO JOBLENGKAP (NO_WO, SPKKET) VALUES ('" & Trim(text(64)) & "','" & Trim(text(115)) & "')")
                    Else
                        If Trim(rsTemp2("SPKKet")) <> Trim(text(115 + i)) Then
                            Call objHT.UpdateData_Bengkel("UPDATE JOBLENGKAP SET SPKKet = '" & Trim(text(115 + i)) & "' WHERE NO_WO = '" & Trim(text(64)) & "'")
                        End If
                    End If
                End If
            Case 1
                '' LAINNYA
                If opLengkap(i) = Checked Then
                    Set rsTemp2 = objO.FetchData("SELECT * FROM JOBLENGKAP WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "'")
                    If rsTemp2.EOF Then
                        Call objHT.InsertData_Bengkel("INSERT INTO JOBLENGKAP (NO_WO, LainKet) VALUES ('" & Trim(text(64)) & "','" & Trim(text(115 + i)) & "')")
                    Else
                        If Trim(rsTemp2("LainKet")) <> Trim(text(115 + i)) Then
                            Call objHT.UpdateData_Bengkel("UPDATE JOBLENGKAP SET LainKet = '" & Trim(text(115 + i)) & "' WHERE NO_WO = '" & Trim(text(64)) & "'")
                        End If
                    End If
                End If
        End Select
    Next
End Sub

Private Function hitungJam(mnt As Integer) As String
Dim jamNy As String
Dim mntNy As String

    jamNy = Format(Round((mnt / 60) - 0.5, 0), "0#")
    mntNy = Format(mnt Mod 60, "0#")
    
    hitungJam = jamNy & ":" & mntNy
End Function

Private Sub simpanJOBDTL()

    'ISI FIELD
    Call ClearsSql
            
    'If cmdPilih.tag = "No. Booking" Then
        gsSql1(0) = "NOFAK"
        gsSql2(0) = text(78) 'Text(0)
    'End If
    
    gsSql1(1) = "NO_CHASSIS"
    gsSql2(1) = text(15)
    gsSql1(2) = "NO_WO"
    gsSql2(2) = Trim(text(64))
    For i = 0 To 5
        Select Case i
            Case 5
                gsSql1(3 + i) = "COND" & (i + 1)
                gsSql2(3 + i) = IIf(text(28 + i) = "LAINNYA", "L| " & Trim(bKondisi(i)), text(28 + i))
            Case Else
                gsSql1(3 + i) = "COND" & (i + 1)
                gsSql2(3 + i) = IIf(text(28 + i) = "LAINNYA", "L| " & Trim(bKondisi(i)), text(28 + i))
                
                gsSql1(9 + i) = "WAK" & (i + 1)
                gsSql2(9 + i) = text(34 + i)
                
                gsSql1(14 + i) = "WAE" & (i + 1)
                gsSql2(14 + i) = text(39 + i)
        End Select
    Next
    gsSql1(19) = "BBM"
    gsSql2(19) = IIf(opBBM(0).value, "0", IIf(opBBM(1).value, "1", IIf(opBBM(2).value, "2", IIf(opBBM(3).value, "3", "4"))))
    gsSql1(20) = "GANTIPART"
    gsSql2(20) = IIf(opGanti(0).value, "0", "1")
    gsSql1(21) = "TUNGGUSVC"
    gsSql2(21) = IIf(opSvc(0).value, "0", "1")
    gsSql1(22) = "PEMBAYARAN"
    gsSql2(22) = IIf(opByr(0).value, "CS", IIf(opByr(1).value, "CR", "T"))
    gsSql1(23) = "HEAVYUSE"
    gsSql2(23) = IIf(chUse.value = Checked, 1, 0)
    
    Call objHT.UpdateData_Bengkel("UPDATE CUSTDAT SET HEAVYUSE = '" & gsSql2(23) & "' WHERE no_chassis = '" & text(15) & "'")
    
    Set rsHlp = objO.FetchData("SELECT * FROM tempFU WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "'") ' AND no_WO = '" &Trim(text(64)) & "'
    If Not rsHlp.EOF Then
        gsSql1(24) = "TGLFU"
        gsSql2(24) = Format(rsHlp("TGLFU"), "mm/dd/yyyy")
        gsSql1(25) = "JAMFU"
        gsSql2(25) = rsHlp("JAMFU")
        gsSql1(26) = "TELPFU"
        gsSql2(26) = rsHlp("TELPFU")
        gsSql1(27) = "KETFU"
        gsSql2(27) = rsHlp("KETFU")
    Else
        gsSql1(24) = "TGLFU"
        gsSql2(24) = ""
        gsSql1(25) = "JAMFU"
        gsSql2(25) = ""
        gsSql1(26) = "TELPFU"
        gsSql2(26) = ""
        gsSql1(27) = "KETFU"
        gsSql2(27) = ""
    End If
    rsHlp.Close
    
    gsSql1(28) = "VERIFIKATOR"
    gsSql2(28) = gVerifikator
    gsSql1(29) = "DATESERAH"
    gsSql2(29) = Format(text(88), "MM/DD/YYYY")
    gsSql1(30) = "TIMESERAH"
    gsSql2(30) = Trim(text(89))
    gsSql1(31) = "DATANG"
    gsSql2(31) = Trim(text(107))
    gsSql1(32) = "DTLDATANG"
    gsSql2(32) = Trim(text(52))
    gsSql1(33) = "KARAKTER"
    gsSql2(33) = Trim(cboPelanggan)
    gsSql1(34) = "DTLKARAKTER"
    gsSql2(34) = Trim(text(75))
    gsSql1(35) = "WASHING"
    gsSql2(35) = IIf(opWash(0), 1, 0)
    
    For i = 1 To 5
        gsSql1(35 + i) = "CMPL" & i
        'gsSql2(35 + i) = text(96 + i)
        gsSql2(35 + i) = Text1(-1 + i)
    Next
    
    Set rsTemp = objO.FetchData("SELECT * FROM jobdtl WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "' AND no_WO = '" & Trim(text(64)) & "'")
    If Not rsTemp.EOF Then
        If (TglDDMMYYYY(rsTemp("DATESERAH")) <> Trim(text(88))) Or (Trim(rsTemp("TIMESERAH")) <> Trim(text(89))) Then
            '' UPDATE JAM PENYERAHAN
            Dim rsHlp2 As ADODB.Recordset
            Set rsHlp2 = objO.FetchData("SELECT RUBAH FROM JOBSERAH WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' ORDER BY RUBAH DESC")
            If Not rsHlp2.EOF Then
                Call objHT.InsertData_Bengkel("INSERT INTO JOBSERAH VALUES ('" & Trim(text(64)) & "','" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN") & "','" & Format(text(88) + " " + Trim(text(89)), "MM/DD/YYYY HH:NN") & "','" & CDbl(rsHlp2("RUBAH")) + 1 & "')")
            End If
        End If
        sSQL = GensSql2("JOBDTL", True) + "WHERE no_chassis = '" & text(15) & "' AND no_WO = '" & Trim(text(64)) & "'"
    Else
        Call objHT.InsertData_Bengkel("INSERT INTO JOBSERAH VALUES ('" & Trim(text(64)) & "','" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN") & "','" & Format(text(88) + " " + Trim(text(89)), "MM/DD/YYYY HH:NN") & "','1')")
        sSQL = GensSql2("JOBDTL", False)
    End If
    Call objHT.UpdateData_Bengkel(sSQL)
    Call objHT.DeleteData_Bengkel("DELETE tempFU WHERE no_chassis = '" & text(15) & "' AND no_WO = '" & Trim(text(64)) & "'")
End Sub

Private Sub simpanJOBKoord()

    'ISI FIELD
    Call ClearsSql
            
    'If cmdPilih.tag = "No. Booking" Then
        gsSql1(0) = "NOFAK"
        gsSql2(0) = text(78)    'Text(0)
    'End If
    
    gsSql1(1) = "NO_CHASSIS"
    gsSql2(1) = text(15)
    gsSql1(2) = "NO_WO"
    gsSql2(2) = Trim(text(64))
    
    Call objHT.DeleteData_Bengkel("DELETE jobKoord WHERE no_chassis = '" & text(15) & "' AND no_WO = '" & Trim(text(64)) & "'")
    
    For i = 0 To nk - 1
        gsSql1(3) = "KOORDX"
        gsSql2(3) = KorX(i)
        gsSql1(4) = "KOORDY"
        gsSql2(4) = KorY(i)
        
'        Set rsTemp = objO.FetchData("SELECT * FROM jobKoord WHERE no_chassis = '" & text(15) & "' AND no_WO = '" &Trim(text(64)) & "' AND KOORDX = '" & KorX(i) & "' AND KOORDY = '" & KorY(i) & "'")
'        If rsTemp.EOF Then
            sSQL = GensSql2("JOBKOORD", False)
'        End If
        Call objHT.UpdateData_Bengkel(sSQL)
    Next
    
    ' SIMPAN STALL
    Dim NoFak As String
    Set rsTemp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE keterangan like '%" & Trim(text(14)) & " |%' ")  'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'")
    If Not rsTemp.EOF Then
    
        If cmdPilih.Tag = "No. WO" Then
            Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(text(64)) & "|%'") 'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'
            Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(text(78)) & "|%'") 'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'
            Call objHT.InsertData_Bengkel("INSERT INTO STALL SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [Group], '" & Trim(text(50)) & "', Flag, '" & text(14) & " |" & Trim(text(64)) & "| " & text(3) & "' AS KETERANGAN,adjust FROM tempStall WITH (NOLOCK) " _
                         & "WHERE NOT EXISTS (SELECT * FROM STALL WITH (NOLOCK) WHERE keterangan like '%|" & Trim(text(64)) & "|%') AND " _
                         & "keterangan like '%|" & Trim(text(64)) & "|%'")
            Call objHT.DeleteData_Bengkel("DELETE TEMPSTALL WHERE KETERANGAN LIKE '%|" & Trim(text(64)) & "|%'") 'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'
            Call objHT.DeleteData_Bengkel("DELETE TEMPSTALL WHERE KETERANGAN LIKE '%|" & Trim(text(78)) & "|%'") 'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'
        ElseIf cmdPilih.Tag = "No. Booking" Then
            'If text(78).text <> "" Then  'By No.Booking
                Call objHT.UpdateData_Bengkel("UPDATE STALL SET NOSTALL = B.NOSTALL, TANGGAL = B.TANGGAL, JAMMULAI = B.JAMMULAI, JAMSELESAI = B.JAMSELESAI, RATE = " & Replace(CDbl(ntRate), ",", ".") & ", " _
                         & "[GROUP] = B.[GROUP], MECH = '" & Trim(text(50)) & "', FLAG = 'N',adjust = B.ADJUST, KETERANGAN = '" & Trim(text(14)) & " |" & Trim(text(64)) & "| " & text(3) & "' " _
                         & "FROM STALL A WITH (NOLOCK) INNER JOIN TEMPSTALL B WITH (NOLOCK) ON " _
                         & "LEFT(A.KETERANGAN, CHARINDEX(' |',A.KETERANGAN)) = LEFT(B.KETERANGAN, CHARINDEx (' |',B.KETERANGAN)) and B.TANGGAL = CONVERT(varchar(12),GETDATE(),101) " _
                         & "WHERE A.KETERANGAN LIKE '%|" & Trim(text(78)) & "|%'")
                         
                         'masalah @STALL  revisi 25/07/2012
                         'SUBSTRING(A.KETERANGAN, CHARINDEX(' |',A.KETERANGAN)+2,CHARINDEX('| ',A.KETERANGAN)-CHARINDEX(' |',A.KETERANGAN)-2) = " _
                         & "SUBSTRING(B.KETERANGAN, CHARINDEX(' |',B.KETERANGAN)+2,CHARINDEX('| ',B.KETERANGAN)-CHARINDEX(' |',B.KETERANGAN)-2)
                         
                Call objHT.DeleteData_Bengkel("DELETE TEMPSTALL WHERE KETERANGAN LIKE '%|" & Trim(text(78)) & "|%' ") 'AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'
            'Else   ' By No.Rangka
        ElseIf cmdPilih.Tag = "No. Rangka" Then
                Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(text(64)) & "|%'") 'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'")
                Call objHT.InsertData_Bengkel("INSERT INTO STALL SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [Group], '" & Trim(text(50)) & "', 'N', '" & text(14) & " |" & Trim(text(64)) & "| " & text(3) & "' AS Keterangan,adjust FROM tempStall WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM STALL WITH (NOLOCK) WHERE keterangan like '%|" & Trim(text(64)) & "|%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "') AND keterangan like '%| " & text(3) & "%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'")
                Call objHT.DeleteData_Bengkel("DELETE TEMPSTALL WHERE KETERANGAN LIKE '%| " & text(3) & "%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'")
            'End If
        End If
    End If
End Sub

Private Sub simpanJOBPART()
Dim nPart As String
Dim nSyaratBayar As String
Dim rsSb As ADODB.Recordset
Set rsSb = objO.FetchDataACC("SELECT TOPGRP as SyaratBayar FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(text(1)) & "'")
If rsSb.RecordCount > 0 Then
    nSyaratBayar = rsSb.Fields(0).value
    Set rsSb = Nothing
Else
    nSyaratBayar = "D0"
End If

Dim rsSOno As ADODB.Recordset
Set rsSOno = objO.FetchDataACC("Select top 1 Nomor from SOORD WITH (NOLOCK) WHERE NOSO = '" & Trim(text(64)) & "' Order By Nomor DESC")

Dim iNo As Double

If rsSOno.RecordCount > 0 Then
    iNo = CDbl(rsSOno.Fields(0).value + 1)
    Set rsSOno = Nothing
Else
    If Trim(text(78)) <> "" Then
        Set rsSOno = objO.FetchDataACC("Select top 1 Nomor from SOORD WITH (NOLOCK) WHERE NOSPK = '" & Trim(text(78)) & "' Order By Nomor DESC")
        If rsSOno.RecordCount > 0 Then
            iNo = CDbl(rsSOno.Fields(0).value + 1)
            Set rsSOno = Nothing
        Else
            iNo = 1
        End If
    Else
        iNo = 1
    End If
    
End If

'Call objO.DeleteDataACC("DELETE SOORD WHERE AddJob = 0 and (NOSO = '" & Trim(text(64)) & "' OR NOSPK = '" & Trim(text(78)) & "') AND ((IsNull(QtyKirimBL,0) <> 0 OR IsNull(QtyKirimBI,0) <> 0 OR [TO] <> 0) Or BO not in ('','CANCEL')) ")
'if fill
    Dim rsSO As ADODB.Recordset
    If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
        If Trim(text(64)) <> "" And Trim(text(78)) <> "" Then
            Set rsSO = objO.FetchDataACC("Select KodeItem + NamaBySA + convert(varchar(10),Nomor) as Kode from SOORD WITH (NOLOCK) WHERE AddJob = 0 and (NOSO = '" & Trim(text(64)) & "' or NOSPK = '" & Trim(text(78)) & "') Order By Nomor DESC")
        ElseIf Trim(text(64)) <> "" Then
            Set rsSO = objO.FetchDataACC("Select KodeItem + NamaBySA + convert(varchar(10),Nomor) as Kode from SOORD WITH (NOLOCK) WHERE AddJob = 0 and (NOSO = '" & Trim(text(64)) & "') Order By Nomor DESC")
        Else
            Set rsSO = objO.FetchDataACC("Select KodeItem + NamaBySA + convert(varchar(10),Nomor) as Kode from SOORD WITH (NOLOCK) WHERE AddJob = 0 and (NOSPK = '" & Trim(text(78)) & "') Order By Nomor DESC")
        End If
    End If

'Set rsSO = objO.FetchDataACC("Select convert(varchar(10),Nomor) as Kode from SOORD WHERE AddJob = 0 and (NOSO = '" & Trim(text(64)) & "' OR NOSPK = '" & Trim(text(78)) & "') Order By Nomor DESC")

    'ISI FIELD
    Call ClearsSql
            
        gsSql1(0) = "NOFAK"
        gsSql2(0) = text(78)
    
    gsSql1(1) = "NO_CHASSIS"
    gsSql2(1) = text(15)
    gsSql1(2) = "NO_WO"
    gsSql2(2) = Trim(text(64))
    
    ' //???// NOTE KE SPAREPART //???//
    ' HAPUS PART YG TIDAK ADA DI BOOK_PART
    Dim xNomor As String
    xNomor = ""
    Call objHT.DeleteData_Bengkel("DELETE JOBPART WHERE NO_WO = '" & Trim(text(64)) & "' and ADDJOB = 0 ")
    'If isUpdateFill = True Then
    '    Call objHT.DeleteData_ACC("DELETE SOORD WHERE NOSO = '" & Trim(text(64)) & "' and isnull(QtyKirimBL,0) + isnull(QtyKirimBi,0) = 0  and ADDJOB = 0")
    'End If
    'If isUpdateFill = False And Trim(text(78)) <> "" Then
    '    Call objHT.DeleteData_ACC("DELETE SOORD WHERE (NOSO = '" & Trim(text(64)) & "' Or NoSPK = '" & Trim(text(78)) & "') and isnull(QtyKirimBL,0) + isnull(QtyKirimBi,0) = 0  and ADDJOB = 0")
    'End If
    'createTblPart
    If Not rsGridPart.State = adStateClosed Then
        If rsGridPart.RecordCount > 0 Then
        'If Not rsGridPart.State = adStateClosed Then
            rsGridPart.MoveFirst
            If Not rsGridPart.EOF Then
                Dim nDelIndex As String
                nDelIndex = GetDelIndex("SV-")
                rsGridPart.MoveFirst
                i = 0
                nPart = ""
                Do While Not rsGridPart.EOF
                    
                    gsSql1(3) = "KODE_ITEM"
                    gsSql2(3) = rsGridPart("KODEITEM")
                    gsSql1(4) = "NAMA_ITEM"
                    gsSql2(4) = rsGridPart("NAMAITEM")
                    gsSql1(5) = "QTY"
                    gsSql2(5) = rsGridPart("QTY")
                    gsSql1(6) = "GANTI"
                    gsSql2(6) = IIf(rsGridPart("GANTI"), "Y", "N")
                    gsSql1(7) = "BAWA"
                    gsSql2(7) = IIf(rsGridPart("BAWA"), "Y", "N")
                    gsSql1(8) = "PARTBAWA"
                    If rsGridPart("BAWA") Then
                        gsSql2(8) = text(69 + i)
                        i = i + 1
                    Else
                        gsSql2(8) = ""
                    End If
                    gsSql1(9) = "TGLPASANG"
                    gsSql2(9) = Format(text(65), "MM/DD/YYYY")
                    
                    ' AMBIL QTY SUPPLY DARI SPAREPART
                        Set rsTemp = objO.FetchDataACC("SELECT INTRN.KODEITEM, SUM(QTY) AS JML FROM INTRN WITH (NOLOCK) INNER JOIN INMST WITH (NOLOCK) ON INTRN.KODEITEM=INMST.KODEITEM " _
                                                    & "WHERE SBRBUKTI = '" & Trim(text(0)) & "' AND INTRN.KODEITEM = '" & rsGridPart("KODEITEM") & "' AND INMST.KODEGDG = '04' GROUP BY INTRN.KODEITEM HAVING SUM(QTY) <> 0")
                        If Not rsTemp.EOF Then
                            gsSql1(10) = "QTYSUPPLY"
                            gsSql2(10) = CDbl(rsTemp("JML"))
                        Else
                            gsSql1(10) = "QTYSUPPLY"
                            gsSql2(10) = 0
                        End If
                    
                    Dim nDis As Integer
                    nDis = 0
                    If opTpSvc(2) And text(21) <> "" Then
                        ' UTK SSC PART TIDAK MINTA KE SPAREPART
                    Else
                        If opTpSvc(1) Then
                            nDis = objO.FetchData_String("SELECT DISCPART FROM HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & text(21) & "'")
                        End If
                        If rsGridPart("GANTI") = "Y" And rsGridPart("BAWA") = "N" Then
                            nPart = nPart + (rsGridPart("KODEITEM") & "^" & rsGridPart("NAMAITEM") & "^" & rsGridPart("QTY") & "^" & nDis & "|")
                        End If
                    End If
                    gsSql1(11) = "NAMABYSA"
                    gsSql2(11) = rsGridPart("NAMABYSA")
                    gsSql1(12) = "nID"
                    gsSql2(12) = iNo + 1 'rsGridPart("nID")
                    
                    'iNo = rsGridPart("nID")
                    'xNomor = xNomor & "," & rsGridPart("nID")
                    sSQL = GensSql2("JOBPART", False)
                    Call objHT.UpdateData_Bengkel(sSQL)
                    If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                        If isUpdateFill = False Then
                            If Trim(text(78)) <> "" Then
                                If rsSO.RecordCount > 0 Then
                                    rsSO.MoveFirst
                                    rsSO.Find "Kode = '" & rsGridPart("KODEITEM") + rsGridPart("NAMABYSA") + Trim(rsGridPart("nID")) & "'", , adSearchForward, 1
                                    'rsSO.Find "Kode = '" & Trim(rsGridPart("nID")) & "'", , adSearchForward, 1
                                    If rsSO.EOF Then
                                        xNomor = xNomor & "," & iNo
                                        Call objHT.InsertData_ACC("Insert Into SOORD " & _
                                            "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                                            "Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & iNo & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                                            '"Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & rsGridPart("nID") & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                                    Else
                                        xNomor = xNomor & "," & rsGridPart("nID")
                                        Call objHT.UpdateData_ACC("Update SOORD Set noSO = '" & Trim(text(64)) & "', KodeUser ='" & gSA & "',TglPerlu =getdate(), Status = '" & gSA & "', TglRange ='" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "' " & _
                                            " Where Nomor = " & rsGridPart("nID") & " And noSPK = '" & Trim(text(78)) & "' And KodeItem = '" & rsGridPart("KODEITEM") & "'")
                                    End If
                                Else
                                    xNomor = xNomor & "," & iNo
                                    Call objHT.InsertData_ACC("Insert Into SOORD " & _
                                        "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                                        "Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & iNo & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                                        '"Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & rsGridPart("nID") & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                                End If
                            Else
                                xNomor = xNomor & "," & iNo
                                Call objHT.InsertData_ACC("Insert Into SOORD " & _
                                    "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                                    "Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & iNo & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                                    '"Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & rsGridPart("nID") & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                            End If
                        Else
                            If rsSO.RecordCount > 0 Then
                                rsSO.MoveFirst
                                rsSO.Find "Kode = '" & rsGridPart("KODEITEM") + rsGridPart("NAMABYSA") + Trim(rsGridPart("nID")) & "'", , adSearchForward, 1
                                'rsSO.Find "Kode = '" & Trim(rsGridPart("nID")) & "'", , adSearchForward, 1
                                If rsSO.EOF Then
                                    xNomor = xNomor & "," & iNo
                                    Call objHT.InsertData_ACC("Insert Into SOORD " & _
                                        "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                                        "Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & iNo & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                                Else
                                    xNomor = xNomor & "," & rsGridPart("nID")
                                End If
                            Else
                                xNomor = xNomor & "," & iNo
                                Call objHT.InsertData_ACC("Insert Into SOORD " & _
                                        "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                                        "Values ('" & Trim(text(64)) & "', '" & rsGridPart("KODEITEM") & "', '" & rsGridPart("NAMAITEM") & "', '" & rsGridPart("QTY") & "','1', '" & rsGridPart("SATITEM") & "', '" & Trim(text(78)) & "', '" & gSA & "', Getdate(),'" & iNo & "','" & rsGridPart("NAMABYSA") & "', '0',getdate(),'" & gSA & "','" & TglNul2(isnow) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "HH:MM:SS") & "')")
                            End If
                        End If
                    End If
                    rsGridPart.MoveNext
                    iNo = iNo + 1
                Loop
                If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                Dim sqlOrder As String
                Dim sqlBooking As String
                
                sqlOrder = "SELECT NoSo FROM sohdr WITH (NOLOCK) WHERE NoSo = '" & Trim(text(64)) & "'"
                sqlBooking = "SELECT NoSo FROM sohdr WITH (NOLOCK) WHERE NoSo = '" & Trim(text(78)) & "'"
                    If isUpdateFill = False Or objHT.FetchDataACC(sqlOrder).RecordCount = 0 Then
                        If Trim(text(78)) <> "" And objHT.FetchDataACC(sqlBooking).RecordCount > 0 Then
                            Call objHT.UpdateData_ACC("Update SOHDR Set NOSO = '" & Trim(text(64)) & "', TglPerlu = getdate() Where NOSPK = '" & Trim(text(78)) & "'")
                        Else
                            Call objHT.InsertData_ACC("insert into SOHDR " & _
                            "(NOSO,TglSO, TypeTrn, TglPerlu, KodeLgn, NamaKirim, AlamatKirim1, AlamatKirim2,AlamatKirim3,SyaratBayar, KODESUMBER, DELINDEX, TglEntry, KodeUser, NoSPK ) " & _
                            "values ('" & Trim(text(64)) & "', Getdate(), 'A',getdate(), '" & Trim(text(1)) & "', '" & Left(Trim(text(3)), 40) & "', '" & Left(Trim(text(4)), 40) & "', '" & Left(Trim(text(117)), 40) & "', '" & Trim(text(5)) & "', '" & nSyaratBayar & "', 'SV', '" & nDelIndex & "', getdate(), '" & gSA & "', '" & Trim(text(78)) & "' )")
                        End If
                    Else
                        Call objHT.UpdateData_ACC("Update SOHDR Set TglPerlu = getdate() Where NOSO = '" & Trim(text(64)) & "'")
                    End If
                    If isUpdateFill = False And Trim(text(78)) <> "" And rsSO.RecordCount > 0 Then
                        Call objHT.UpdateData_ACC("Update SOORD Set NOSO = '" & Trim(text(64)) & "', TglPerlu = getdate() Where NOSPK = '" & Trim(text(78)) & "'")
                    End If
                    If xNomor = "" Then
                    
                    Else
                       Call objHT.UpdateData_ACC("Delete SOORD Where NOSO = '" & Trim(text(64)) & "' and isnull(tglsupply,'') = '' and ADDJOB = 0 and  BO <> 'CANCEL' and Nomor not in (" & Right(xNomor, Len(xNomor) - 1) & ")")
                    End If
                    
                End If
            End If
        End If
    Else
        If partGridFill.EOF Then
            'jika tidak ada data dalam grid
            If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                Call objHT.UpdateData_ACC("Delete SOORD Where NOSO = '" & Trim(text(64)) & "' and isnull(tglsupply,'') = '' and ADDJOB = 0 and  BO <> 'CANCEL' ")
            End If
            
        End If
    End If
       
    
    ' INPUT SUBLET
    Set rsTemp = objO.FetchData("SELECT * FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE no_Chassis = '" & Trim(text(15)) & "' AND NO_WO = '" & IIf(cmdPilih.Tag = "No. WO", Trim(text(64)), "") & "'")
    If Not rsTemp.EOF Then
        gPO = IIf(gPO <> "", gPO, isiNoPoSublet)
        Call objHT.DeleteData_Bengkel("DELETE jobSUBLET WHERE no_wo = '" & Trim(text(64)) & "' AND no_Chassis = '" & Trim(text(15)) & "' AND NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
        Call objHT.InsertData_Bengkel("INSERT INTO JOBSUBLET SELECT DISTINCT '" & gPO & "', '" & Trim(text(64)) & "', '" & text(67) & "', no_chassis, nofak, kodeItem, NamaBarang, kodeSatuan, Qty, Harga, nonota, keterangan,FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, '', '',estwaktu " _
                     & "FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE no_wo = '" & Trim(text(64)) & "' AND no_Chassis = '" & Trim(text(15)) & "') AND " _
                     & "no_Chassis = '" & Trim(text(15)) & "' AND NO_WO = '" & IIf(cmdPilih.Tag <> "No. WO", "", Trim(text(64))) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
        Call objHT.DeleteData_Bengkel("DELETE TEMPJOBSUBLET WHERE no_wo = '" & IIf(cmdPilih.Tag <> "No. WO", "", Trim(text(64))) & "' AND no_Chassis = '" & text(15) & "'  AND NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
    Else
        Call objHT.DeleteData_Bengkel("DELETE jobSUBLET WHERE no_wo = '" & Trim(text(64)) & "' AND no_Chassis = '" & Trim(text(15)) & "' AND NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
    End If
End Sub

Private Function isiNoPoSublet() As String
    Set rsHlp = objO.FetchData("SELECT DISTINCT NOPO FROM JOBSUBLET WITH (NOLOCK) WHERE SUBSTRING(NOPO,4,4) = '" & Format(Right(Year(isnow), 2), "00") & Format(Month(isnow), "00") & "' ORDER BY NOPO DESC")
    If Not rsHlp.EOF Then
        isiNoPoSublet = "GR." & Format(Right(Year(isnow), 2), "00") & Format(Month(isnow), "00") & Format(CDbl(Right(rsHlp("NOPO"), 5)) + 1, "00000")
    Else
        isiNoPoSublet = "GR." & Format(Right(Year(isnow), 2), "00") & Format(Month(isnow), "00") & Format(1, "00000")  'Format(Right(Year(isnow), 2), "00")
    End If
End Function

Public Sub hitungEstimasiPenyerahan()
          Dim rsJob As ADODB.Recordset
          Dim rsSublet As ADODB.Recordset
          Dim totalJam As Integer
          Dim totalSublet As Integer
          Dim totalJob As Double
          Dim hasilJam As String
          Dim totalMenit As Integer
          Dim keterangan As String
          Dim jamTemSkrg As Integer
          Dim menitTemSkrg As Integer
          Dim jamMulai As String
          
1         If woOnProgres Then
2             jamMulai = Trim(text(66))
          ElseIf woOnProgres = False And Trim(text(78)) <> "" Then
              jamMulai = Trim(text(66))
3         Else
4             jamMulai = Trim(text(86))
5         End If
          
6         keterangan = ""
7         totalSublet = 0

8         If Trim(text(78)) <> "" Then
9            Set rsSublet = objO.FetchData("SELECT ESTWAKTU FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE no_Chassis = '" & Trim(text(15)) & "' AND no_WO = '" & Trim(text(78)) & "'")
10        Else
11            Set rsSublet = objO.FetchData("SELECT ESTWAKTU FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE no_Chassis = '" & Trim(text(15)) & "' AND no_WO = '" & Trim(text(64)) & "'")
12        End If
13        If Not rsSublet.EOF Then
14            rsSublet.MoveFirst
15            Do While Not rsSublet.EOF
16                totalSublet = totalSublet + hitungTotalMenit(rsSublet("ESTWAKTU"))
17            rsSublet.MoveNext
18            Loop
19        End If

20        totalJob = 0
21        If Trim(text(64)) <> "" Then
22            keterangan = Trim(text(14)) & " |" & Trim(text(64))
23        Else
24            keterangan = Trim(text(14)) & " |"
25        End If
          
26        Set rsJob = objO.FetchData("SELECT adjust,jammulai,TANGGAL FROM tempstall WITH (NOLOCK) WHERE flag NOT IN ('k','x') and keterangan like '" & keterangan & "%'")
27        If Not rsJob.EOF Then
28            jamTemSkrg = CInt(Left(Trim(rsJob("adjust")), InStr(rsJob("adjust"), ":") - 1))
29            menitTemSkrg = Mid(Trim(rsJob("adjust")), InStr(Trim(rsJob("adjust")), ":") + 1, Len(Trim(rsJob("adjust"))) - InStr(Trim(rsJob("adjust")), ":"))
30            totalJob = (jamTemSkrg * 60) + menitTemSkrg
31            jamMulai = Trim(rsJob("jammulai"))
32            If Format(rsJob("tanggal"), "DD-MM-YYYY") > Trim(text(65)) Then
33              jamMulai = Trim(text(66))
34            End If
35        Else
36            Set rsJob = objO.FetchData("SELECT adjust,jammulai,TANGGAL FROM stall WITH (NOLOCK) WHERE flag NOT IN ('k','x') and keterangan like '" & keterangan & "%'")
37            If Not rsJob.EOF Then
38                jamTemSkrg = CInt(Left(Trim(rsJob("adjust")), InStr(rsJob("adjust"), ":") - 1))
39                menitTemSkrg = Mid(Trim(rsJob("adjust")), InStr(Trim(rsJob("adjust")), ":") + 1, Len(Trim(rsJob("adjust"))) - InStr(Trim(rsJob("adjust")), ":"))
40                totalJob = (jamTemSkrg * 60) + menitTemSkrg
41                jamMulai = Trim(rsJob("jammulai"))
42                  If Format(rsJob("tanggal"), "DD-MM-YYYY") > Trim(text(65)) Then
43                      jamMulai = Trim(text(66))
44                  End If
45            End If
46        End If
          
47        If chEM = Checked Then
48            totalJob = 0.5 * 60
49        ElseIf chSBNP = Checked Then
50            totalJob = 1 * 60
      '        If Trim(text(114)) <> "" Then
      '            totalJob = CDbl(CDbl(Trim(text(114))) * 60)
      '        End If
51        End If
          
52        totalMenit = CInt(totalSublet) + CInt(totalJob) + CInt(Log.estInv)
          'LAKUKAN PENGECEKAN JIKA TANGGAL DAN JAM
53        hasilJam = hitungTotalPenyerahan(totalMenit, jamMulai, text(65))
          
54        text(63) = Left(hasilJam, InStr(hasilJam, "|") - 1)
55        text(62) = TglDDMMYYYY(Right(hasilJam, Len(hasilJam) - InStr(hasilJam, "|")))
          
56        If Not woOnProgres And Trim(text(64)) = "" Then
57            text(89) = Left(hasilJam, InStr(hasilJam, "|") - 1)
58            text(88) = TglDDMMYYYY(Right(hasilJam, Len(hasilJam) - InStr(hasilJam, "|")))
59        End If
          
End Sub

Private Function panggilCustHO(nRangka As String, rs As ADODB.Recordset) As Boolean
    text(1) = rs("FS_KD_PLG")
    text(2) = rs("FS_KD_WIL")
    text(3) = rs("CUSTOMER")
    If Len(Trim(Nul(rs("CUST_ADRES")))) > 60 Then
        text(4) = Left(rs("CUST_ADRES"), 60)
        text(117) = Mid(rs("CUST_ADRES"), 61, 60)
    Else
        text(4) = Left(rs("CUST_ADRES"), 60)
    End If
    text(5) = rs("KOTA")
    text(6) = Replace(rs("TELPON"), " ", "")
    text(7) = rs("GSM")
    text(8) = rs("CDMA")
    text(9) = rs("FS_TH_MOBIL")
    text(10) = rs("MDL")
    text(11) = rs("MODEL")
    text(12) = rs("FMDL")
    text(13) = rs("COLOR")
    text(14).Mask = ""
    text(14) = rs("NO_POLISI")
    text(15) = rs("NO_CHASSIS")
    text(16).Mask = ""
    text(16) = TglDDMMYYYY(rs("DOTAM"))
    text(16).Mask = "99-99-9999"
    text(20) = Format(CDbl(rs("FS_KM_AKHIR")), "0#,#")
    If Trim(text(20)) = "" Then text(20) = 0
    chUse = IIf(rs("HEAVYUSE") = 1, Checked, Unchecked)
    text(76) = rs("NO_ENGINE")
    text(79).Mask = ""
    text(79) = rs("DDATE")
    text(118).Mask = ""
    text(118) = TglDDMMYYYY(rs("DDATE"))
    text(118).Mask = "99-99-9999"
    text(81) = rs("FS_KD_DEALER")
    text(80) = rs("FS_NM_DEALER")
    text(90) = cekMOU(rs("MDL"), nRangka)
    cmdLengkap.Visible = False
    If Trim(text(90)) <> "" Then
        cmdLengkap.Visible = True
    End If
    panggilCustHO = True
End Function

Private Sub btnBatal_Click(Index As Integer)
    frmEvent.Visible = False
End Sub

Private Sub btnCancel_Click()
    frmLengkap.Visible = False
End Sub

Private Sub btnSimpan_Click()
Dim bisaSimpan As Boolean

    bisaSimpan = True
    If Trim(text(115)) = "" And Trim(text(116)) = "" Then
        bisaSimpan = False
    End If
    
    If bisaSimpan Then
        frmLengkap.Visible = False
    End If
End Sub

Private Sub cboPelanggan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub chEM_Click()
    If chEM = Checked Then
        If chSBNP = Unchecked Then
            If chEM = Checked Then
                'Text(86) = Left(hitungWaktu(1, Text(66), Text(65)), 5)
                text(63) = Left(hitungWaktu(1, text(86), text(65)), 5)
                Call hitungEstimasiPenyerahan
            Else
                If text(86) <> ">>:>>" And text(65) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(ntRate), text(86), text(65))
                    'Text(86) = Left(nWKT, 5) 'text(66)
                    text(63) = Left(nWKT, 5) 'text(66)
                    Call hitungEstimasiPenyerahan
                End If
            End If
        Else
            MsgBox "Pilih tipe salah satu EM / SBE NP", vbCritical, "Tipe Service"
            chEM = Unchecked
        End If
    End If
    Call hitungEstimasiPenyerahan
End Sub

Private Sub chEM_Validate(Cancel As Boolean)
    If chEM = Checked Then
        If chSBNP = Unchecked Then
            If chEM = Checked Then
                'Text(86) = Left(hitungWaktu(1, Text(66), Text(65)), 5)
                text(63) = Left(hitungWaktu(1, text(86), text(65)), 5)
                Call hitungEstimasiPenyerahan
            Else
                If text(86) <> ">>:>>" And text(65) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(ntRate), text(86), text(65))
                    'Text(86) = Left(nWKT, 5) 'text(66)
                    text(63) = Left(nWKT, 5) 'text(66)
                    Call hitungEstimasiPenyerahan
                End If
            End If
        Else
            'MsgBox "Pilih tipe salah satu EM / SBE NP", vbCritical, "Tipe Service"
            chEM = Unchecked
        End If
    End If
    Call hitungEstimasiPenyerahan
End Sub

Private Sub chSBNP_Click()
    If chSBNP = Checked Then
        If chEM = Unchecked Then
            If chSBNP = Checked Then
                'Text(86) = Left(hitungWaktu(1, Text(66), Text(65)), 5)
                text(63) = Left(hitungWaktu(1, text(86), text(65)), 5)
                Call hitungEstimasiPenyerahan
            Else
                If text(66) <> ">>:>>" And text(65) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(ntRate), text(66), text(65))
                    'Text(86) = Left(nWKT, 5) 'text(66)
                    text(63) = Left(nWKT, 5) 'text(66)
                    Call hitungEstimasiPenyerahan
                End If
            End If
        Else
            MsgBox "Pilih tipe salah satu EM / SBE NP", vbCritical, "Tipe Service"
            chSBNP = Unchecked
        End If
    End If
    Call hitungEstimasiPenyerahan
End Sub

Private Sub chSBNP_Validate(Cancel As Boolean)
    If chSBNP = Checked Then
        If chEM = Unchecked Then
            If chSBNP = Checked Then
                'Text(86) = Left(hitungWaktu(1, Text(66), Text(65)), 5)
                text(63) = Left(hitungWaktu(1, text(86), text(65)), 5)
                Call hitungEstimasiPenyerahan
            Else
                If text(66) <> ">>:>>" And text(65) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(ntRate), text(66), text(65))
                    'Text(86) = Left(nWKT, 5) 'text(66)
                    text(63) = Left(nWKT, 5) 'text(66)
                    Call hitungEstimasiPenyerahan
                End If
            End If
        Else
            'MsgBox "Pilih tipe salah satu EM / SBE NP", vbCritical, "Tipe Service"
            chSBNP = Unchecked
        End If
    End If
    Call hitungEstimasiPenyerahan
End Sub

Private Sub chPDS_Click()
    If chPDS = Checked Then
        'CboStatus = "INT"
        text(103) = "AFFPPN"
        text(23) = "PDS - PRE DELIVERY SERVICE"
        text(44) = "PDS"
        text(46) = "PRE DELIVERY SERVICE"
        opByr(2).value = True
        opByr(0).Enabled = False
        opByr(1).Enabled = False
        Jrate = 0
        cmdStall.Tag = ""
        If SSTab1.Enabled Then
            text(46).SetFocus
        End If
                
        For i = 1 To 8
            If DGrid.TextMatrix(i, 4) = "SBI" Or DGrid.TextMatrix(i, 4) = "SBE" Or DGrid.TextMatrix(i, 4) = "GRP" Then
                MsgBox "Pekerjaan PDS, tidak boleh dikerjakan bersamaan dengan Job lain (SBI/SBE/GRP) !", vbInformation + vbOKOnly, "Validasi PDS"
                chPDS.value = 0
                Exit Sub
            End If
        Next i
        
    Else
        'CboStatus = "EXT"
        text(103) = "EXT"
        text(23) = ""
        text(44) = ""
        text(46) = ""
        opByr(0).value = True
        opByr(0).Enabled = True
        opByr(1).Enabled = True
        Jrate = 0
        'cmdStall.Tag = "belum"
    End If
End Sub

Private Sub chPDS_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub chUse_GotFocus()
    SSTab1.Tab = 3
End Sub

Private Sub chUse_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdAddJob_Click()
Dim tanda As String
Dim boleh As Boolean
Dim rsSearch As ADODB.Recordset
Dim xDate As Date
xDate = isnow
    boleh = False
    If Trim(text(64)) <> "" Then
        If DateDiff("d", Format(text(65), "DD/MM/YYYY"), Format(xDate, "DD/MM/YYYY")) = 0 Then
            boleh = True
        Else
            If DateDiff("d", Format(text(65), "DD/MM/YYYY"), Format(xDate, "DD/MM/YYYY")) > 0 Then
                tanda = "kemarin. "
                Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(text(64)) & "|%'")
                If Not rsSearch.EOF Then
                    If DateDiff("d", Format(rsSearch("TANGGAL"), "DD/MM/YYYY"), Format(xDate, "DD/MM/YYYY")) = 0 Then
                        boleh = True
                    Else
                        boleh = False
                    End If
                End If
                rsSearch.Close
            ElseIf DateDiff("d", Format(text(65), "DD/MM/YYYY"), Format(xDate, "DD/MM/YYYY")) < 0 Then
                tanda = "besok. "
                boleh = False
            End If
        End If
        
        If text(44).BackColor <> &HC0E0FF Then
            boleh = False
            tanda = "yang belum dikerjakan. " & vbCr & "Tambahkan di kolom Order biasanya. "
        End If
        
        If boleh Then
            With frmOAddJob
                .text(0) = text(15)
                .text(1) = text(10)
                .text(2) = text(11)
                .text(3) = text(12)
                .text(4) = text(3)
                .text(5) = text(4)
                .text(6) = Trim(text(64))
                .text(28) = text(62)
                '.text(29) = text(63)
                .rateLama = ntRate
                .text(30) = text(63)
                .text(8) = text(14)
                .text(7) = text(78)
                .text(12) = text(50)
                .text(31) = Trim(text(106))
                .text(30).Visible = True
                .text(32) = lastService
                .text(65) = text(65)
                .isNewID = objO.FetchData_Double("select top 1 nID from JobPart WITH (NOLOCK) where No_Wo = '" & Trim(text(64)) & "' Order By nID DESC")
                'TOChild frmOrderDMSNonDSS.hWnd, frmOAddJob.hWnd
                .WindowState = vbMaximized
                .Show 'vbModal
            End With
        Else
            MsgBox "Tidak bisa menambah Job Order untuk unit " & tanda
            Exit Sub
        End If
    Else
        MsgBox "Tidak boleh mengisi di Additional Job. "
        Exit Sub
    End If
End Sub

Private Sub cmdBatal_Click()
    frmeGantiNoPol.Visible = False
End Sub

Private Sub cmdBtlPass_Click()
    opByr(1).value = True
End Sub

Private Sub cmdCekHO_Click()
    If cmdPilih.Tag = "No. Rangka" Then
        If Trim(text(0)) <> "" Then
                Set rsHlp = objO.FetchDataHO("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
                If Not rsHlp.EOF Then
                    custKetemu = True
                    Call panggilCustHO(text(0), rsHlp)
                    Call simpanCustomer(text(0), rsHlp)
                    'Call simpanHistoryHO(text(0))
                End If
            
        Else
            MsgBox "Isi No Rangka"
            text(0).SelLength = Len(text(0))
            text(0).SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Fungsi ini hanya bisa untuk No. Rangka"
        Exit Sub
    End If
End Sub

Private Sub simpanCustomer(nRgk As String, rs As ADODB.Recordset)
Dim rsSearch As ADODB.Recordset
    
    If rs.State <> adStateClosed Then
        Set rsSearch = objO.FetchData("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(nRgk) & "'")
        If rsSearch.EOF Then
            Call objO.InsertData("INSERT INTO CUSTDAT VALUES (  '" & rs("CUSTOMER") & "'                    , '" & rs("CUST_ADRES") & "', '" & rs("KOTA") & "'   , '" & rs("KODEPOS") & "', '" & rs("TELPON") & "'                     , '" & rs("NO_POLISI") & "' , '" & rs("NO_ENGINE") & "'  , '" & rs("WMI") & "'        , '" & rs("NO_CHASSIS") & "'   , '" & rs("MDL") & "'      , '" & rs("MODEL") & "'                        , '" & rs("FMDL") & "'     , '" & rs("COLOR") & "'  , " _
                         & "'" & rs("FS_KD_DEALER") & "', '" & rs("FS_NM_DEALER") & "'                , '" & rs("DDATE") & "'     , '" & rs("REMARK") & "' , '" & rs("CODE") & "'   , '" & rs("NPWP") & "'                       , '" & rs("PPN") & "'       , '" & rs("DISC") & "'       , '" & rs("KEYNO") & "'      , '" & rs("USER") & "'         , '" & rs("SPK") & "'      , '" & Format(rs("TGL_REP"), "MM/DD/YYYY") & "' , '" & rs("JOB") & "'      , '" & rs("JOBCD") & "'  , '" & rs("FS_KM_AKHIR") & "', " _
                         & "'" & rs("USEJOBDSC") & "'   , '" & Format(rs("RCDATE"), "MM/DD/YYYY") & "', '" & rs("KD_K") & "'      , '" & rs("CR") & "'     , '" & rs("CDK") & "'    , '" & Format(rs("TGL_D"), "MM/DD/YYYY") & "', '" & rs("CDD") & "'       , '" & rs("CRDTCD") & "'     , '" & rs("FS_TH_MOBIL") & "', '" & rs("USEFS_NM_HOBI") & "', '" & rs("FS_KD_PLG") & "', '" & rs("FS_KD_WIL") & "'                    , '" & rs("DOTAM") & "'    , '" & rs("USENAME") & "', " _
                         & "'" & rs("USEADDR") & "'     , '" & rs("USECITY") & "'                     , '" & rs("USEPOS") & "'    , '" & rs("USEPSTN") & "', '" & rs("USEGSM") & "' , '" & rs("USECDMA") & "'                    , '" & rs("USET4LAHIR") & "', '" & rs("USETGLLAHIR") & "', '" & rs("ASURANSI") & "'   , '" & rs("ASR_ADDR") & "'     , '" & rs("ASR_CITY") & "' , '" & Format(rs("ASR_TGLM"), "MM/DD/YYYY") & "', '" & Format(rs("ASR_TGLS"), "MM/DD/YYYY") & "', " _
                         & "'" & rs("ASR_LIMIT") & "'   , '" & rs("ASR_NOTE") & "'                    , '" & rs("NOTE") & "'      , '" & rs("GSM") & "'    , '" & rs("CDMA") & "'   , '" & rs("USEAGAMA") & "'                   , '" & rs("USERADD") & "'   , '" & rs("FLAGKM") & "'     , '" & rs("HEAVYUSE") & "'   , '" & rs("PKP") & "'          , '" & rs("KDPLG_ASR") & "')")
        End If
    End If
End Sub

Private Sub cmdCetak_Click()
    'Call CetakForm
    If Not sgnPreDiag.Visible Then
        Call CetakForm1
    Else
        cetakFormPrediag
    End If
End Sub

Private Sub cmdFU_Click()
    Set rsHlp = objO.FetchData("SELECT * FROM JOBDTL WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "' AND no_wo = '" & Trim(text(64)) & "'")
    If Not rsHlp.EOF Then
        If Trim(rsHlp("JAMFU")) <> "" Then
        Call objO.InsertData("INSERT INTO tempFU SELECT NO_WO, NO_CHASSIS, TGLFU,JAMFU,TELPFU,KETFU FROM JOBDTL WITH (NOLOCK) " _
                     & "WHERE NOT EXISTS (SELECT * FROM tempFU WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "' AND no_wo = '" & Trim(text(64)) & "') AND " _
                     & "no_chassis = '" & text(15) & "' AND no_wo = '" & Trim(text(64)) & "' ORDER BY TGLFU")
        End If
    End If
    
    With frmO3Follow
        .text(0) = text(15)
        .text(1) = text(10)
        .text(2) = text(13)
        .text(3) = text(12)
        .text(4) = text(3)
        .text(5) = text(4)
        .text(6) = IIf(sgnPreDiag.Visible, text(22), Trim(text(64)))
        .text(7) = TglDDMMYYYY(Format(DateAdd("d", 3, TglDDMMYYYY(text(88))), "dd/mm/yyyy"))
        
        Set rsHlp = objO.FetchData("SELECT * FROM TEMPFU WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_CHASSIS = '" & text(15) & "'")
        If Not rsHlp.EOF Then
            .text(8) = rsHlp("JAMFU")
            .text(9) = rsHlp("TELPFU")
            .text(10) = rsHlp("KETFU")
        End If
        rsHlp.Close
        
        .Show vbModal
    End With
End Sub

Private Sub cmdGanti_Click()
    Dim rsSearch As ADODB.Recordset
    Dim noPolBaru As String
    erTranAll = False
    Call objHT.OpenTransactionBengkel
    
    noPolBaru = Trim(text(111)) & IIf(Trim(text(111)) <> "", "-", "") & Trim(text(112)) & IIf(Trim(text(112)) <> "", "-", "") & Trim(text(113))
    Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    If Not rsSearch.EOF Then
        Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET NO_POLISI = '" & Trim(noPolBaru) & "' WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    End If
    
    Set rsSearch = objO.FetchData("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(15)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    If Not rsSearch.EOF Then
        Call objHT.UpdateData_BengkelOnly("UPDATE CUSTDAT SET NO_POLISI = '" & Trim(noPolBaru) & "' WHERE NO_CHASSIS = '" & Trim(text(15)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    End If
    
    Set rsSearch = objO.FetchData("SELECT * FROM DTLFU WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    If Not rsSearch.EOF Then
        Call objHT.UpdateData_BengkelOnly("UPDATE DTLFU SET NO_POLISI = '" & Trim(noPolBaru) & "' WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    End If
    
    Set rsSearch = objO.FetchData("SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    If Not rsSearch.EOF Then
        Call objHT.UpdateData_BengkelOnly("UPDATE JOBSUBLET SET NO_POLISI = '" & Trim(noPolBaru) & "' WHERE NO_WO = '" & Trim(text(64)) & "' AND NO_POLISI = '" & Trim(text(110)) & "'")
    End If
    
    Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '" & Trim(text(110)) & " |" & Trim(text(64)) & "|%' ")
    If Not rsSearch.EOF Then
        Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET KETERANGAN = '" & noPolBaru & "'+SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN), CHARINDEX('| ',KETERANGAN)+2-CHARINDEX(' |',KETERANGAN))+ '" & Trim(text(3)) & "' WHERE KETERANGAN LIKE '" & Trim(text(110)) & " |" & Trim(text(64)) & "|%'")
        Call objHT.UpdateData_BengkelOnly("UPDATE TEMPSTALL SET KETERANGAN = '" & noPolBaru & "'+SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN), CHARINDEX('| ',KETERANGAN)+2-CHARINDEX(' |',KETERANGAN))+ '" & Trim(text(3)) & "' WHERE KETERANGAN LIKE '" & Trim(text(110)) & " |" & Trim(text(64)) & "|%'")
    End If
    
    If erTranAll = False Then
        text(14) = Trim(noPolBaru)
        Call objHT.CloseTransactionBengkel
    Else
        Call objHT.CloseTransactionBengkel_err
    End If
    frmeGantiNoPol.Visible = False
End Sub

Private Sub cmdGantiNoPol_Click()
    text(110) = Trim(text(14))
    frmeGantiNoPol.Top = 2520
    frmeGantiNoPol.Left = 4200
    frmeGantiNoPol.Visible = True
    text(111).SetFocus
End Sub

Private Sub cmdGateS_Click()
    Dim Report As New RptGatePassSementara
    Dim rsRepo As ADODB.Recordset
    
    Set rsRepo = objO.FetchData("SELECT NO_WO, NOINV, NO_POLISI, NO_CHASSIS, MDL, FMDL, CUSTOMER, CUST_ADRES, ALASAN_GATE FROM JOBDTA WITH (NOLOCK) WHERE NOINV = '" & text(22) & "' AND NO_WO = '" & text(0) & "'")
    If Not rsRepo.EOF Then
        Report.Database.SetDataSource rsRepo
        Report.txtPT.SetText Trim(gPerush)
        Report.txtAlamat.SetText Trim(gAlamat) & "  " & Trim(gKotaCbg)
    
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Zoom 100
            .Visible = True
        End With
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Data belum disimpan"
    End If
    'rsRepo.Close
End Sub

Private Sub cmdHistory_Click()
    If Replace(Replace(text(15), ">", ""), " ", "") = "" Then Exit Sub
    If Log.LogHOConn = 0 Then
        MsgBox "Gagal Koneksi ke Server HO !", vbInformation + vbOKOnly, "History On Line HO"
        Exit Sub
    End If
    gTag = "ONLINE"
    dRangka = text(15)
    nData(0) = text(15)
    nData(1) = text(3)
    nData(2) = Replace(Replace(text(14), ">", ""), " ", "")
    Dim fH As New FrmHistory_new
    fH.Show (vbModal)
End Sub
Private Sub Command1_Click()
    If Replace(Replace(text(15), ">", ""), " ", "") = "" Then Exit Sub
    gTag = "LOCAL"
    dRangka = text(15)
    nData(0) = text(15)
    nData(1) = text(3)
    nData(2) = Replace(Replace(text(14), ">", ""), " ", "")
    Dim fH As New FrmHistory_new
    fH.isOrder = True
    fH.noBaris = baris
    fH.modelKendaraan = Trim(text(10))
    fH.Show (vbModal)
End Sub


Private Sub cmdKodePlg_Click()
    With frmtCustomer
        gAmbilPlg = "order"
        If Left(text(15), 3) <> "" Then
            .text(50) = Left(text(15), 3)
        End If
        .text(1) = Mid(text(15), 4, 6)
        .text(0) = Mid(text(15), 10, IIf(Len(text(15)) > 0, Len(text(15)) - 9, 0))
        .text(2) = text(1)
        .text(3) = text(2)
        .text(4) = text(3)
        .text(5) = text(4)
        .text(6) = text(5)
        .text(8) = text(6)
        .text(9) = text(7)
        .text(10) = text(8)
        .text(20) = text(9)
        If Trim(text(20)) = "" Then text(20) = 0
        .text(19).Mask = ""
        .text(19) = IIf(Trim(text(16)) = "01-01-3000", "", IIf(Trim(text(16)) = "01-01-1900", "", Trim(text(16))))
        .text(19).Mask = "99-99-9999"
        'text(16)
        .text(21) = text(10)
        .text(22) = text(11)
        .text(23) = text(12)
        .text(24) = text(13)
        If Not SgnRekom.Visible Then
            If Trim(text(14)) <> "" Then
        Dim nPol As String
                nPol = Trim(text(14))
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""
                Call pecahNoPol(nPol, nDpn1, nTgh1, nBlk1)
                .text(25) = nDpn1
                .text(46) = nTgh1
                .text(47) = Trim(nBlk1)
            End If
        End If
        .text(26) = text(76)
        gBantuan = IIf(cmdPilih.Tag = "No. Booking", "booking", "")
        gBantuan2 = IIf(cmdPilih.Tag = "No. Booking", text(0), "")
        frmtCustomer.Show vbModal
    End With
End Sub

Private Sub cmdLengkap_Click()
    frmLengkap.Visible = True
End Sub

Public Sub cmdPilih_Click()
    cmdCekHO.Visible = False
    cmdGantiNoPol.Visible = False
    If cmdPilih.Tag = "No. Booking" Then
        cmdPilih.Tag = "No. Rangka"
        'cmdCekHO.Visible = True
        cmdPilih.Picture = ImageList1.ListImages(2).Picture
    ElseIf cmdPilih.Tag = "No. Rangka" Then
        cmdPilih.Tag = "No. WO"
        cmdPilih.Picture = ImageList1.ListImages(3).Picture
    ElseIf cmdPilih.Tag = "No. WO" Then
        cmdPilih.Tag = "No. Booking"
        cmdPilih.Picture = ImageList1.ListImages(1).Picture
    End If
    If text(3) <> "" Then Call hapusForm
End Sub

Private Sub cmdPilih_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0) = ""
        text(0).SetFocus
        cmdPilih.Tag = "No. Rangka"
        cmdPilih.Picture = ImageList1.ListImages(2).Picture
    End If
End Sub

Private Sub cmdStall_Click()
Dim pesan As String
Dim nActRate As Double
    
    boleh = True
    If baris = 1 Then
        For i = 1 To 7
            If DGrid.TextMatrix(i, 1) <> "" Then
                baris = baris + 1
            End If
        Next
    End If
    If baris = 1 Then
        pesan = pesan + "Job Order. "
        boleh = False
    End If
    If text(14) = "" Then
        pesan = pesan + "No Polisi. "
        boleh = False
    End If
    If text(15) = "" Then
        pesan = pesan + "No Rangka. "
        boleh = False
    End If
    If text(3) = "" Then
        pesan = pesan + "Kode Pelanggan. "
        boleh = False
    End If
    If woOnProgres Then
        pesan = pesan + "Unit Sudah dikerjakan, Tidak bisa diganti. "
        boleh = False
    End If
        
    If boleh Then
        If Not opTpSvc(2).value Then
            ' ISI TEMPSTALL
            Call objO.DeleteData("DELETE TEMPSTALL WHERE tanggal = '" & Format(Trim(text(65)), "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(text(14), " ", ""), ">", "") & " |%'")
            Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                         & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(65)), "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(text(14), " ", ""), ">", "") & " |%') AND " _
                         & "tanggal = '" & Format(Trim(text(65)), "mm/dd/yyyy") & "' AND keterangan like '%" & Replace(Replace(text(14), " ", ""), ">", "") & " |" & Trim(text(64)) & "| " & "%' ORDER BY noStall, Tanggal, JamMulai")
        Else
            Call objO.DeleteData("DELETE TEMPSTALL WHERE tanggal = '" & Format(Trim(text(65)), "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(text(14), " ", ""), ">", "") & " |%'")
            Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                         & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(65)), "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(text(14), " ", ""), ">", "") & " |%') AND " _
                         & "tanggal = '" & Format(Trim(text(65)), "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(text(14), " ", ""), ">", "") & " |" & Trim(text(64)) & "| %' AND [GROUP] = 'W' ORDER BY noStall, Tanggal, JamMulai")
        End If
        
        nActRate = 0
        For i = 1 To 8
            If DGrid.TextMatrix(i, 1) <> "" Then
                Set rsTemp = objO.FetchData("SELECT * FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(text(10)) & "' AND FS_LOG_DEL = ''")
                If Not rsTemp.EOF Then
                    nActRate = nActRate + CDbl(rsTemp("RATE"))
                End If
            End If
        Next
        
        If chEM = Checked Then
            nActRate = 0.5
        End If
        If chSBNP = Checked Then
            nActRate = 1
        End If
        
        If nActRate > 0 Then
        gStall = "order"
        gTanggalStall = Format(Trim(text(65)), "DD-MM-YYYY")
        
        If Log.isEmStall = 1 Then
            With frmStall
                .text(7) = text(65)
                .text(4).Mask = ""
                .text(4) = Replace(text(14), ">", " ")
                .text(5) = text(1)
                .text(6) = text(3)
                .text(8) = text(50)
                .text(9) = nActRate
                .text(12) = IIf(sgnPreDiag.Visible, text(22), Trim(text(64)))
                .frmStal.Visible = False
                .chkThs.value = chTHS.value
                .chWar.value = IIf(opTpSvc(2).value, Checked, Unchecked)
                .Show
                .WindowState = vbMaximized
            End With
        Else
            With frmStallNonEM
                .text(7) = text(65)
                .text(4).Mask = ""
                .text(4) = Replace(text(14), ">", " ")
                .text(5) = text(1)
                .text(6) = text(3)
                .text(8) = text(50)
                .text(9) = nActRate
                .text(12) = IIf(sgnPreDiag.Visible, text(22), Trim(text(64)))
                .frmStal.Visible = False
                .chkThs.value = chTHS.value
                .chWar.value = IIf(opTpSvc(2).value, Checked, Unchecked)
                .Show
                .WindowState = vbMaximized
            End With
        End If
        
        'lakukan proses perhitungan estimasi
        newWo = False
        If Trim(text(64)) = "" Then newWo = True
        
            If newho Then
                nWKT = hitungWaktu(nActRate, text(66), text(65))
            Else
                Dim gTglTampung As Date
                gTglTampung = gTglServer
                nWKT = hitungWaktu(nActRate, Format(gTglTampung, "HH:NN"), TglDDMMYYYY(gTglTampung))
            End If
            text(63) = Left(nWKT, InStr(nWKT, "|") - 1)
            text(62) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
            Call hitungEstimasiPenyerahan
            
        Else
            MsgBox "Belum mengisi Actual Rate untuk job order tersebut atau model tersebut !!!", vbCritical, "ACTUAL RATE"
            Exit Sub
        End If
    Else
        MsgBox "Lengkapi data Order !! " & vbCr & pesan, vbCritical, "Isi Field Kosong"
        Exit Sub
    End If
End Sub

Private Sub cmdSublet_Click()
Dim subletValid As Boolean
Dim rsSublet As ADODB.Recordset
    
    subletValid = False
    Set rsSublet = objO.FetchData("SELECT DISTINCT VALID FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(text(64)) & "' AND no_chassis = '" & Trim(text(15)) & "' ORDER BY VALID DESC")
    If Not rsSublet.EOF Then
        If Trim(rsSublet("VALID")) = "1" Then
            subletValid = True
        End If
    End If
    
    If Not subletValid Then
        If Trim(text(15)) <> "" Then
            Call objO.UpdateData("UPDATE TEMPJOBSUBLET SET NO_WO = '" & Trim(text(64)) & "', NO_CHASSIS = '" & Trim(text(15)) & "' WHERE NO_POLISI = '" & Trim(text(14)) & "' AND NOFAK = '" & Trim(text(78)) & "' AND NO_WO = '' AND NO_CHASSIS = ''")
            If Trim(text(64)) <> "" Then
                Dim rsSupply As ADODB.Recordset
                Set rsSupply = objO.FetchData("SELECT * FROM jobsublet WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & Trim(text(64)) & "' AND no_chassis = '" & text(15) & "') and no_wo = '" & Trim(text(64)) & "' AND no_chassis = '" & Trim(text(15)) & "' ORDER BY nonota, kodeitem")
                If Not rsSupply.EOF Then
                    Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE no_wo = '" & Trim(text(64)) & "' AND no_chassis = '" & Trim(text(15)) & "'")
                    Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI,PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
                                 & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & Trim(text(64)) & "' AND no_chassis = '" & text(15) & "') and no_wo = '" & Trim(text(64)) & "' AND no_chassis = '" & Trim(text(15)) & "' " _
                                 & "ORDER BY nonota, kodeitem")
                End If
            End If
            With frmOSublet
                gTag = "order"
                gPO = objO.FetchData_String("SELECT NOPO FROM JOBSUBLET WITH (NOLOCK) WHERE no_wo = '" & Trim(text(64)) & "' AND no_chassis = '" & text(15) & "'")
                .text(0) = text(15)
                .text(1) = text(10)
                .text(2) = text(11)
                .text(3) = text(12)
                .text(4) = text(3)
                .text(5) = text(4)
                .text(6) = Trim(text(64))
                .text(7) = gPO
                .text(14) = text(14)
                .text(15) = text(78)
                .cmdKeluar.Caption = "Kembali ke &Order"
                .chAdd = Unchecked
                .text(13).Visible = False
                .Label11.Visible = False
                .Show vbModal
            End With
            
            Call hitungEstimasiPenyerahan
        Else
            MsgBox "Data Customer, No. Rangka Belum ada !!!"
            cmdKodePlg.SetFocus
        End If
    Else
        MsgBox "Sublet sudah divalidasi, Lakukan retur sublet !!"
    End If
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub cmdWaktu_Click()
    If frmeWaktu.Visible Then
        frmeWaktu.Visible = False
    Else
        If text(65) <> "" And text(65) <> ">>->>->>>>" Then
            frmeWaktu.Visible = True
            frmeWaktu.Top = cmdWaktu.Top + cmdWaktu.Height + 100
            frmeWaktu.Left = cmdWaktu.Left - (frmeWaktu.Width - cmdWaktu.Width)
            Call TampilWaktu(Trim(text(65)))
        End If
    End If
End Sub

Private Sub cmdWarranty_Click()
    gTag = "ORDER"
    opTpSvc(2).value = True
    With frmWarranty
        .text(0) = text(15)
        .text(1) = text(3)
        .text(2) = Replace(Replace(text(14), ">", ""), " ", "")
        .Brs.Caption = baris
        .BrsPart.Caption = barispart
        .text(3) = text(65)
        .text(4) = text(66)
        .Show (vbModal)
    End With
End Sub

Private Sub DGrid_DblClick()
    If baris > 1 And DGrid.TextMatrix(DGrid.row, 1) <> "" Then

        If Trim(text(64)) <> "" Or Trim(text(78)) <> "" Then
            Dim nCexPP As String
            Dim rsPP As ADODB.Recordset
            'Set rsPP = objO.FetchData("Select Item_Part from HO_PartPack Where MDL = '" & Text(10) & "' and FS_Item = '" & dGrid.TextMatrix(dGrid.row, 1) & "' and JOB = '" & dGrid.TextMatrix(dGrid.row, 4) & "'")
            Set rsPP = objO.FetchData("Select Nama_Part from HO_PartPack WITH (NOLOCK) Where MDL = '" & text(10) & "' and FS_Item = '" & DGrid.TextMatrix(DGrid.row, 1) & "' and JOB = '" & DGrid.TextMatrix(DGrid.row, 4) & "'")
            If rsPP.RecordCount > 0 Then
                nCexPP = "Select NOSO FROM SOORD WITH (NOLOCK) Where (NoSO = '" & Trim(text(64)) & "' OR NOSPK = '" & Trim(text(78)) & "') and ("
                rsPP.MoveFirst
                Do Until rsPP.EOF
                    'nCexPP = nCexPP & " KodeItem like '%" & rsPP.Fields(0).value & "' OR "
                    nCexPP = nCexPP & " NamaBYSA = '" & rsPP.Fields(0).value & "' OR "
                    rsPP.MoveNext
                Loop
                nCexPP = Left(nCexPP, Len(nCexPP) - 3) & ")"
                nCexPP = nCexPP & " AND ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0  Or IsNull(QtyReserve,0) <> 0) AND ([TO] <> 0 Or BO not in ('','CANCEL') )"
            
                Dim rsPSO As ADODB.Recordset
                Set rsPSO = objO.FetchDataACC(nCexPP)
                If rsPSO.RecordCount > 0 Then
                    MsgBox "Paket Part sudah di proses oleh partman !," & Chr(13) & "Job Order tidak dapat di hapus" & Chr(13) & Chr(13) & "Lakukan pembatalan proses part oleh Partman, untuk menghapus Job Order"
                    Exit Sub
                End If
            End If
        End If
        
        a = DGrid.row
        'AMBIL
        With DGrid
            text(44) = .TextMatrix(a, 4)
            text(46) = .TextMatrix(a, 1)
            Jrate = .TextMatrix(a, 2)
            If .TextMatrix(a, 3) <> "" Then
                If .TextMatrix(a, 3) Then
                    fOther = True
                    text(47).Visible = True
                    text(47).Mask = ""
                    text(47) = Jrate
                Else
                    fOther = False
                    text(47).Visible = False
                End If
            Else
                fOther = False
                text(47).Visible = False
            End If
        End With
        
        ' HITUNG KURANG JAM KERJA MEKANIK
        'dGrid.TextMatrix(a, 4) = objO.FetchData_String("SELECT JOB FROM jobOrder WHERE fs_item = '" & Text(46) & "'")
        If rsGridPart.State <> adStateClosed Then
            Call cekPartPack(DGrid.TextMatrix(a, 4), DGrid.TextMatrix(a, 1), text(10), 2, 2, IIf(rsGridPart.State = adStateClosed, 0, rsGridPart.RecordCount), frmOrderDMSNonDSS, rsGridPart)
        End If
        
        'hrgLabour = hitungLabour(DGrid.TextMatrix(baris - 1, 2), Text(11))
        myChassis = Trim(text(15))
        'text(56) = Format(CDbl(text(56) - CDbl(hitungLabour(DGrid.TextMatrix(baris - 1, 2), text(11), IIf(sgnPreDiag.Visible, True, False)))), "0#,#") ' * CDbl(Jrate)
        text(56) = Format(CDbl(text(56) - CDbl(hitungLabour(DGrid.TextMatrix(a, 2), text(11), IIf(sgnPreDiag.Visible, True, False)))), "0#,#")
        text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
        If isJualNRM = True Then
            Me.JualNRM.Caption = nJualNRM
        Else
            Me.JualNRM.Caption = ""
        End If
        'HAPUS
        For i = a To baris
            With DGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
                .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
            End With
        Next
        
        baris = baris - 1
        If baris = 1 Then
            text(86) = text(66)
            text(62) = text(65)
            Call hitungEstimasiPenyerahan
        End If
                
        text(44).SelLength = Len(text(44))
        text(44).SetFocus
        If (text(44)) = "SBI" Then isSBI = False
    End If
End Sub

Private Sub dGrid_GotFocus()
Dim bolehTambah As Boolean
'Dim hrgLabour
'Dim nA As Double
    If chPDS.value = 1 Then
        If text(44).text = "SBE" Or text(44).text = "SBI" Or text(44).text = "GRP" Then
            MsgBox "Unit PDS, tidak boleh dengan Job Order SBI/SBE/GRP !", vbInformation + vbOKOnly, "Validasi PDS"
            text(44).text = ""
            Exit Sub
        End If
    End If
    If Trim(text(46)) <> "" And Trim(text(50)) <> "" And Nul(Jrate) <> "" Then  'And CDbl(Jrate) > 0
        For i = 0 To baris
            If DGrid.TextMatrix(i, 1) = Trim(text(46)) Then
                bolehTambah = False
                text(46) = ""
                Exit Sub
            Else
                bolehTambah = True
            End If
        Next
        If bolehTambah Then
            If Trim(text(44)) = "SBE" Or Trim(text(44)) = "SBI" Then
                '' banding last sbe service dengan yg baru masuk
                Dim rsSearch As ADODB.Recordset
                Set rsSearch = objO.FetchData("SELECT FS_NILAI FROM T_REGULAR WITH (NOLOCK) WHERE FS_KETERANGAN = '" & Mid(text(46), InStr(text(46), "KALA ") + 5, InStr(text(46), "KM") - (InStr(text(46), "KALA ") + 5)) & "'")
                If Not rsSearch.EOF Then
                    If CDbl(lastService) >= CDbl(rsSearch("FS_NILAI")) Then
                        MsgBox "Service berkala yang dipilih tidak boleh lebih kecil atau sama dengan History sebelumnya !!", vbCritical, "Salah SBE"
                        bolehTambah = False
                    Else
                        bolehTambah = True
                    End If
                End If
            End If
        End If
        If Trim(text(44)) = "SBI" Then isSBI = True
        If bolehTambah Then
            ' JIKA PEKERJAAN INPUT SENDIRI
            If fOther Then
                Jrate = CDbl(Replace(Replace(text(47), " ", ""), ">", ""))
            End If
            
            myChassis = Trim(text(15))
            If MaxBrs >= baris Then
                text(56) = Format(CDbl(text(56)) + Val(hitungLabour(CDbl(Jrate), Trim(text(11)), IIf(sgnPreDiag.Visible, True, False))), "#0,0")
                With DGrid
                    .TextMatrix(baris, 1) = Trim(text(46))
                    .TextMatrix(baris, 2) = Jrate
                    .TextMatrix(baris, 3) = fOther
                    .TextMatrix(baris, 4) = Trim(text(44))
                    .TextMatrix(baris, 5) = Val(hitungLabour(CDbl(Jrate), Trim(text(11)), IIf(sgnPreDiag.Visible, True, False)))
                    .TextMatrix(baris, 6) = Trim(text(50))
                End With
                If isJualNRM = True Then
                    Me.JualNRM.Caption = nJualNRM
                Else
                    Me.JualNRM.Caption = ""
                End If
                baris = baris + 1
                
                ' ISI PART PACK
                If rsGridPart.State = adStateClosed Then
                    Call createTblPart
                End If
                Call cekPartPack(text(44), text(46), text(10), 1, 2, , frmOrderDMSNonDSS, rsGridPart)
                Set partGrid.DataSource = rsGridPart
                Set partGridFill.DataSource = rsGridPart
                Call HeadColumn
                
                ntRate = 0
                text(114) = ntRate
                For i = 1 To 7
                    If Trim(DGrid.TextMatrix(i, 2)) <> "" Then
                        ntRate = ntRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(text(46)) & "' AND MDL = '" & Trim(text(10)) & "'")
                    End If
                Next
                text(114) = ntRate
                
                If text(66) <> ">>:>>" And text(65) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(ntRate), text(66), text(65))
                    text(86) = text(66)
                End If
                text(46) = ""
                text(44) = ""
                text(45) = ""
                text(47) = ">>,>>"
                text(47).Visible = False
                text(44).SetFocus
                text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
            
            Else
                MsgBox "Maksimal " & str(MaxBrs) & " !"
                text(46).SetFocus
            End If
        End If
    End If
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            Call DGrid_DblClick
    End Select
End Sub

'Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        If X > (ImgCrop.Left - 600) And X < ((ImgCrop.Left - 600) + ImgCrop.Width) Then
'            If Y > (ImgCrop.Top - 600) And Y < ((ImgCrop.Top - 600) + ImgCrop.Height) Then
'                imgTanda(nk).Top = 600 + Y - (imgTanda(nk).Height / 2)
'                imgTanda(nk).Left = 600 + X - (imgTanda(nk).Width / 2)
'                imgTanda(nk).Visible = True
'                KorX(nk) = X
'                KorY(nk) = Y
'                nk = nk + 1
'            End If
'        End If
'    End If
'End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If X > (ImgCrop.Left - 600) And X < ((ImgCrop.Left - 600) + ImgCrop.Width) Then
            If Y > (ImgCrop.Top - 600) And Y < ((ImgCrop.Top - 600) + ImgCrop.Height) Then
                imgTanda(nk).Top = 600 + Y - (imgTanda(nk).Height / 2)
                imgTanda(nk).Left = 600 + X - (imgTanda(nk).Width / 2)
                imgTanda(nk).Visible = True
                KorX(nk) = X
                KorY(nk) = Y
                nk = nk + 1
            End If
        End If
    End If
End Sub


Private Sub imgTanda_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then
        For i = Index To nk - 1
            KorX(i) = KorX(i + 1)
            KorY(i) = KorY(i + 1)

            imgTanda(i).Top = imgTanda(i + 1).Top
            imgTanda(i).Left = imgTanda(i + 1).Left

            If i = nk - 1 Then
                imgTanda(i).Visible = False
            End If
        Next
        nk = nk - 1
    'End If
End Sub

Private Sub opBBM_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub opByr_Click(Index As Integer)
    Select Case Index
        Case 0
            Updated = True
        Case 1
            Updated = True
        Case 2
            If Not opTpSvc(2).value And chPDS.value <> 1 Then
                If Updated Then
                    If Trim(text(90)) = "" Then
                        MsgBox "Tidak bisa melakukan proses tagihan. " & vbCr & "Harus untuk unit yang terdaftar MOU."
                        opByr(0).value = True
                    Else
                        If Log.LogHOConn = 1 Then
                            Set rsHlp = objO.FetchDataHO("SELECT LIMITAR, HUTANG - BAYAR AS SISA, [TOP] FROM MOUHDR WITH (NOLOCK) WHERE NO_MOU = '" & Trim(text(90)) & "'")
                            If Not rsHlp.EOF Then
                                If CDbl(rsHlp("LIMITAR")) <> 0 Then
                                    If CDbl(rsHlp("LIMITAR")) <= CDbl(rsHlp("SISA")) + CDbl(text(61)) Then
                                        MsgBox "Tagihan melebihi batas perjanjian. Tidak bisa dengan Tagihan. " & vbCr & "Harap melakukan pelunasan terlebih dahulu. "
                                        opByr(0).value = True
                                    End If
                                Else
                                    MsgBox "Tidak bisa melakukan proses tagihan. " & vbCr & "Perjanjian pembayaran " & rsHlp("TOP")
                                    opByr(0).value = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Private Sub opByr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub opGanti_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub opLengkap_Click(Index As Integer)

    text(115 + Index).Visible = False
    If opLengkap(Index) = Checked Then
        text(115 + Index).Visible = True
    Else
        text(115 + Index) = ""
    End If
End Sub

Private Sub opSvc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub opTpSvc_Click(Index As Integer)
Dim getPrediag As Boolean

    opByr(0).value = True
    opByr(0).Enabled = True
    opByr(1).Enabled = True
    sgnPreDiag.Visible = False
    Select Case Index
        Case 0, 2
            ' SERVICE BIASA
            text(21).Visible = False
            text(22).Visible = False
            MaskEdBox12.Visible = False
            cmdWarranty.Enabled = False
            text(44) = ""
            text(44).BackColor = &H80000005
            chPDS.Enabled = True
            If Index = 2 Then
                If cmdPilih.Tag = "No. Rangka" Then
                    '' Warranty
                    opByr(2).value = True
                    opByr(0).Enabled = False
                    opByr(1).Enabled = False
                    chPDS.Enabled = False
                    
                    getPrediag = False
                    text(22).Visible = True
                    If lastPrediag(Trim(text(15))) Then
                        If MsgBox("Edit PreDiagnosis lama ?", vbYesNo + vbQuestion, "PreDiagnosis") = vbYes Then
                            FrDiag.Left = 0
                            Set rsHlp = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(15)) & "' AND DELIV <> 'BTL' AND VALIDTWC = ''")
                            If Not rsHlp.EOF Then
                                Do While Not rsHlp.EOF
                                    If MsgBox("Edit No Prediagnosis : " & Trim(rsHlp("NO_DIAG")) & vbCr & "Permintaan : " & Trim(rsHlp("REQ1")), vbYesNo + vbDefaultButton1 + vbQuestion, "PreDiagnosis") = vbYes Then
                                        sgnPreDiag.Visible = True
                                        Call panggilPreDiagnosis(Trim(rsHlp("NO_DIAG")), Trim(text(15)), "Edit")
                                        getPrediag = True
                                        GoTo endFunction
                                    End If
                                    rsHlp.MoveNext
                                Loop
                                If Not getPrediag Then
                                    MsgBox "Tidak ada preDiagnosis lain", vbExclamation
                                End If
                            End If
                            Call hapusForm
                        Else
                            If allowPrediagToWO(Trim(text(15))) Then
                                If MsgBox("Buat PreDiagnosis Jadi WO ?", vbYesNo + vbDefaultButton2 + vbQuestion, "PreDiagnosis to WO Warranty") = vbYes Then
                                    FrDiag.Left = -9000
                                    Set rsHlp = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(text(15)) & "' AND DELIV <> 'BTL' AND VALIDTWC = 'OKE' AND NO_WO = ''")
                                    If Not rsHlp.EOF Then
                                        Do While Not rsHlp.EOF
                                            If MsgBox("Proses No Prediagnosis : " & Trim(rsHlp("NO_DIAG")) & vbCr & "Permintaan : " & Trim(rsHlp("REQ1")), vbYesNo + vbDefaultButton2 + vbQuestion, "PreDiagnosis") = vbYes Then
                                                Call panggilPreDiagnosis(Trim(rsHlp("NO_DIAG")), Trim(text(15)), "WO")
                                                getPrediag = True
                                                GoTo endFunction
                                            End If
                                            rsHlp.MoveNext
                                        Loop
                                        If Not getPrediag Then
                                            MsgBox "Tidak ada preDiagnosis lain", vbExclamation
                                            Call hapusForm
                                        End If
                                    End If
                                Else
                                    If allowNewPrediag(Trim(text(15))) Then
                                        If MsgBox("Buat PreDiagnosis baru ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Warranty") = vbYes Then
                                            sgnPreDiag.Visible = True
                                        Else
                                            Call hapusForm
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox "Tidak ada preDiagnosis yang sudah disetujui", vbInformation
                                If allowNewPrediag(Trim(text(15))) Then
                                    If MsgBox("Buat PreDiagnosis baru ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Warranty") = vbYes Then
                                        FrDiag.Left = 0
                                        sgnPreDiag.Visible = True
                                    Else
                                        Call hapusForm
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If allowNewPrediag(Trim(text(15))) Then
                            If MsgBox("Buat PreDiagnosis baru ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Warranty") = vbYes Then
                                FrDiag.Left = 0
                                sgnPreDiag.Visible = True
                            Else
                                Call hapusForm
                            End If
                        End If
                    End If
                ElseIf cmdPilih.Tag = "No. Booking" Then
                    'lakukan pengecekan apakah dari prediagnose
                    Dim diagOrderDiagnose As New ADODB.Recordset
                    Set diagOrderDiagnose = objO.FetchData("select da.*,d.no_polisi,d.* from diagdta d WITH (NOLOCK) inner join diagapproval da WITH (NOLOCK) on da.no_diag = d.no_diag where da.no_wo = '' and isnull(da.alasan_tdksetuju,'') = '' and  da.nofak = '" & Trim(text(0)) & "'")
                    If Not diagOrderDiagnose.EOF Then
                        'jika ada data dari diagnose
                        opTpSvc(2).value = True
                        opTpSvc(0).Enabled = False
                        opTpSvc(1).Enabled = False
                        chPDS.Enabled = False
                        
                        getPrediag = False
                        text(22).Visible = True
                    Else
                        opTpSvc(0).value = True
                        opTpSvc(0).Enabled = True
                        opTpSvc(1).Enabled = True
                        opTpSvc(2).Enabled = False
                        MsgBox "Order Warranty tidak bisa diproses dari Booking ! (batalkan Booking untuk melanjutkan Warranty)", vbExclamation
                        Call hapusForm
                    End If
                    
                End If
endFunction:
            End If
        Case 1
            ' PROGRAM
            text(21).Visible = True
            text(22).Visible = True
            MaskEdBox12.Visible = True
            
            text(21) = ""
            text(22) = ""
            cmdWarranty.Enabled = False
            If Index = 2 Then
                opByr(2).value = True
                opByr(0).Enabled = False
                opByr(1).Enabled = False
                text(44) = "TWC"
                text(44).BackColor = &HC0E0FF   '&HC0E0FF
            Else
                text(21).SelLength = Len(text(21))
                text(44) = ""
                text(44).BackColor = &H80000005
            End If
    End Select
End Sub

Private Sub opTpSvc_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        text(23).SetFocus
    End If
End Sub

Private Sub opWash_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    
    
End Sub

Private Sub partGrid_AfterColUpdate(ByVal ColIndex As Integer)
    
    With partGrid
        If ColIndex = 5 Then
            If .Columns(4).value <> -1 Then
                .Columns(5).value = 0
            Else
                If .Columns(5).value = -1 Then
                    If Not rsGridPart("OLI") Then
                        text(57) = CDbl(text(57)) - HargaPart(rsGridPart)
                    Else
                        text(59) = CDbl(text(59)) - HargaPart(rsGridPart)
                    End If
                Else
                    If Not rsGridPart("OLI") Then
                        text(57) = CDbl(text(57)) + HargaPart(rsGridPart)
                    Else
                        text(59) = CDbl(text(59)) + HargaPart(rsGridPart)
                    End If
                End If
            End If
        ElseIf ColIndex = 4 Then
            If .Columns(4).value <> -1 Then
                If Not rsGridPart("OLI") Then
                    text(57) = CDbl(text(57)) - HargaPart(rsGridPart)
                Else
                    text(59) = CDbl(text(59)) - HargaPart(rsGridPart)
                End If
            Else
                If Not rsGridPart("OLI") Then
                    text(57) = CDbl(text(57)) + CDbl(HargaPart(rsGridPart))
                Else
                    text(59) = CDbl(text(59)) + CDbl(HargaPart(rsGridPart))
                End If
            End If
        End If
        
        text(57) = Format(text(57), "#0,#")
        text(59) = Format(text(59), "#0,#")
    End With
End Sub

Private Sub partGridFill_AfterColUpdate(ByVal ColIndex As Integer)
    
    With partGrid
        If ColIndex = 5 Then
            If .Columns(4).value <> -1 Then
                .Columns(5).value = 0
            Else
                If .Columns(5).value = -1 Then
                    If Not rsGridPart("OLI") Then
                        text(57) = CDbl(text(57)) - HargaPart(rsGridPart)
                    Else
                        text(59) = CDbl(text(59)) - HargaPart(rsGridPart)
                    End If
                Else
                    If Not rsGridPart("OLI") Then
                        text(57) = CDbl(text(57)) + HargaPart(rsGridPart)
                    Else
                        text(59) = CDbl(text(59)) + HargaPart(rsGridPart)
                    End If
                End If
            End If
        ElseIf ColIndex = 4 Then
            If .Columns(4).value <> -1 Then
                If Not rsGridPart("OLI") Then
                    text(57) = CDbl(text(57)) - HargaPart(rsGridPart)
                Else
                    text(59) = CDbl(text(59)) - HargaPart(rsGridPart)
                End If
            Else
                If Not rsGridPart("OLI") Then
                    text(57) = CDbl(text(57)) + CDbl(HargaPart(rsGridPart))
                Else
                    text(59) = CDbl(text(59)) + CDbl(HargaPart(rsGridPart))
                End If
            End If
        End If
        text(57) = Format(text(57), "#0,#")
        text(59) = Format(text(59), "#0,#")
    End With
End Sub

Private Function HargaPart(rsS As ADODB.Recordset) As Double
Dim hrg As Double
    
    With rsS
        hrg = 0
        If rsS.State <> adStateClosed Then
            If Not rsS.EOF Then
                    Set rsTemp = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM = '" & rsS("KODEITEM") & "'")
                    If Not rsTemp.EOF Then
                        hrg = CDbl(rsTemp("HARGAJUALA"))
                    End If
                HargaPart = CDbl(hrg) * CDbl(rsS("QTY"))
            End If
        End If
    End With
End Function

Private Sub partGrid_DblClick()

    If partGrid.Columns(0) <> "" Then
        text(48) = partGrid.Columns(0)
        hNamaPart = partGrid.Columns(1)
        text(49) = CDbl(partGrid.Columns(3))
            Set rsTemp = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & text(48) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
            If Not rsTemp.EOF Then
                nStok = CDbl(rsTemp("QTYSTOCK"))
            End If
        
        If Not rsGridPart.State = adStateClosed Then
            With rsGridPart
                If .RecordCount > 1 Then
                    Call hapusPart(Trim(text(48)), "", Val(rsGridPart("nID")))
                Else
                    Set rsGridPart = Nothing
                    Set partGrid.DataSource = Nothing
                End If
            End With
            partGrid.Refresh
        End If
    End If
End Sub

Private Sub partGridFill_DblClick()
    If Trim(text(64)) <> "" Then
        Dim rsSO As ADODB.Recordset
        Set rsSO = objO.FetchDataACC("Select NOSO from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(text(64)) & "' and KodeItem = '" & partGridFill.Columns(0) & "' and isNull(NamaBYSA,'') = '" & partGridFill.Columns(2) & "' and Nomor = '" & rsGridPart.Fields(8).value & "' and ADDJOB = 0  AND ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0 ) AND ([TO] <> 0 Or BO not in ('','CANCEL') )")
        If rsSO.RecordCount > 0 Then
            MsgBox "Spare Part : " & partGridFill.Columns(0) & " " & partGridFill.Columns(2) & " Sejumlah : " & partGridFill.Columns(4) & Chr(13) & "Sudah diproses !, tidak dapat dihapus/dirubah !", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    If Trim(text(78)) <> "" Then
        Set rsSO = objO.FetchDataACC("Select NOSPK from SOORD WITH (NOLOCK) Where NOSPK = '" & Trim(text(78)) & "' and KodeItem = '" & partGridFill.Columns(0) & "' and isNull(NamaBYSA,'') = '" & partGridFill.Columns(2) & "' and Nomor = '" & rsGridPart.Fields(8).value & "' and ADDJOB = 0   AND ((IsNull(QtyKirimBL,0) <> 0 OR IsNull(QtyKirimBI,0) <> 0 OR [TO] <> 0) Or BO not in ('','CANCEL') )")
        If rsSO.RecordCount > 0 Then
            MsgBox "Spare Part : " & partGridFill.Columns(0) & " " & partGridFill.Columns(2) & Chr(13) & "Sudah disupply (BOOKING)!, tidak dapat dihapus/dirubah !", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    
    If partGridFill.Columns(0) <> "" Or partGridFill.Columns(2) <> "" Then
        text(119) = IIf(partGridFill.Columns(0) <> "", partGridFill.Columns(0), partGridFill.Columns(2))
        hNamaPart = partGridFill.Columns(1)
        text(120) = CDbl(partGridFill.Columns(4))
            Set rsTemp = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & text(119) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
            If Not rsTemp.EOF Then
                nStok = CDbl(rsTemp("QTYSTOCK"))
            Else
                nStok = -1
            End If
        If Not rsGridPart.State = adStateClosed Then
            With rsGridPart
                If .RecordCount > 1 Then
                    Call hapusPart(partGridFill.Columns(0), partGridFill.Columns(2), Val(rsGridPart.Fields(8).value))
                Else
                    Set rsGridPart = Nothing
                    Set partGridFill.DataSource = Nothing
                End If
            End With
            Set partGrid.DataSource = rsGridPart
            Set partGridFill.DataSource = rsGridPart
            partGridFill.Refresh
        End If
        text(119).SetFocus
        Call Text_Validate(119, False)
    End If
End Sub

Public Sub hapusPart(hKdItem As String, hNmBySA As String, iID As Double)
    With rsGridPart
        If rsGridPart.State <> adStateClosed Then
            If rsGridPart.RecordCount > 0 Then
                If rsGridPart.EOF Then .MoveFirst
                If Not rsGridPart.BOF Then
                    .MoveFirst
                    Do While Not .EOF
                        If Trim(!KodeItem) = Trim(hKdItem) And Trim(!namabySA) = Trim(hNmBySA) And !nID = Val(iID) Then
                            If !ganti Then
                                If Not !BAWA Then
                                    If !oli Then
                                        text(59) = CDbl(text(59)) - HargaPart(rsGridPart)
                                    Else
                                        text(57) = CDbl(text(57)) - HargaPart(rsGridPart)
                                    End If
                                    text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                                End If
                            End If
                            .Delete
                        End If
                        .MoveNext
                    Loop
                End If
            End If
        End If
    End With
End Sub

Private Sub tblCetakPO_Click()
Dim Report As New RptPO
Dim rsRepo As ADODB.Recordset

    If gPO <> "" Then
        Set rsRepo = objO.FetchData("SELECT * , '" & Format(text(65), "DD/MM/YYYY") & "' AS [DATE] FROM JOBSUBLET WITH (NOLOCK) WHERE NOPO = '" & gPO & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND NO_WO = '" & Trim(text(64)) & "' ORDER BY NAMABARANG")
        Report.Database.SetDataSource rsRepo
        
        Report.txtPerusahaan.SetText gPerush
        Report.txtAlamat.SetText gAlamat
        Report.txtKota.SetText gKota
        Report.txtTelp.SetText gTelp
        Report.txtAdh.SetText gADH
        Report.txtKabeng.SetText gSecmanager
    
        With CRViewer1
            .ReportSource = Report
            .ViewReport
            .Visible = True
            
            cmdTutup.Visible = True
            cmdTutup.Left = .Left
            cmdTutup.Top = .Top + 520
            
            .Zoom 100
        End With
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub TblKeluar_Click()
    If Not simpan Then
        If MsgBox("Data Belum disimpan, akan kembali ke Menu ?", vbYesNo, "Konfirmasi") = vbNo Then
            Exit Sub
        End If
        'Call hapusForm
    End If
    If opTpSvc(2) Then
        Call cekWarranty(Trim(text(15)), Trim(text(21)))
    End If
    
    If rsGridPart.State <> adStateClosed Then
        If rsGridPart.State = adStateOpen Then rsGridPart.Close
    End If
    If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status = '" & gSA & "'")
    Unload Me
End Sub

Private Sub batalMekanik(noOrder As String)
Dim rsData As ADODB.Recordset
Dim adaSama As Boolean
Dim hlp As String
    
    Set rsData = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & noOrder & "'")
    If Not rsData.EOF Then
        For j = 1 To 7
            adaSama = False
            If Trim(DGrid.TextMatrix(j, 1)) <> "" Then
                For i = 1 To 8
                    hlp = "REP" & i
                    If Trim(rsData(hlp)) <> "" Then
                        If Trim(rsData(hlp)) = Trim(DGrid.TextMatrix(j, 1)) Then
                            adaSama = True
                            Exit For
                        End If
                    End If
                Next
                If Not adaSama Then
                    Call objO.UpdateData("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI - " & (Round(CDbl(DGrid.TextMatrix(j, 2)), 1) * 60) & " WHERE MECH = '" & Trim(DGrid.TextMatrix(j, 6)) & "' AND HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "'")
                End If
            End If
        Next
    Else
        For j = 1 To 7
            If Trim(DGrid.TextMatrix(j, 1)) <> "" Then
                Call objO.UpdateData("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI - " & (Round(CDbl(DGrid.TextMatrix(j, 2)), 1) * 60) & " WHERE MECH = '" & Trim(DGrid.TextMatrix(j, 6)) & "' AND HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "'")
            End If
        Next
    End If
    rsData.Close
End Sub

Private Sub cekWarranty(nRgk As String, kdWar As String)
    Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & nRgk & "' AND KODETIPE = '" & kdWar & "'")
    If rsHlp.EOF Then
       ' HAPUS TEMPWAR
        If onTR = True Then
            Call objHT.DeleteData_Bengkel("DELETE TEMPWAR WHERE NO_RANGKA = '" & nRgk & "' AND ID_WARRANTY = '" & kdWar & "'")
        Else
            Call objO.DeleteData("DELETE TEMPWAR WHERE NO_RANGKA = '" & nRgk & "' AND ID_WARRANTY = '" & kdWar & "'")
        End If
    End If
End Sub

Private Sub tblOkPass_Click()
Dim nPass As String
        
    nPass = Replace(text(85), "'", "")
    If (UCase(nPass) = fPassword And text(84) <> "" And (nTipe = "KACAB" Or nTipe = "KABENG")) Then
    Else
        opByr(1).value = True
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
        cmdPilih.SetFocus
    End If
End Sub

Private Sub Text_Change(Index As Integer)
      Dim rsSearch As New ADODB.Recordset
1         If Index = 56 Or Index = 57 Or Index = 58 Or Index = 59 Or Index = 77 Or Index = 60 Or Index = 86 Or Index = 44 Or Index = 96 Or Index = 28 Or Index = 29 Or Index = 30 Or Index = 31 Or Index = 32 Or Index = 33 Or Index = 64 Or Index = 78 Then
              Dim rsPM As ADODB.Recordset
2             Select Case Index
                  Case 56, 57, 58, 59, 77, 60
3                     If text(0) <> "" Then
4                         If text(Index) = "" Then
5                             text(Index) = 0
6                         ElseIf CDbl(text(Index)) >= 0 Then
7                             text(Index) = Format(text(Index), "#0,#")
8                             text(Index).SelStart = Len(text(Index))
9                             text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
10                        End If
11                    End If
12                Case 86
13                    If text(86) <> ">>:>>" And text(65) <> ">>->>->>>>" Then
                          Dim rsFlag As ADODB.Recordset
14                        Set rsFlag = objO.FetchData("select top 1 Tanggal, Pause FROM TEMPPROGRESS WITH (NOLOCK) Where NoOrder <> '' and noOrder = '" & Trim(text(64)) & "' and FlagP = 'F' Order By Tanggal DESC, Pause DESC")
15                        If rsFlag.RecordCount > 0 Then
16                            text(63) = rsFlag.Fields(1).value
17                            text(62) = TglDDMMYYYY(rsFlag.Fields(0).value)
                              
18                            Call hitungEstimasiPenyerahan
19                        Else
20                            nWKT = hitungWaktu(IIf(chEM.value = Checked, 1, CDbl(ntRate)), text(86), text(65))
21                            text(63) = Left(nWKT, InStr(nWKT, "|") - 1)
22                            text(62) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
23                            Call hitungEstimasiPenyerahan
24                            If Trim(text(64)) = "" Then
25                                If text(89) < text(86) Then
26                                    Call hitungEstimasiPenyerahan
                                      
      '                                text(89) = Left(nWKT, InStr(nWKT, "|") - 1)
      '                                text(88) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
27                                End If
28                            End If
29                        End If
30                    End If
      '
31                Case 96
32                    ntRate = CDbl(text(Index))
33                    text(114) = ntRate
34                Case 64
35                    If Trim(text(64)) <> "" Then
36                        If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
37                            Set rsPM = objO.FetchDataACC("select Top 1 Status from SOORD WITH (NOLOCK) Where NoSO = '" & Trim(text(64)) & "'")
38                            If rsPM.RecordCount > 0 Then
39                                If rsPM.Fields(0).value <> "" And rsPM.Fields(0).value <> gSA Then
40                                    MsgBox "Sparepart sedang diproses oleh Partman (" & rsPM.Fields(0).value & ")!", vbInformation + vbOKOnly, "Supply Sparepart"
41                                    text(64) = ""
42                                    Call hapusForm
43                                    Exit Sub
44                                Else
45                                    Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSo = '" & Trim(text(64)) & "'")
46                                End If
47                            Else
48                                Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSo = '" & Trim(text(64)) & "'")
49                            End If
50                        End If
51                    Else
52                        If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
53                            If Trim(text(78)) <> "" Then
54                                Call objO.UpdateDataACC("Update SOORD Set Status ='" & gSA & "' Where nospk = '" & Trim(text(78)) & "'")
55                            End If
56                        End If
57                    End If
58                Case 78
59                    If Trim(text(78)) <> "" Then
60                        If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
61                            Set rsPM = objO.FetchDataACC("select Top 1 Status from SOORD WITH (NOLOCK) Where NoSPK = '" & Trim(text(78)) & "'")
62                            If rsPM.RecordCount > 0 Then
63                                If rsPM.Fields(0).value <> "" And rsPM.Fields(0).value <> gSA Then
64                                    MsgBox "Sparepart sedang diproses oleh Partman (" & rsPM.Fields(0).value & ") !", vbInformation + vbOKOnly, "Supply Sparepart"
65                                    text(78) = ""
66                                    Call hapusForm
67                                    Exit Sub
68                                Else
69                                    Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSPK = '" & Trim(text(78)) & "'")
70                                End If
71                            Else
72                                Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSPK = '" & Trim(text(78)) & "'")
73                            End If
74                        End If
75                    Else
76                        If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
77                            If Trim(text(64)) <> "" Then
78                                Call objO.UpdateDataACC("Update SOORD Set Status ='" & gSA & "' Where Noso = '" & Trim(text(64)) & "'")
79                            End If
80                        End If
81                    End If
82            End Select
83        End If
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
        
        Select Case Index
            Case 23, 33, 95
                SSTab1.Tab = 0
            Case 34, 43
                SSTab1.Tab = 1
            Case 44
                SSTab1.Tab = 2
            Case 54, 53
                SSTab1.Tab = 3
        End Select
    End If
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
'    text(Index).SetFocus
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                If cmdPilih.Tag = "No. Booking" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "BOOKING"
                    frmCari.Show (vbModal)
                ElseIf cmdPilih.Tag = "No. Rangka" Then
                    gMainKeyCari = text(Index)
                    'gSearchOption = "ORDER"
                    gSearchOption = "CUSTOMER2"
                    frmCari.Show (vbModal)
                ElseIf cmdPilih.Tag = "No. WO" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "ORDER"
                    gBantuan3 = ""
                    'gBantuan3 = IIf(gTipeUser = "SA", " AND FRONT = '" & Trim(gSA) & "'", "")
                    frmCari.Show (vbModal)
                End If
            Case 1
                gMainKeyCari = text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
            Case 21
                If opTpSvc(1).value Then
                    ' PROGRAM
                    gMainKeyCari = text(Index)
                    gSearchOption = "PROGRAM"
                    gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                    frmCari.Show (vbModal)
                ElseIf opTpSvc(2).value Then
                    ' WARRANTY
                    gMainKeyCari = text(Index)
                    gSearchOption = "WARRANTY"
                    frmCari.Show (vbModal)
                End If
            Case 28, 29, 30, 31, 32, 33
                gMainKeyCari = text(Index)
                gSearchOption = "KONDISI"
                gBantuan = Index - 27
                frmCari.Show (vbModal)
            Case 44
                If Trim(text(10)) <> "" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.Show (vbModal)
                Else
                    MsgBox "Model Kendaraan belum ada !!" & vbCr & "Lengkapi data master. "
                End If
            Case 46
                If Trim(text(44)) <> "" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "JOBLIST2"
                    gBantuan = text(10)
                    gBantuan2 = text(44)
                    frmCari.Show (vbModal)
                Else
                    text(44).SetFocus
                End If
            Case 48
                gMainKeyCari = text(Index)
                gSearchOption = "PART"
                gBantuan3 = Format(text(65), "dd/mm/yyyy")
                frmCariAcess.Show (vbModal)
            Case 119
                gMainKeyCari = text(Index)
                gSearchOption = "PART"
                gBantuan3 = Format(text(65), "dd/mm/yyyy")
                
                frmCariAcess.Show (vbModal)
            Case 50
                gMainKeyCari = text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ORDER"
                gTanggal = isnow
                gBantuan2 = IIf(Weekday(gTanggal) = 1, 6, Weekday(gTanggal) - 2) & WeekdayName(IIf(Weekday(gTanggal) = 1, 8, Weekday(gTanggal)) - 1)
                gBantuan3 = Format(gTanggal, "DD/MM/YYYY")
                gBantuan4 = ""
                frmCari.Show (vbModal)
            Case 103
                gMainKeyCari = text(Index)
                gSearchOption = "STATUSAKUN"
                frmCari.Show (vbModal)
            Case 107
                gMainKeyCari = text(Index)
                gSearchOption = "ALASANDTG"
                frmCari.Show (vbModal)
            Case 108
                gTanggal = TglDDMMYYYY(Format(gTglServer, "DD-MM-YYYY"))
                gMainKeyCari = text(Index)
                gSearchOption = "ANTRIAN"
                gBantuan = gSA
                gBantuan2 = Format(TglYYYYMMDD(text(65)), "MM/DD/YYYY")
                gBantuan3 = Format(isnow, "HH:NN")
                'Format(IIf(Trim(text(66)) <> "", Trim(text(66)), isnow), "HH:NN")
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
        cmdPilih.SetFocus
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If text(Index).BackColor = &HC0E0FF Then KeyAscii = 0
    Exit Sub
    Select Case Index
        Case 51, 47, 49, 58, 77
            KeyAscii = CekNumerik(KeyAscii)
            simpan = False
        Case 113
            KeyAscii = IIf(CekNumerik(KeyAscii) = 0, CekAlfabet(KeyAscii), CekNumerik(KeyAscii))
    End Select
    If KeyAscii = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0) = ""
        text(0).SetFocus
        cmdPilih.Tag = "No. Rangka"
        cmdPilih.Picture = ImageList1.ListImages(2).Picture
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim inputBantu As String
Dim boleh As Boolean
Dim noB As String
Dim kdB As String
Dim nonToyota As Boolean
Dim birthDay As String
Dim rsOpsi As ADODB.Recordset
Dim xDate As Date
xDate = isnow
    Select Case Index
        Case 0
            isWashing = False
            gTanggal = TglDDMMYYYY(Format(gTglServer, "DD-MM-YYYY"))
            If text(Index) <> "" Then
                boleh = True
                If simpan = False Then
                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                        boleh = False
                    End If
                End If
                If boleh Then
                    '----- proses hapus form ----
                    Call hapusForm
                    Call bukaForm
                    
                    partGrid.DataSource = Nothing
                    partGridFill.DataSource = Nothing
                    Select Case cmdPilih.Tag
                        Case "No. Booking"
                            If text(0) <> "" Then
                                Set rsOpsi = objO.FetchData("SELECT nofak FROM book WITH (NOLOCK) WHERE nofak = '" & Trim(text(Index)) & "' AND [DATE] >= '" & Format(xDate, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
                                If rsOpsi.EOF Then
                                    gMainKeyCari = text(Index)
                                    gSearchOption = "BOOKING"
                                    frmCari.Show (vbModal)
                                Else
                                    noB = Trim(text(Index))
                                    Call hapusForm
                                    Call bukaForm
                                    text(Index) = noB
                                    text(78) = noB
                                    SSTab1.Enabled = True
                                    opTpSvc(2).Enabled = False
                                    inputBantu = PanggilDataBooking(Trim(text(Index)))
                                End If
                            End If
                        Case "No. Rangka"
                            text(15) = Trim(text(Index))
                            noB = Trim(text(Index))
                            Call hapusForm
                            Call bukaForm
                            text(Index) = noB
                            If text(Index) <> "" Then
                                Select Case Len(text(Index))
                                    Case Is < 16
                                        MsgBox "No Rangka belum lengkap !!"
                                        If MsgBox("Unit NON Toyota ?", vbYesNo, "MERK KENDARAAN") = vbYes Then nonToyota = True
                                    Case Is > 17
                                        MsgBox "No Rangka kelebihan !!"
                                        text(Index).SelLength = Len(text(Index))
                                        text(Index).SetFocus
                                        Exit Sub
                                End Select
                            End If
                            
                            ' CHECK CUSTDAT
                            custKetemu = PanggilDataCustomer(Trim(text(Index)), IIf(cmdPilih.Tag = "No. WO", text(0), ""))
                            SSTab1.Enabled = True
                            cmdStall.Tag = "belum"
                                                        
                            If Not custKetemu Then
                                ' JK DI HO tidak ada
                                temp = MsgBox("Customer baru ?", vbDefaultButton1 + vbYesNo, "NEW CUSTOMER ?")
                                If temp = vbYes Then
                                    If rsTemp.State <> adStateClosed Then rsTemp.Close
                                    Set rsTemp = objO.FetchData("SELECT * FROM T_COUNTRY WITH (NOLOCK) WHERE KODE = '" & Trim(Left(text(Index), 3)) & "'")
                                    If rsTemp.EOF Then
                                        If Not nonToyota Then
                                            MsgBox "KODE NEGARA DARI NO RANGKA SALAH, ULANGI PENGISIAN !"
                                            text(Index).SetFocus
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    With frmtCustomer
                                        .text(50) = Left(text(Index), 3)
                                        If Len(text(Index)) > 4 Then
                                            .text(1) = Mid(text(Index), 4, 6)
                                        Else
                                            .text(1) = Mid(text(Index), 4, Len(text(Index)) - 4)
                                        End If
                                        If Len(text(Index)) > 9 Then
                                            .text(0) = Right(text(Index), Len(text(Index)) - 9)
                                        Else
                                            .text(0) = ""
                                        End If
                                        .Show (vbModal)
                                        text(1).SelLength = Len(text(1))
                                        text(1).SetFocus
                                    End With
                                Else
                                    gMainKeyCari = text(Index)
                                    gSearchOption = "CUSTOMER2"
                                    frmCari.Show (vbModal)
                                    text(0).SelLength = Len(text(0))
                                    text(0).SetFocus
                                End If
                            Else
                                If Not opTpSvc(2).value Then
                                    Call cekBooking(Trim(text(14)))
                                End If
                                
                            End If
                        
                        Case "No. WO"
                            ' CEK KE JOBDTA TANPA LIHAT KE HO
                            If text(Index) <> "" Then
                                Ket = ""
                                Set rsOpsi = objO.FetchData("SELECT NO_CHASSIS FROM jobdta WITH (NOLOCK) WHERE no_wo = '" & Trim(text(Index)) & "' AND deliv <> 'BTL' AND RPR_COND <> 'K'" & Ket)
                                If rsOpsi.EOF Then
                                    gMainKeyCari = text(Index)
                                    gSearchOption = "ORDER"
                                    gBantuan3 = ""
                                    frmCari.Show (vbModal)
                                Else
                                    noB = Trim(text(Index))
                                    Call hapusForm
                                    Call bukaForm
                                    text(Index) = noB
                                    Call PanggilDataOrder(Trim(rsOpsi("NO_CHASSIS")), text(Index))
                                    SSTab1.Enabled = True
                                End If
                            End If
                    End Select
                    
                    Call panggilWarranty(IIf(cmdPilih.Tag = "No. Booking", inputBantu, Trim(text(15))), text(10))
                    Call panggilHistory(IIf(cmdPilih.Tag = "No. Booking", inputBantu, Trim(text(15)))) ', text(14)
                    lastService = panggilLastSBE(IIf(cmdPilih.Tag = "No. Booking", inputBantu, Trim(text(15))), text(14))
                    
                    
                    
''                    frmEvent.Visible = False
''                    Dim ketEvent As String
''                    ketEvent = ""
''                    ketEvent = cekEvent(IIf(cmdPilih.Tag = "No. Booking", inputBantu, Trim(text(15))), text(14))
''                    If Trim(ketEvent) <> "" Then
''                        sgnEvent = Left(ketEvent, InStr(ketEvent, "|") - 1)
''                        frmEvent.Width = sgnEvent.Width + 200
''                        frmEvent.Height = sgnEvent.Height + 200
''                        frmEvent.Left = (frmOrderDMSNonDSS.Width / 2) - frmEvent.Width
''                        frmEvent.Top = (frmOrderDMSNonDSS.Height / 2) - frmEvent.Height
''                        frmEvent.Visible = True
''                        Set rsTemp = support.eventProgram(Replace(Trim(Mid(ketEvent, InStr(ketEvent, "|") + 1, Len(ketEvent) - InStr(ketEvent, "|") - 1)), vbNewLine, ""))
''                        If Not rsTemp.EOF Then
''                            opTpSvc(1).value = True
''                            text(21) = Trim(rsTemp("ID_PROGRAM"))
''                            text(22) = Trim(rsTemp("PROGRAM"))
''                        End If
''                    End If
                    
                    frmEvent.Visible = False
                    Dim ketEvent As String
                    ketEvent = ""
                    ketEvent = cekEvent(IIf(cmdPilih.Tag = "No. Booking", inputBantu, Trim(text(15))), text(14))
                    If Trim(ketEvent) <> "" Then
                        Dim nCounter As Integer
                        Dim lenText As Integer
                        Dim potText As String
                        Dim potAwal As String
                        
                        nCounter = 0
                        lenText = Len(Trim(ketEvent))
                        potText = Trim(ketEvent)
                        Do While lenText > 1
                            Pos = InStr(Trim(potText), vbNewLine)
                            If Pos > 0 Then
                                nCounter = nCounter + 1
                                potText = Mid(Trim(potText), Pos + 1, lenText - Pos)
                                lenText = lenText - Pos
                            End If
                        Loop
                        
                        potText = Trim(ketEvent)
                        sgnEvent = ""
                        Do While nCounter > 0
                            
                            sgnEvent = sgnEvent & Left(potText, InStr(potText, "|") - 1) & vbCrLf
                            potAwal = Left(potText, (InStr(Trim(potText), vbNewLine)))
                            potText = Mid(Trim(potText), (InStr(Trim(potText), vbNewLine)) + 1, Len(potText) - (InStr(Trim(potText), vbNewLine)))
                            
                            Set rsTemp = support.eventProgram(Mid(Trim(potAwal), InStr(potAwal, "|") + 1, Len(potAwal) - InStr(potAwal, "|") - 1))
                            If Not rsTemp.EOF Then
                                opTpSvc(1).value = True
                                text(21) = Trim(rsTemp("ID_PROGRAM"))
                                text(22) = Trim(rsTemp("PROGRAM"))
                            End If
                            nCounter = nCounter - 1
                        Loop
                        frmEvent.Width = sgnEvent.Width + 200
                        frmEvent.Height = sgnEvent.Height + 200
                        frmEvent.Left = (frmOrderDMSNonDSS.Width / 2) - frmEvent.Width
                        frmEvent.Top = (frmOrderDMSNonDSS.Height / 2) - frmEvent.Height
                        frmEvent.Visible = True
                    End If
                    
                    birthDay = ""
                    birthDay = birthDayReminder(Trim(text(15)))
                    If birthDay <> "" Then
                        If Abs(DateDiff("d", Format(birthDay, "mm/dd/yyyy"), xDate)) <= 5 Then
                            MsgBox "Selamat Ulang tahun Bp. / Ibu. / Sdr. " & cariNamaBirth(Trim(text(15))) & vbCr & "Pada Tanggal " & birthDay
                        End If
                    End If
                    
                    
                    If cmdPilih.Tag <> "No. WO" Then
                        If cekPiutang(Trim(text(1)), Trim(text(10)), IIf(cmdPilih.Tag = "No. Booking", inputBantu, Trim(text(15)))) Then
                            'Call hapusForm
                        End If
                        If Trim(cekMOU(text(10), text(15))) <> "" Then
                            opByr(2).value = True
                            cmdLengkap.Visible = True
                        End If
                        If capailimitMOU(hLimitMou, Trim(text(10)), Trim(text(15))) Then
                            MsgBox "Sisa piutang yang diijinkan sudah mencapai batasan " & hLimitMou & "%" & vbCr & "Segera melunasi hutang. "
                        End If
                    End If
                    simpan = False
                    Call hitungEstimasiPenyerahan
                End If
            End If
        Case 1
            If text(Index) <> "" Then
                Call isCurrentAR_Info
                If Log.LockAR = True And woOnProgres = False And Trim(text(64).text) = "" Then
''                    Me.PicARO.Visible = False
''                    If isOverdueAR = True Then
''                        Me.lblOto.Caption = " WO TER-BLOK !!!"
''                        Me.PicARO.Visible = True
''                        MsgBox "Ada AR Overdue !" & Chr(13) & Chr(13) & "WORK ORDER TIDAK DAPAT DIPROSES !"
''                        Exit Sub
''                    End If
''                    If isCurrentAR = True Then
''                        If nPersenAR >= 100 Then
''                            Me.lblOto.Caption = " WO TER-BLOK !!!"
''                            MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "WORK ORDER TIDAK DAPAT DIPROSES !"
''                        ElseIf nPersenAR >= 80 Then
''                            If objO.FetchIsValid_AR_Kabeng(Text(1).Text) = False Then
''                                Me.lblOto.Caption = "OTORITAS ADH"
''                                MsgBox "AR sudah >= 80% dari Batas AR !"
''                            End If
''                        End If
''                    End If
                    
                    If isOverdueAR = True Then
                        Me.PicARO.Visible = True
                        If objO.FetchIsValid_AR_Kabeng(text(1).text) = False Then
                            Me.lblOto.Caption = "OTORITAS KACAB"
                            Me.PicARO.Visible = True
                            MsgBox "Ada AR Overdue !" & Chr(13) & Chr(13) & "WORK ORDER TIDAK DAPAT DIPROSES !" & Chr(13) & "OTORITAS KEPALA CABANG", vbInformation + vbOKOnly
                            'Exit Sub
                        Else
                            MsgBox "AR OVERDUE" & Chr(13) & "TELAH DIVALIDASI OLEH KACAB", vbInformation + vbOKOnly
                            FrAR.Visible = False
                        End If
                    Else
                        If isCurrentAR = True Then
                            If nPersenAR >= 100 Then
                                If objO.FetchIsValid_AR_Kabeng(text(1).text) = False Then
                                    Me.lblOto.Caption = "OTORITAS KACAB"
                                    Me.PicARO.Visible = False
                                    MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "Work Order Tidak Dapat Di Proses !" & Chr(13) & Chr(13) & "OTORITAS KEPALA CABANG", vbInformation + vbOKOnly
                                Else
                                    MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "TELAH DIVALIDASI OLEH KACAB", vbInformation + vbOKOnly
                                End If
                            ElseIf nPersenAR >= 80 Then
                                Me.lblOto.Caption = "W A R N I N G"
                                Me.PicARO.Visible = False
                                If MsgBox("AR sudah >= 80% dari Batas AR !" & Chr(13) & "LANJUTKAN PROSES ?", vbQuestion + vbYesNo) = vbNo Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            
                    'Set rsOpsi = objO.FetchDataACC("SELECT * FROM armst WHERE kodeLgn = '" & text(Index) & "'")
                    '20/12/2012 Cust Non Axtif
                    Set rsOpsi = objO.FetchDataACC("SELECT * FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & text(Index) & "' and nonaktifflag = 0")
                    If rsOpsi.EOF Then
                        noB = Trim(text(0))
                        kdB = Trim(text(Index))
                        Call hapusForm
                        Call bukaForm
                        text(0) = noB
                        text(Index) = kdB
                        
                        gMainKeyCari = text(Index)
                        gSearchOption = "KODEPLG"
                        frmCariAcess.Show (vbModal)
                    Else
                        SSTab1.Enabled = True
                        text(2) = Nul(rsOpsi("KODEWIL"))
                        text(3) = Replace(Nul(rsOpsi("NAMALGN")), vbNewLine, "")
                        text(4) = Nul(rsOpsi("ALAMAT1"))
                        text(5) = Nul(rsOpsi("ALAMAT2"))
                        
                        Dim nDpn1 As String
                        Dim nTgh1 As String
                        Dim nBlk1 As String
                        nDpn1 = ""
                        nTgh1 = ""
                        nBlk1 = ""
                        Call pecahNoTelp(Trim(Nul(rsOpsi("TELEPON"))), nDpn1, nTgh1, nBlk1, ",")
                        text(6) = Trim(nDpn1)
                        text(7) = Trim(nTgh1)
                        text(8) = Trim(nBlk1)
                        
                        If Trim(text(14)) <> "" Or Trim(text(15)) <> "" Then
                            If Trim(text(14)) <> "" Then
                                Set rsTemp = objO.FetchData("SELECT no_CHASSIS FROM custdat WITH (NOLOCK) WHERE FS_KD_PLG = '" & text(Index) & "' AND no_polisi = '" & text(14) & "'")
                            ElseIf Trim(text(15)) <> "" Then
                                Set rsTemp = objO.FetchData("SELECT no_CHASSIS FROM custdat WITH (NOLOCK) WHERE FS_KD_PLG = '" & text(Index) & "' AND no_CHASSIS = '" & text(15) & "'")
                            End If
                            If rsTemp.EOF Then
                                If MsgBox("Kode Pelanggan Berubah. " & vbCr & "Akan mengupdate data Customer ? ", vbDefaultButton2 + vbYesNo, "KELENGKAPAN") = vbYes Then
                                    cmdKodePlg_Click
                                    text(2).SelLength = Len(text(2))
                                    text(2).SetFocus
                                End If
                            End If
                        End If
                            
                        Dim SQL As String
                        If cmdPilih.Tag = "No. Rangka" And Trim(text(0)) <> "" Then
                            SQL = "SELECT NO_CHASSIS FROM custdat WITH (NOLOCK) WHERE NO_CHASSIS = '" & text(0) & "' AND FS_KD_PLG = '" & text(Index) & "'"
                        Else
                            SQL = "SELECT NO_CHASSIS FROM custdat WITH (NOLOCK) WHERE FS_KD_PLG = '" & text(Index) & "' AND no_polisi = '" & text(14) & "' AND NO_CHASSIS LIKE '%" & Trim(text(15)) & "%'"
                        End If
                        
                        Set rsHlp = objO.FetchData(SQL)
                        If Not rsHlp.EOF Then
                            Call PanggilDataCustomer(rsHlp("NO_CHASSIS"), IIf(cmdPilih.Tag = "No. WO", text(0), ""))
                            If cmdPilih.Tag = "No. Rangka" Then
                                If Not opTpSvc(2).value Then
                                    Call cekBooking(Trim(text(14)))
                                End If
                            End If
                            Call objO.UpdateData("UPDATE TEMPJOBSUBLET SET NO_WO = '" & Trim(text(64)) & "', NO_CHASSIS = '" & Trim(text(15)) & "' WHERE NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "' AND NO_WO = '" & IIf(cmdPilih.Tag = "No. WO", Trim(text(64)), "") & "'")
                        End If
                    End If
                                
                If Not dataCustomerLengkap(Trim(text(15))) Then
                    MsgBox "Validasi Data Customer !!", vbCritical, "Kelengkapan"
                    Call cmdKodePlg_Click
                End If
                
                'lakukan pengecekan data booking / wo prediagnose
                Dim rsTempDiagnose As ADODB.Recordset
                If Trim(text(78)) <> "" Then
                    Set rsTempDiagnose = objO.FetchData("SELECT * FROM diagapproval WITH (NOLOCK) WHERE NOFAK = '" & Trim(text(78)) & "'")
                    If Not rsTempDiagnose.EOF Then
                        'jika booking tersebut termasuk dalam diagnose
                        opTpSvc(2).Enabled = True
                        opTpSvc(2).value = True
                        text(22) = rsTempDiagnose.Fields(0)
                        opTpSvc(0).Enabled = False
                        opTpSvc(1).Enabled = False
'                    Else
'                        opTpSvc(0).value = True
'                        opTpSvc(0).Enabled = True
'                        opTpSvc(1).Enabled = True
'                        opTpSvc(2).Enabled = False
'                        opTpSvc(2).value = False
'                        If Text(22) <> "" Then
'                            Text(22).Visible = False
'                        End If
                        
                    End If
                End If
            End If
        Case 21
            If text(Index) <> "" Then
                If opTpSvc(1).value Then
                    Set rsOpsi = objO.FetchData("SELECT A.ID_Program as Kode_Program,A.FLAGMODEL,A.PROGRAM  FROM HO_program A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                    If rsOpsi.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "PROGRAM"
                        gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                        frmCari.Show (vbModal)
                    Else
                        text(22) = rsOpsi("PROGRAM")
                        'Dim nMDL As Integer
                        'nMDL = objO.FetchData_Count("SELECT MDL FROM ho_promdl WHERE id_program = '" & Trim(rsOpsi("id_program")) & "'") ' AND mdl = '" & Trim(text(10)) & "'
                        'If nMDL > 0 Then
                            If Trim(rsOpsi("FLAGMODEL")) <> "A" Then
                                Set rsTemp = objO.FetchData("SELECT id_program FROM ho_promdl WITH (NOLOCK) WHERE id_program = '" & Trim(rsOpsi("kode_program")) & "' AND mdl = '" & Trim(text(10)) & "'")
                                If rsTemp.EOF Then
                                    MsgBox "Program yang dipilih tidak berlaku untuk Model kendaraan bersangkutan !!!"
                                    text(Index) = ""
                                    text(22) = ""
                                    text(Index).SelLength = Len(text(Index))
                                    SendKeys ("+{TAB}")
                                End If
                            End If
                        'End if
                    End If
                ElseIf opTpSvc(2).value Then
                    Set rsOpsi = objO.FetchData("SELECT WARRANTY,id_warranty FROM HO_warranty WITH (NOLOCK) WHERE id_warranty = '" & Trim(text(Index)) & "' and FS_LOG_DEL = ''")
                    If rsOpsi.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "WARRANTY"
                        frmCari.Show (vbModal)
                    Else
                        text(22) = rsOpsi("WARRANTY")
                        Set rsTemp = objO.FetchData("SELECT id_warranty FROM ho_warDtl WITH (NOLOCK) WHERE id_warranty = '" & rsOpsi("id_warranty") & "' AND mdl = '" & text(10) & "' AND no_rangka1 <= '" & text(15) & "' AND no_rangka2 >= '" & text(15) & "'")
                        If rsTemp.EOF Then
                           text(Index) = ""
                           text(22) = ""
                           text(Index).SelLength = Len(text(Index))
                           SendKeys ("+{TAB}")
                           Exit Sub
                        End If
                    End If
                End If
            End If
        Case 28, 29, 30, 31, 32, 33
            If text(Index) <> "" Then
                gBantuan = Index - 27
                Select Case gBantuan
                    Case 1
                        PSQL = "SELECT waktu FROM T_KONDISI WITH (NOLOCK) WHERE waktu = '" & text(Index) & "' ORDER BY id"
                    Case 2
                        PSQL = "SELECT frekuensi FROM T_KONDISI WITH (NOLOCK) WHERE frekuensi = '" & text(Index) & "' ORDER BY id"
                    Case 3
                        PSQL = "SELECT tempat FROM T_KONDISI WITH (NOLOCK) WHERE tempat = '" & text(Index) & "' ORDER BY id"
                    Case 4
                        PSQL = "SELECT warn FROM T_KONDISI WITH (NOLOCK) WHERE warn = '" & text(Index) & "' ORDER BY id"
                    Case 5
                        PSQL = "SELECT mulai FROM T_KONDISI WITH (NOLOCK) WHERE mulai = '" & text(Index) & "' ORDER BY id"
                    Case 6
                        PSQL = "SELECT selesai FROM T_KONDISI WITH (NOLOCK) WHERE selesai = '" & text(Index) & "' ORDER BY id"
                End Select
                Set rsOpsi = objO.FetchData(PSQL)
                If rsOpsi.EOF Then
                    text(Index) = ""
                Else
                    If text(Index) = "LAINNYA" Then
                        Dim nT As Integer
                        Select Case gBantuan
                            Case 1
                                text(74) = ""
                                text(74) = bKondisi(0)
                                nT = 74
                            Case 2
                                text(91) = ""
                                text(91) = bKondisi(1)
                                nT = 91
                            Case 3
                                text(92) = ""
                                text(92) = bKondisi(2)
                                nT = 92
                            Case 4
                                text(93) = ""
                                text(93) = bKondisi(3)
                                nT = 93
                            Case 5
                                text(94) = ""
                                text(94) = bKondisi(4)
                                nT = 94
                            Case 6
                                text(95) = ""
                                text(95) = bKondisi(5)
                                nT = 95
                                If Not text(nT).Visible Then
                                    SendKeys ("+{TAB}")
                                End If
                        End Select
                        If Not text(nT).Visible Then
                            SendKeys ("+{TAB}")
                        End If
                        text(nT).Visible = True
                        text(nT).SetFocus
                    End If
                End If
            End If
        Case 44
            If Trim(text(10)) <> "" Then
                If text(Index) <> "" And text(Index).BackColor <> &HC0E0FF Then
                    Set rsOpsi = objO.FetchData("SELECT DISTINCT JOB, REMARK FROM HO_FLATE WITH (NOLOCK) WHERE job = '" & Trim(text(Index)) & "' AND FS_LOG_DEL = ''")
                    If rsOpsi.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "JENIS_KERJA"
                        frmCari.Show (vbModal)
                    Else
                        text(45) = rsOpsi("REMARK")
                        If Trim(text(44)) = "PDS" Then
                            If chPDS = Unchecked Then
                                chPDS = Checked
                            End If
                        End If
                    End If
                Else
                    text(Index + 1) = ""
                End If
            Else
                MsgBox "Model Kendaraan belum ada !!" & vbCr & "Lengkapi data master. "
            End If
        Case 46
            If text(Index) <> "" Then
                If text(44) <> "" Then
                    If chPDS.value = 1 Then
                        If text(44).text = "SBE" Or text(44).text = "SBI" Or text(44).text = "GRP" Then
                            MsgBox "Unit PDS, tidak boleh dengan Job Order SBI/SBE/GRP !", vbInformation + vbOKOnly, "Validasi PDS"
                            text(44).text = ""
                            Exit Sub
                        End If
                    End If
                    Set rsOpsi = objO.FetchData("select DISTINCT FS_ITEM, RATE from HO_FLATE WITH (NOLOCK) where fs_item = '" & Trim(text(Index)) & "' AND mdl = '" & Trim(text(10)) & "' AND JOB = '" & Trim(text(44)) & "'")
                    If rsOpsi.BOF Then
                        If fOther Then
                            text(47).SelLength = Len(text(47))
                        Else
                            Cancel = True
                            gMainKeyCari = text(Index)
                            gSearchOption = "JOBLIST2"
                            gBantuan = text(10)
                            gBantuan2 = text(44)
                            frmCari.Show (vbModal)
                        End If
                    Else
                        If Trim(rsOpsi("FS_ITEM")) = "OTHER" Then
                            fOther = True
                            text(Index) = ""
                            text(Index).Mask = ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                            text(47).Visible = True
                            text(47).Enabled = True
                            text(Index).SelLength = Len(text(Index))
                            SendKeys ("+{TAB}")
                        Else
                            fOther = False
                            text(47).Visible = False
                            Jrate = CDbl(rsOpsi("Rate"))
                        End If
                    End If
                Else
                    MsgBox "Job Order Harus diisi", vbCritical, "KELENGKAPAN"
                    text(44).SelLength = Len(text(44))
                    text(44).SetFocus
                    Exit Sub
                End If
            End If
        Case 47
            If text(Index) = "" Then
                text(Index) = 0
            Else
                If text(44) = "" Then
                    MsgBox "Isi Tipe Job Order !!!"
                    text(44).SelLength = Len(text(44))
                    text(44).SetFocus
                    Exit Sub
                End If
            End If
            
                     
            Case 119
            If text(Index) <> "" Then
                    Set rsOpsi = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
                    If rsOpsi.EOF Then
                        hNamaPart = text(Index)
                        nStok = -1
                        hHrgPart = 0
                        hSatPart = ""
                        If rsTemp.State Then rsTemp.Close
                        Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(text(Index), 1) & "' AND TIPE = 'O'")
                        If Not rsTemp.EOF Then
                            hOliPart = True
                        Else
                            hOliPart = False
                        End If
                    Else
                        hNamaPart = rsOpsi("NAMABARANG")
                        nStok = rsOpsi("QTYSTOCK")
                        hHrgPart = rsOpsi("HARGAJUALA")
                        hSatPart = rsOpsi("KODESATUAN")
                        If rsTemp.State Then rsTemp.Close
                        Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(text(Index), 1) & "' AND TIPE = 'O'")
                        If Not rsTemp.EOF Then
                            hOliPart = True
                        Else
                            hOliPart = False
                        End If
                    End If
            End If
            
        Case 48
            If text(Index) <> "" Then
                Set rsOpsi = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
                If rsOpsi.EOF Then
                    gMainKeyCari = text(Index)
                    text(Index) = ""
                    gSearchOption = "PART"
                    frmCariAcess.Show (vbModal)
                    nStok = gStok
                    hNamaPart = gNamaPart
                    hSatPart = gSatPart
                    hHrgPart = gHargaJual
                    If rsTemp.State Then rsTemp.Close
                    Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(text(Index), 1) & "' AND TIPE = 'O'")
                    If Not rsTemp.EOF Then
                        hOliPart = True
                    Else
                        hOliPart = False
                    End If
                Else
                    hNamaPart = rsOpsi("NAMABARANG")
                    nStok = rsOpsi("QTYSTOCK")
                    hHrgPart = rsOpsi("HARGAJUALA")
                    hSatPart = rsOpsi("KODESATUAN")
                    If rsTemp.State Then rsTemp.Close
                    Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(text(Index), 1) & "' AND TIPE = 'O'")
                    If Not rsTemp.EOF Then
                        hOliPart = True
                    Else
                        hOliPart = False
                    End If
                End If
            End If
            
        
        Case 49
            If Trim(text(Index)) = "" Then
                text(Index) = 0
                SendKeys ("+{TAB}")
                Exit Sub
            Else
                If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                    If hNamaPart <> "" And Trim(text(48)) <> "" Then
                        If text(Index) > 0 Then
                            If nStok >= CDbl(text(Index)) Then
                                'ISI RS PART
                                Call isiPart(text(48), hNamaPart, hSatPart, CDbl(text(Index)), hOliPart, "", NewId)
                                text(48) = ""
                                text(Index) = 0
                                text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                            Else
                                MsgBox "Stok item tidak mencukupi !!" & vbCr & "Stok : " & nStok, vbCritical, "STOK ITEM"
                                text(Index) = nStok
                                text(Index).SelLength = Len(text(Index))
                                SendKeys ("+{TAB}")
                                Exit Sub
                            End If
                        End If
                    Else
                        text(48).SetFocus
                        SendKeys ("+{TAB}")
                        Exit Sub
                    End If
                Else
                    If Trim(text(48)) <> "" Then
                        If text(Index) > 0 Then
                            If nStok >= CDbl(text(Index)) Then
                                'ISI RS PART
                                Call isiPart(text(48), hNamaPart, hSatPart, CDbl(text(Index)), hOliPart, "", NewId)
                                text(48) = ""
                                text(Index) = 0
                                text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                            Else
                                MsgBox "Stok item tidak mencukupi !!" & vbCr & "Stok : " & nStok, vbCritical, "STOK ITEM"
                                text(Index) = nStok
                                text(Index).SelLength = Len(text(Index))
                                SendKeys ("+{TAB}")
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
            End If

        Case 120
            If Trim(text(Index)) = "" Then
                text(Index) = 0
                SendKeys ("+{TAB}")
                Exit Sub
            Else
                If hNamaPart <> "" And Trim(text(119)) <> "" Then
                    If text(Index) > 0 Then
                        If nStok >= CDbl(text(Index)) Then
                            'ISI RS PART
                            Call isiPart(text(119), hNamaPart, hSatPart, CDbl(text(Index)), hOliPart, "", NewId)
                            text(119) = ""
                            text(Index) = 0
                            text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                        Else
                            If nStok <> -1 Then
                                'If MsgBox("Stok item tidak mencukupi !!" & vbCr & "Stok : " & nStok & vbCr & "Tetap inputkan data ?", vbQuestion + vbYesNo, "STOK ITEM") = vbYes Then
                                    Call isiPart(text(119), hNamaPart, hSatPart, CDbl(text(Index)), hOliPart, "", NewId)
                                    text(119) = ""
                                    text(Index) = 0
                                    text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                                'Else
                                '    text(Index) = nStok
                                '    text(Index).SelLength = Len(text(Index))
                                '    SendKeys ("+{TAB}")
                                '    Exit Sub
                                'End If
                            Else
                                If MsgBox("Stok item tidak terdaftar !!" & vbCr & "Inputkan data Nama Part By SA ?", vbQuestion + vbYesNo, "STOK ITEM") = vbYes Then
                                    Call isiPart("", "", "ITEM", CDbl(text(Index)), hOliPart, Trim(text(119)), NewId)
                                    text(119) = ""
                                    text(Index) = 0
                                    text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                                Else
                                    text(Index) = "0"
                                    text(Index).SelLength = Len(text(Index))
                                    SendKeys ("+{TAB}")
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Else
                    text(119).SetFocus
                    SendKeys ("+{TAB}")
                    Exit Sub
                End If
            End If
            
        Case 50
            If text(Index) <> "" Then
                If text(46) <> "" And text(44) <> "" Then
                    'LIHAT LAGI
                    Set rsOpsi = objO.FetchData("SELECT isnull(MENITKERJA,0) AS MENITKERJA, ISNULL(TERPAKAI,0) AS TERPAKAI, ISNULL(ALOKASI, 0) AS ALOKASI, LEVMECH, [GROUP] FROM mekanik A WITH (NOLOCK) LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH AND B.HARI = '" & IIf(Weekday(xDate) = 1, 6, Weekday(xDate) - 2) & WeekdayName(IIf(Weekday(xDate) = 1, 8, Weekday(xDate)) - 1) & "' AND B.TANGGAL = '" & Format(xDate, "DD/MM/YYYY") & "' WHERE A.mech = '" & text(Index) & "'")
                    If rsOpsi.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "MEKANIK"
                        gBantuan = "ORDER"
                        'LIHAT LAGI
                        gBantuan2 = IIf(Weekday(xDate) = 1, 6, Weekday(xDate) - 2) & WeekdayName(IIf(Weekday(xDate) = 1, 8, Weekday(xDate)) - 1)
                        gBantuan3 = Format(xDate, "DD/MM/YYYY")
                        gBantuan4 = ""
                        frmCari.Show (vbModal)
                    Else
                        text(82) = rsOpsi("MENITKERJA") - rsOpsi("ALOKASI")
                        text(106) = Trim(rsOpsi("GROUP"))
                        If CDbl(text(82)) <= 0 Then
                            text(82) = ""
                            text(50) = ""
                            text(50).SelLength = Len(text(50))
                            text(50).SetFocus
                            MsgBox "Waktu Mekanik tidak mencukupi !!!" & vbCr & "Pilih Mekanik Lain"
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Case 51
            If Trim(text(20)) = "" Then text(20) = 0
            If text(Index) = "" Then
                text(Index) = 0
            ElseIf CDbl(text(Index)) < CDbl(text(20)) Then
                MsgBox "KM sekarang harus lebih besar dari KM terakhir !!" & vbCr & "Ulangi Pengisian & Cek Rangka Kendaraan", vbCritical, "KM"
                text(Index) = 0
                text(Index).SetFocus
            End If
        Case 74
            If text(Index) <> "" Then
                bKondisi(0) = text(Index)
            End If
        Case 88
            If CekTgl(text(Index)) Then
                If DateDiff("d", Format(xDate, "dd/mm/yyyy"), Format(text(Index), "DD/MM/YYYY")) < 0 Then
                    text(Index) = TglDDMMYYYY(Format(xDate, "dd/mm/yyyy"))
                End If
            ElseIf Not CekTgl(text(Index)) Then
                text(Index) = TglDDMMYYYY(Format(xDate, "dd/mm/yyyy"))
            End If
        Case 89
            If Not CekJam(text(Index)) Then
                text(Index) = Format(xDate, "HH:NN")
            Else
                
                If DateDiff("d", Format(text(88), "DD/MM/YYYY"), Format(text(65), "DD/MM/YYYY")) = 0 Then
                    If DateDiff("s", Format(text(63), "HH:NN"), Format(text(Index), "HH:NN")) < 0 Then
                        MsgBox "Waktu Penyerahan tidak boleh lebih kecil dari Waktu Selesai ! "
                        text(Index) = Format(text(63), "HH:NN")
                        
                    End If
                End If
            End If
                Case 91
            If text(Index) <> "" Then
                bKondisi(1) = text(Index)
            End If
        Case 92
            If text(Index) <> "" Then
                bKondisi(2) = text(Index)
            End If
        Case 93
            If text(Index) <> "" Then
                bKondisi(3) = text(Index)
            End If
        Case 94
            If text(Index) <> "" Then
                bKondisi(4) = text(Index)
            End If
        Case 95
            If text(Index) <> "" Then
                bKondisi(5) = text(Index)
            End If
        Case 103
            If text(Index) <> "" Then
                Set rsOpsi = objO.FetchData("SELECT TIPE FROM T_STATUSAKUN WITH (NOLOCK) WHERE TIPE = '" & Trim(text(Index)) & "'")
                If rsOpsi.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "STATUSAKUN"
                    frmCari.Show (vbModal)
                End If
            End If
        Case 108
            If text(Index) <> "" Then
                If text(109) <> "" Then
                    Dim txtHelp As String
                    txtHelp = ""
                    If cmdPilih.Tag = "No. WO" Then
                        txtHelp = " AND WO <> '" & Trim(text(64)) & "'"
                    End If
                    Set rsOpsi = objO.FetchData("SELECT TIPE FROM LISTANTRIAN WITH (NOLOCK) WHERE TIPE = '" & Trim(text(109)) & "' AND QUEUE = '" & Trim(text(Index)) & "' AND TANGGAL = '" & Format(TglYYYYMMDD(text(65)), "MM/DD/YYYY") & "'" & txtHelp)
                    If Not rsOpsi.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "ANTRIAN"
                        gBantuan = gSA
                        gBantuan2 = Format(TglYYYYMMDD(text(65)), "MM/DD/YYYY")
                        gBantuan3 = Format(isnow, "HH:NN") 'Format(IIf(Trim(text(66)) <> "", Trim(text(66)), xDate), "HH:NN")
                        
                        frmCari.Show (vbModal)
                    End If
                Else
                    gMainKeyCari = text(Index)
                    gSearchOption = "ANTRIAN"
                    gBantuan = gSA
                    gBantuan2 = Format(TglYYYYMMDD(text(65)), "MM/DD/YYYY")
                    gBantuan3 = Format(isnow, "HH:NN") 'Format(IIf(Trim(text(66)) <> "", Trim(text(66)), xDate), "HH:NN")
                    frmCari.Show (vbModal)
                End If
            End If
    End Select
End Sub

Private Sub cekBooking(nPol As String)
Dim rsBooking As ADODB.Recordset
    Set rsBooking = objO.FetchData("SELECT NOFAK, DATE FROM BOOK WITH (NOLOCK) WHERE NO_POLISI = '" & Trim(nPol) & "' AND [DATE] >= '" & Format(isnow, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
    If Not rsBooking.EOF Then
        If MsgBox("Telah dibuat booking untuk tanggal " & TglDDMMMMYYYY(rsBooking("DATE")) & vbCr & "Akan menginput Order ?", vbYesNo) = vbYes Then
            
            cmdPilih.Tag = "No. Booking"
            cmdPilih.Picture = ImageList1.ListImages(1).Picture

            text(0) = Trim(rsBooking("NOFAK"))
            text(78) = Trim(rsBooking("NOFAK"))
            inputBantu = PanggilDataBooking(Trim(text(0)))
            text(1).SetFocus
        End If
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Trim(Text1(Index)) = "" Then
        Text1(Index) = "                              "
    End If
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

''====================================00000000000000000000000===============================================
''====================================00000000000000000000000===============================================

Private Function lastPrediag(noRangka As String) As Boolean
    
    On Error GoTo vbError
    lastPrediag = False
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND NO_WO = '' AND DELIV <> 'BTL' AND (VALIDTWC <> 'NOK')")
    If Not rsHlp.EOF Then
        lastPrediag = True
    End If
    Exit Function

vbError:
    lastPrediag = False
    Exit Function
End Function

Private Function allowPrediagToWO(noRangka As String) As Boolean
    
    On Error GoTo vbError
    allowPrediagToWO = False
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND DELIV <> 'BTL' AND VALIDTWC = 'OKE' AND NO_WO = ''")
    If Not rsHlp.EOF Then
        allowPrediagToWO = True
    End If
    Exit Function

vbError:
    allowPrediagToWO = False
    Exit Function
End Function

Private Function allowNewPrediag(noRangka As String) As Boolean
    
    On Error GoTo vbError
    allowNewPrediag = True
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND DELIV <> 'BTL' AND (VALIDTWC <> 'NOK' AND VALIDTWC <> 'OKE')")
    If Not rsHlp.EOF Then
        If CDbl(rsHlp.RecordCount) >= 3 Then
            allowNewPrediag = False
        End If
    End If
    Exit Function

vbError:
    allowNewPrediag = True
    Exit Function
End Function

Private Sub panggilPreDiagnosis(nDiag As String, nRgk As String, tipe As String)
Dim rsTmp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
Dim rsOli As ADODB.Recordset
Dim n As Integer
        
    Set rsHlp2 = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(nDiag) & "' AND NO_CHASSIS = '" & Trim(nRgk) & "' AND USERDEL = ''")
    If rsHlp2.EOF Then
        Set rsTemp = objO.FetchData("SELECT a.*,b.gsm,b.cdma,b.tgl_rep,b.rcdate,b.fs_km_akhir, b.Ddate, b.FS_kd_dealer, b.Fs_nm_dealer " _
                                 & "FROM diagdta a WITH (NOLOCK) INNER JOIN custdat b WITH (NOLOCK) ON a.no_chassis = b.no_chassis " _
                                 & "WHERE a.no_chassis = '" & nRgk & "' AND a.NO_DIAG = '" & nDiag & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' AND VALIDTWC = '" & IIf(tipe = "Edit", "", "OKE") & "'")
                                 
    
        If Not rsTemp.EOF Then
            FrDiag.Left = 0
            If rsTemp("RPR_COND") <> "K" Then
                text(1) = rsTemp("FS_KD_PLG")
                text(2) = rsTemp("FS_KD_WIL")
                text(3) = rsTemp("CUSTOMER")
                text(4) = Left(rsTemp("CUST_ADRES"), 40)
                text(5) = rsTemp("KOTA")
                text(6) = rsTemp("TELPON")
                text(7) = rsTemp("GSM")
                text(8) = rsTemp("CDMA")
                text(9) = rsTemp("FS_TH_MOBIL")
                text(10) = rsTemp("MDL")
                text(11) = rsTemp("MODEL")
                text(12) = rsTemp("FMDL")
                text(13) = rsTemp("COLOR")
                text(14).Mask = ""
                text(14) = rsTemp("NO_POLISI")
                cmdGantiNoPol.Visible = True
                text(15) = Nul(rsTemp("NO_CHASSIS"))
                text(16).Mask = ""
                text(16) = TglDDMMYYYY(rsTemp("DOTAM"))
                text(16).Mask = "99-99-9999"
                text(79) = TglDDMMYYYY(rsTemp("DDATE"))
                text(118).Mask = ""
                text(118) = TglDDMMYYYY(rsTemp("DDATE"))
                text(118).Mask = "99-99-9999"
                text(17).Mask = ""
                text(17) = TglDDMMYYYY(rsTemp("TGL_REP"))
                text(18).Mask = ""
                text(18) = TglDDMMYYYY(rsTemp("RCDATE"))
                text(19) = "" 'rsTemp("")
                text(20) = Format(CDbl(rsTemp("FS_KM_AKHIR")), "0#,#")
                If Trim(text(20)) = "" Then text(20) = 0
                text(23) = rsTemp("REQ1")
                text(24) = rsTemp("REQ2")
                text(25) = rsTemp("REQ3")
                text(26) = rsTemp("REQ4")
                text(27) = rsTemp("REQ5")
                text(90) = cekMOU(rsTemp("MDL"), Nul(rsTemp("NO_CHASSIS")))
                cmdLengkap.Visible = False
                If Trim(text(90)) <> "" Then
                    cmdLengkap.Visible = True
                End If
                chPDS = IIf(rsTemp("PDS") = 1, Checked, Unchecked)
                chEM = IIf(rsTemp("CODE") = 1, Checked, Unchecked)
                chSBNP = IIf(rsTemp("KD_K") = 1, Checked, Unchecked)
                text(76) = rsTemp("NO_ENGINE")
                text(104) = rsTemp("FRONT")
                text(105) = rsTemp("TELPANCUST")
                text(106) = rsTemp("GROUP")
                
                Dim hlp As String
                Set rsTmp = objO.FetchData("SELECT * FROM DIAGDTL WITH (NOLOCK) WHERE NO_diag = '" & rsTemp("NO_DIAG") & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                
                    ' WALK AROUND CHECK
                    ' %%%% KONDISI KENDARAAN %%%%
                    text(34) = Nul(rsTmp("WAK1"))
                    text(35) = Nul(rsTmp("WAK2"))
                    text(36) = Nul(rsTmp("WAK3"))
                    text(37) = Nul(rsTmp("WAK4"))
                    text(38) = Nul(rsTmp("WAK5"))
                    
                    ' %%%% BARANG BERHARGA %%%%
                    text(39) = Nul(rsTmp("WAE1"))
                    text(40) = Nul(rsTmp("WAE2"))
                    text(41) = Nul(rsTmp("WAE3"))
                    text(42) = Nul(rsTmp("WAE4"))
                    text(43) = Nul(rsTmp("WAE5"))
                    
                    'cboAlasan = Trim(rsTmp("DATANG"))
                    text(107) = Trim(rsTmp("DATANG"))
                    text(52) = Trim(rsTmp("DTLDATANG"))
                    cboPelanggan = Trim(rsTmp("KARAKTER"))
                    text(75) = Trim(rsTmp("DTLKARAKTER"))
                    
                    ' ISI KELUHAN
                    For i = 1 To 5
                        hlp = "CMPL" & i
                        Text1(-1 + i) = Trim(rsTmp(hlp))
                    Next
                    
                    chUse.value = IIf(rsTmp("HEAVYUSE") = 1, Checked, Unchecked)
                    
                    ' BBM
                    n = IIf(Nul(rsTmp("BBM")) = "", 4, Nul(rsTmp("BBM")))
                    opBBM(n).value = True
                    
                    ' FIELD _ FIELD LAIN
                    n = IIf(Nul(rsTmp("GANTIPART")) = "", 0, Nul(rsTmp("GANTIPART")))
                    opGanti(n).value = True
                    
                    n = IIf(Nul(rsTmp("TUNGGUSVC")) = "", 0, Nul(rsTmp("TUNGGUSVC")))
                    opSvc(n).value = True
                    
                    n = IIf(Nul(rsTmp("WASHING")) = 0, 1, 0)
                    opWash(n).value = True
                    
                    Dim nBYR As Integer
                    nBYR = IIf(Nul(rsTmp("PEMBAYARAN")) = "CS", 0, IIf(rsTmp("PEMBAYARAN") = "CR", 1, 2))
                    If nBYR = 2 Then Updated = False
                End If
                
                ' %%%% KOORDINAT %%%%
                nk = 0
                bodyClear = True
                Set rsTmp = objO.FetchData("SELECT KOORDX,KOORDY FROM DIAGKoord WITH (NOLOCK) WHERE no_DIAG = '" & rsTemp("NO_DIAG") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        imgTanda(nk).Top = 600 + rsTmp("KOORDY") - (imgTanda(nk).Height / 2)
                        imgTanda(nk).Left = 600 + rsTmp("KOORDX") - (imgTanda(nk).Width / 2)
                        imgTanda(nk).Visible = True
                        KorX(nk) = rsTmp("KOORDX")
                        KorY(nk) = rsTmp("KOORDY")
                        nk = nk + 1
                        rsTmp.MoveNext
                    Loop
                    bodyClear = False
                End If
        
                ' TAMPIL CONDITION
                Dim kCond As String
                Dim kBant As String
                Set rsTmp = objO.FetchData("SELECT * FROM DIAGDTL WITH (NOLOCK) WHERE NO_DIAG = '" & rsTemp("NO_DIAG") & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    For i = 1 To 6
                        hlp = "COND" & i
                        kBant = ""
                        kCond = ""
                        If rsTmp(hlp) <> "" Then
                            If Left(rsTmp(hlp), 3) = "L| " Then
                                kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
                                kCond = "LAINNYA"
                            Else
                                kCond = Trim(rsTmp(hlp))
                                kBant = ""
                            End If
                        End If
                        text(27 + i) = kCond
                        bKondisi(i - 1) = kBant
                    Next
                End If
                
                ' ISI JOBORDER
                ntRate = 0
                text(114) = 0
                myChassis = Trim(text(15))
                For i = 1 To 8
                    tmp1 = "REP" & i
                    tmp2 = "RTE" & i
                    tmp3 = "MEKANIK" & i
                    DGrid.TextMatrix(i, 1) = rsTemp(tmp1)
                    If rsTemp(tmp1) <> "" Then
                        baris = baris + 1
                        Set rsTmp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(tmp1) & "'")
                        If Not rsTmp.EOF Then
                            DGrid.TextMatrix(i, 3) = False
                            DGrid.TextMatrix(i, 4) = rsTmp("JOB")
                        Else
                            DGrid.TextMatrix(i, 3) = True
                        End If
                        DGrid.TextMatrix(i, 2) = rsTemp(tmp2)
                        DGrid.TextMatrix(i, 5) = hitungLabour(CDbl(rsTemp(tmp2)), text(11))
                        DGrid.TextMatrix(i, 6) = rsTemp(tmp3)
                        ntRate = ntRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(rsTemp(tmp1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                        text(114) = ntRate
                        If isJualNRM = True Then
                            Me.JualNRM.Caption = nJualNRM
                        Else
                            Me.JualNRM.Caption = ""
                        End If
                    End If
                Next
                
                ' ISI SUBLET
                Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & nDiag & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND KODEITEM <> 'A'")
                Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, NO_WO, NOINV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, HARGA, NONOTA, KETERANGAN, FLAG, NO_POLISI,PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR, '00:00' AS ESTWAKTU FROM DIAGSUBLET WITH (NOLOCK) WHERE NO_DIAG = '" & nDiag & "' AND NO_CHASSIS = '" & Trim(text(15)) & "' AND KODEITEM <> 'A'")
                gPO = objO.FetchData_String("SELECT NOPO FROM DIAGSUBLET WITH (NOLOCK) WHERE NO_DIAG = '" & nDiag & "' AND NO_CHASSIS = '" & Trim(text(15)) & "'")
                ' ISI FU
                Set rsTmp = objO.FetchData("SELECT * FROM DIAGDTL WITH (NOLOCK) WHERE NO_DIAG = '" & rsTemp("NO_DIAG") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
                If Not rsTmp.EOF Then
                    If Trim(rsTmp("TELPFU")) <> "" And Not IsNull(rsTmp("TELPFU")) Then
                        Call objO.InsertData("INSERT INTO TEMPFU SELECT NO_WO, NO_CHASSIS, TGLFU, JAMFU, TELPFU, KETFU FROM DIAGDtl WITH (NOLOCK) " _
                                     & "WHERE NO_diag = '" & rsTemp("NO_DIAG") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' AND " _
                                     & "NOT EXISTS (SELECT * FROM TEMPFU WITH (NOLOCK) WHERE NO_WO = '" & rsTemp("NO_DIAG") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "')")
                    End If
                    If Trim(rsTmp("TIMESERAH")) <> "" Then
                        text(88) = TglDDMMYYYY(rsTmp("DATESERAH"))
                        text(89) = Trim(rsTmp("TIMESERAH"))
                    End If
                End If
                
                ' ISI PART
                If rsGridPart.State = adStateClosed Then createTblPart
                Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, SAT_ITEM, NamaBySA FROM DIAGPart A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & rsTemp("NO_CHASSIS") & "' AND no_DIAG = '" & rsTemp("NO_DIAG") & "' AND ADDJOB = 0")
                If Not rsTmp.EOF Then
                    i = 1
                    n = 69
                    Do While Not rsTmp.EOF
                        With rsGridPart
                            .AddNew
                            !KodeItem = rsTmp("KODE_ITEM")
                            !NAMAITEM = rsTmp("NAMA_ITEM")
                            !SATITEM = rsTmp("SAT_ITEM")
                            !qty = CDbl(rsTmp("QTY"))
                            !ganti = IIf(rsTmp("GANTI") <> "N", True, False)
                            !BAWA = IIf(rsTmp("BAWA") <> "N", True, False)
                            !oli = IIf(rsTmp("KDOLI") <> "", True, False)
                            !namabySA = rsTmp("NamaBySA")
                            !nID = i + 1
                            .Update
                            i = i + 1
                        End With
                        If rsTmp("PARTBAWA") <> "" Then
                            text(n) = rsTmp("PARTBAWA")
                            n = n + 1
                        End If
                    rsTmp.MoveNext
                    Loop
                    Call HeadColumn
                    Set partGrid.DataSource = rsGridPart
                    Set partGridFill.DataSource = rsGridPart
                    FrDiag.Left = -9000
                End If
                
                cmdWarranty.Enabled = False
                If rsTemp("TPSVC") = "W" Then
                    opTpSvc(0).Enabled = False
                    opTpSvc(1).Enabled = False
                    opTpSvc(2).value = True
                    opTpSvc(2).Enabled = True
                    text(21) = "Pre"
                    text(22) = Trim(rsTemp("NO_DIAG"))
                ElseIf rsTemp("TPSVC") = "P" Then
                    opTpSvc(1).value = True
                    text(21) = rsTemp("KODETIPE")
                    text(22) = objO.FetchData_String("SELECT PROGRAM FROM HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & Trim(text(21)) & "'")
                ElseIf rsTemp("TPSVC") = "" Then
                    opTpSvc(0).value = True
                    text(21) = ""
                End If
                
                text(102) = rsTemp("ANCUST")
                text(50) = rsTemp("MECH")
                text(51) = rsTemp("KM")
                Txt(53) = Left(rsTemp("CATAT"), 250)
                Txt(87) = Right(rsTemp("CATAT"), Len(rsTemp("CATAT")) - IIf(Len(rsTemp("CATAT")) > 250, 250, Len(rsTemp("CATAT"))))
                Txt(54) = rsTemp("REMARK1")
                Txt(55) = rsTemp("SUGGEST")
                text(56) = Format(CDbl(rsTemp("LABOR")), "0#,#")
                text(57) = Format(CDbl(rsTemp("PART")), "0#,#")
                text(58) = Format(CDbl(rsTemp("MTRL")), "0#,#")
                text(59) = Format(CDbl(rsTemp("OIL")), "0#,#")
                text(60) = Format(CDbl(rsTemp("SBLT")), "0#,#")
                text(61) = Format(CDbl(rsTemp("TOTAL")), "0#,#")
                text(61) = Format(CDbl(text(56)) + CDbl(text(57)) + CDbl(text(58)) + CDbl(text(59)) + CDbl(text(60)) + CDbl(text(77)), "#0,#")
                
                '' NO ANTRIAN
                Set rsTmp = objO.FetchData("SELECT * FROM LISTANTRIAN WITH (NOLOCK) WHERE WO = '" & Trim(text(0)) & "'")
                If Not rsTmp.EOF Then
                    text(108) = Trim(rsTmp("QUEUE"))
                    text(109) = Trim(rsTmp("TIPE"))
                    noAntriLama = Trim(text(108))
                End If
                
                ' HITUNG NILAI PART & OLI
                If rsGridPart.State <> adStateClosed Then
                    text(59) = 0
                    text(57) = 0
                    If Not rsGridPart.EOF Then
                    rsGridPart.MoveFirst
                    Do While Not rsGridPart.EOF
                            Set rsOli = objO.FetchDataACC("SELECT hargaJualA FROM inmst WITH (NOLOCK) WHERE kodeItem = '" & rsGridPart("KODEITEM") & "'")
                            If Not rsOli.EOF Then
                                If rsGridPart("OLI") Then
                                    text(59) = text(59) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                                Else
                                    text(57) = text(57) + (CDbl(rsOli("HARGAJUALA")) * rsGridPart("QTY"))
                                End If
                            End If
                            rsOli.Close
                        rsGridPart.MoveNext
                    Loop
                    End If
                End If
                
                'CboStatus = rsTemp("STATUS")
                text(103) = rsTemp("STATUS")
                text(64) = Nul(rsTemp("NO_WO"))
                text(65) = TglDDMMYYYY(isnow)  'TglDDMMYYYY(rsTemp("DATE"))
                text(66) = rsTemp("TIME")
                text(67) = rsTemp("NOINV")
                text(78) = Nul(rsTemp("NOFAK"))
                ChkReturnOrder = IIf(Trim(rsTemp("FN_RETURN")) = 1, Checked, Unchecked)
                ChkReturnOrder.Enabled = True
                chTHS = IIf(rsTemp("THS") = 1, Checked, Unchecked)
                simpan = False
                opByr(nBYR).value = True
                
                text(44).BackColor = &H80000005
                text(46).BackColor = &H80000005
                
                If tipe = "Edit" Then
                    ' CHECK STALL
                    Set rsTmp = objO.FetchData("SELECT * FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & nDiag & "|%'")
                    If Not rsTmp.EOF Then
                        ' TUTUP TAMBAH PEKERJAAN
                        If rsTmp("FLAG") <> "N" Then
                            text(44).BackColor = &HC0E0FF
                            text(46).BackColor = &HC0E0FF
                            woOnProgres = True
                        End If
                        
                        cmdStall.Tag = ""
                        Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN like '%|" & nDiag & "|%'")
                        Call objO.InsertData("INSERT INTO tempStall SELECT noStall, tanggal, JamMulai, JamSelesai, rate, [Group], Mech, Flag, Keterangan,adjust FROM STALL WITH (NOLOCK) " _
                                     & "WHERE KETERANGAN like '%|" & nDiag & "|%' AND NOT EXISTS " _
                                     & "(SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE KETERANGAN like '%|" & nDiag & "|%' )") 'AND Tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'
                                     
                        text(86) = Trim(rsTmp("JAMMULAI"))
                        
                        Select Case rsTmp("FLAG")
                            Case "I"
                                If rsTemp("NOINV") <> "" Then
                                    pesan = "Sudah diinvoice dengan No Invoice : " & Trim(rsTemp("NOINV")) & "  ,Tanggal : " & TglDDMMMMYYYY(rsTemp("INVDATE"))
                                    Call tutupForm
                                Else
                                    pesan = IIf(Nul(rsTemp("WASHING")) = 0, "Kendaraan sudah selesai dikerjakan", "Kendaraan dalam proses washing")
                                End If
                                MsgBox pesan
                            Case "S"
                                MsgBox "Kendaraan dalam proses pengerjaan"
                            Case "P"
                                MsgBox "Kendaraan dalam proses dihentikan sementara"
                            Case "H"
                                MsgBox "Kendaraan menunggu proses Final Inspection"
                            Case "Q"
                                MsgBox "Kendaraan dalam proses Final Inspection"
                            Case "G"
                                MsgBox "Kendaraan menunggu proses Washing"
                            Case "A"
                                MsgBox "Kendaraan dalam proses Washing"
                            Case "F"
                                MsgBox "Kendaraan sudah selesai dikerjakan"
                            Case "K"
                                MsgBox "Kendaraan sudah Closing"
                                Call tutupForm
                            Case ""
                                MsgBox "Kendaraan dalam proses menunggu perbaikan"
                        End Select
                    Else
                        cmdStall.Tag = "belum"
                        If chPDS.value = Checked Then
                            woOnProgres = True
                        End If
                    End If
                End If
                TblSimpan.Enabled = True
                cmdAddJob.Enabled = True
                If rsTemp("NOINV") <> "" Then
                    TblSimpan.Enabled = False
                    cmdAddJob.Enabled = False
                End If
            Else
                MsgBox "WO ini sudah di Closing "
                Exit Sub
            End If
        End If
    Else
        MsgBox "WO ini sudah divalidasi"
        TblSimpan.Enabled = False
    End If
End Sub

Private Sub simpanPrediag()
    'ISI FIELD
    Call ClearsSql
    
    newPre = False
    If Trim(text(22)) = "" Then
        text(22) = isiNoPRE
        newPre = True
    End If
    
    gsSql1(0) = "NOFAK"
    gsSql2(0) = text(78)   'Text(0)
    gsSql1(1) = "FS_KD_PLG"
    gsSql2(1) = text(1)
    gsSql1(2) = "FS_KD_WIL"
    gsSql2(2) = text(2)
    gsSql1(3) = "CUSTOMER"
    gsSql2(3) = text(3)
    gsSql1(4) = "CUST_ADRES"
    gsSql2(4) = text(4)
    gsSql1(5) = "KOTA"
    gsSql2(5) = text(5)
    gsSql1(6) = "TELPON"
    gsSql2(6) = text(6)
    gsSql1(7) = "FS_TH_MOBIL"
    gsSql2(7) = text(9)
    gsSql1(8) = "MDL"
    gsSql2(8) = text(10)
    gsSql1(9) = "MODEL"
    gsSql2(9) = text(11)
    gsSql1(10) = "FMDL"
    gsSql2(10) = text(12)
    gsSql1(11) = "COLOR"
    gsSql2(11) = text(13)
    gsSql1(12) = "NO_POLISI"
    gsSql2(12) = text(14)
    gsSql1(13) = "NO_CHASSIS"
    gsSql2(13) = text(15)
    gsSql1(14) = "DOTAM"
    gsSql2(14) = Replace(text(16), ">", "")
    gsSql1(15) = "NO_ENGINE"
    gsSql2(15) = text(76)
    gsSql1(16) = "FS_KD_DEALER"
    gsSql2(16) = Trim(gKdCbg)
    gsSql1(17) = "FS_NM_DEALER"
    gsSql2(17) = Trim(gNmCbg)
    gsSql1(18) = "FRONT"
    gsSql2(18) = gSA
    gsSql1(19) = "MECH"
    gsSql2(19) = text(50)
    
    gsSql1(99) = "[GROUP]"
    gsSql2(99) = Trim(text(106))
    
    gsSql1(20) = "KM"
    gsSql2(20) = CDbl(text(51))
    gsSql1(21) = "DATE"
    gsSql2(21) = Format(text(65), "mm/dd/yyyy")
    gsSql1(22) = "TIME"
    gsSql2(22) = text(66)
    gsSql1(23) = "PART"
    gsSql2(23) = CDbl(text(57))
    gsSql1(24) = "MTRL"
    gsSql2(24) = CDbl(text(58))
    gsSql1(25) = "OIL"
    gsSql2(25) = CDbl(text(59))
    gsSql1(26) = "SBLT"
    gsSql2(26) = CDbl(text(60))
    gsSql1(27) = "LABOR"
    gsSql2(27) = CDbl(text(56))
    gsSql1(28) = "BODY"
    gsSql2(28) = CDbl(text(77))
    gsSql1(29) = "TOTAL"
    gsSql2(29) = CDbl(text(61))
    gsSql1(30) = "REQ1"
    gsSql2(30) = text(23)
    gsSql1(31) = "REQ2"
    gsSql2(31) = text(24)
    gsSql1(32) = "REQ3"
    gsSql2(32) = text(25)
    gsSql1(33) = "REQ4"
    gsSql2(33) = text(26)
    gsSql1(34) = "REQ5"
    gsSql2(34) = text(27)
    For i = 1 To 7
        gsSql1(34 + i) = "REP" & i
        gsSql2(34 + i) = DGrid.TextMatrix(i, 1)
        gsSql1(41 + i) = "RTE" & i
        gsSql2(41 + i) = Replace(IIf(DGrid.TextMatrix(i, 2) = "", 0, DGrid.TextMatrix(i, 2)), ",", ".")
        
        ' Actual Rate
        Dim nActRate As Double
        Set rsTemp = objO.FetchData("SELECT * FROM actrate WITH (NOLOCK) " _
                                 & "WHERE FS_ITEM = '" & DGrid.TextMatrix(i, 1) & "' AND MDL = '" & Trim(text(10)) & "' AND FS_LOG_DEL = '' " _
                                 & "ORDER BY levelMech")
        If Not rsTemp.EOF Then
            gsSql1(48 + i) = "ACTRATE" & i
            gsSql2(48 + i) = Replace(CDbl(rsTemp("RATE")), ",", ".")
            'gSBE = IIf(rsTemp("SBE") = "Y", True, False)
        Else
            gsSql1(48 + i) = "ACTRATE" & i
            gsSql2(48 + i) = 0
            gSBE = False
        End If
        nActRate = nActRate + gsSql2(48 + i)
        
        ' GROUP JOB
        gsSql1(55 + i) = "KLPKERJA" & i
        gsSql2(55 + i) = DGrid.TextMatrix(i, 4)
        
        ' NILAI HARGA JOB
        gsSql1(62 + i) = "OFF" & i
        gsSql2(62 + i) = IIf(DGrid.TextMatrix(i, 5) <> "", DGrid.TextMatrix(i, 5), 0)
        
        ' MECH YG MENGERJAKAN PER JOB
        gsSql1(69 + i) = "MEKANIK" & i
        gsSql2(69 + i) = DGrid.TextMatrix(i, 6)
        
        Set rsHlp = objO.FetchData("SELECT [GROUP] FROM MEKANIK WITH (NOLOCK) WHERE MECH = '" & Trim(gsSql2(69 + i)) & "'")
        If Not rsHlp.EOF Then
            gsSql1(76 + i) = "GRP" & i
            gsSql2(76 + i) = rsHlp("GROUP")
        Else
            gsSql1(76 + i) = "GRP" & i
            gsSql2(76 + i) = ""
        End If
    Next
    gsSql1(84) = "NO_DIAG"
    gsSql2(84) = text(22)
    gsSql1(85) = "TPSVC"
    gsSql2(85) = IIf(opTpSvc(2).value, "W", IIf(opTpSvc(1).value, "P", ""))
    gsSql1(86) = "KODETIPE"
    gsSql2(86) = IIf(opTpSvc(2).value, text(21), IIf(opTpSvc(1).value, text(21), ""))
    gsSql1(87) = "NOINV"
    gsSql2(87) = text(67)
    gsSql1(88) = "STATUS"
    gsSql2(88) = Trim(text(103)) 'CboStatus
    gsSql1(89) = "PAID"
    gsSql2(89) = IIf(opByr(0).value, "CS", IIf(opByr(1).value, "CR", "T"))
    gsSql1(90) = "REMARK1"
    gsSql2(90) = Txt(54)
    gsSql1(91) = "SUGGEST"
    gsSql2(91) = Txt(55)
    gsSql1(92) = "FS_LOG_ADD"
    gsSql2(92) = gSA
    gsSql1(93) = "CATAT"
    gsSql2(93) = Replace(Txt(53), "'", "`") & " " & Replace(Txt(87), "'", "`")
    gsSql1(94) = "FN_RETURN"
    gsSql2(94) = IIf(ChkReturnOrder.value = Checked, 1, 0)
    gsSql1(95) = "THS"
    gsSql2(95) = IIf(chTHS.value = Checked, "1", "0")
    'gsSql1(82) = "SBE"
    'gsSql2(82) = IIf(gSBE, "SBE", "")
    
    gsSql1(96) = "ANCUST"
    gsSql2(96) = Replace(Trim(text(102)), "'", "`")
    gsSql1(97) = "MOU"
    gsSql2(97) = Trim(text(90))
    gsSql1(98) = "TELPANCUST"
    gsSql2(98) = Trim(text(105))
    gsSql1(100) = "[CODE]"
    gsSql2(100) = IIf(chEM.value = Checked, "1", "0")
    gsSql1(101) = "KD_K"
    gsSql2(101) = IIf(chSBNP.value = Checked, "1", "0")
    
    If Not woOnProgres Then
        If chPDS = Checked Then
            gsSql1(102) = "RPR_COND"
            gsSql2(102) = "S"
            gsSql1(103) = "DELIV"
            gsSql2(103) = "OKE"
            gsSql1(104) = "PDS"
            gsSql2(104) = "1"
        Else
            gsSql1(102) = "RPR_COND"
            gsSql2(102) = IIf(onlySublet, "S", "")
            gsSql1(103) = "DELIV"
            gsSql2(103) = IIf(onlySublet, "OKE", "")
            gsSql1(104) = "PDS"
            gsSql2(104) = "0"
        End If
    Else
        If chPDS = Checked Then
            gsSql1(102) = "PDS"
            gsSql2(102) = "1"
        Else
            gsSql1(102) = "PDS"
            gsSql2(102) = "0"
        End If
    End If
    
    ' KURANGI JAM ALOKASI LAMA
    Dim rsLama As ADODB.Recordset
    Dim rsAct As ADODB.Recordset
    Dim hlp As String
    Dim nTotAct As Integer
    
    If cmdPilih.Tag = "No. WO" Then
        Set rsLama = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_DIAG = '" & Trim(text(22)) & "'")
        If Not rsLama.EOF Then
            nTotAct = 0
            For i = 1 To 7
                hlp = "REP" & i
                If rsLama(hlp) <> "" Then
                    nTotAct = nTotAct + (objO.FetchData_Double("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'"))
                End If
            Next
            Call createMekanik
            'Call decresRestTime(rsLama("MECH"), CDbl(nTotAct), 1, Format(rsLama("DATE"), "MM/DD/YYYY"))
        End If
    End If
    
    Set rsTemp = objO.FetchData("SELECT * FROM DIAGdta WITH (NOLOCK) WHERE no_chassis = '" & Trim(text(15)) & "' AND no_DIAG = '" & Trim(text(22)) & "'")
    If Not rsTemp.EOF Then
        sSQL = GensSql2("DIAGDTA", True) + "WHERE no_chassis = '" & text(15) & "' AND no_DIAG = '" & text(22) & "'"
    Else
        sSQL = GensSql2("DIAGDTA", False)
    End If
    Call objHT.UpdateData_Bengkel(sSQL)
    
'    UPDATE JAM TERPAKAI MEKANIK
    nTotAct = 0
    For i = 1 To 7
        If DGrid.TextMatrix(i, 1) <> "" Then
            Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
            If Not rsAct.EOF Then
               nTotAct = nTotAct + CDbl(rsAct("RATE"))
            End If
        End If
    Next
    Call countRestTime(Trim(text(50)), CDbl(nTotAct), 1, Format(text(65), "MM/DD/YYYY"))
    
    ' UPDATE CLOCK TUNGGU DILAYANI & CLOCK SA
    Set rsTemp = objO.FetchData("SELECT jam, jamPanggil, DATEDIFF(MINUTE, jam, jamPanggil) as waktu, DATEDIFF(MINUTE, jamPanggil, GETDATE()) as waktuSA FROM T_ANTRIAN WITH (NOLOCK) WHERE loket = '" & gLoket & "' AND noAntrian = '" & gNoAntri & "' AND statusCust = '" & gTpAntri & "'" _
                             & "AND jam >= '" & Format(isnow, "mm/dd/yyyy 00:00:00") & "' AND jam < '" & Format(isnow, "mm/dd/yyyy 23:59:59") & "'")
    If Not rsTemp.EOF Then
        Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET FS_RC = '', TIME_AT = '" & hitungJam(rsTemp("WAKTU")) & "', TIME_SA = '" & hitungJam(rsTemp("WAKTUSA")) & "' where NO_DIAG ='" & text(22) & "'")
        Call objHT.UpdateData_Bengkel("UPDATE T_ANTRIAN SET NoPol = '" & text(14) & "', NO_WO = '" & Trim(text(22)) & "' WHERE noAntrian = '" & gNoAntri & "' AND StatusCust = '" & gTpAntri & "' AND JAM BETWEEN '" & Format(isnow, "mm/dd/yyyy 00:00:00") & "' AND '" & Format(isnow, "mm/dd/yyyy 23:59:59") & "'")
    End If
    
    ' UPDATE NILAI REGULAR CHECK
    Call objHT.UpdateData_Bengkel("UPDATE JOBDTA SET FS_RC ='', SBE = '' where NO_WO ='" & Trim(text(64)) & "' ")
    Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC <>''")
    If Not rsTemp.BOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_DIAG ='" & text(22) & "' " _
                 & " AND REP1+ REP2+ REP3+ REP4+ REP5+ REP6+ REP7 like '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%'")
            rsTemp.MoveNext
        Loop
    End If
    
    ' UPDATE NILAI REGULAR CHECK
    Dim rsTemp2 As ADODB.Recordset
    Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC <>''")
    If Not rsTemp.BOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            Set rsTemp2 = objO.FetchData("SELECT * FROM DIAGADD2 WITH (NOLOCK) WHERE NO_DIAG = '" & text(22) & "' AND REP LIKE '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%'")
            If Not rsTemp2.EOF Then
                Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_DIAG ='" & text(22) & "' ")
            End If
            rsTemp.MoveNext
        Loop
    End If
    
End Sub

Private Sub simpanPreDtl()

    'ISI FIELD
    Call ClearsSql
            
    gsSql1(0) = "NO_DIAG"
    gsSql2(0) = text(22)
    gsSql1(1) = "NO_CHASSIS"
    gsSql2(1) = text(15)
    gsSql1(2) = "NO_WO"
    gsSql2(2) = Trim(text(64))
    For i = 0 To 5
        Select Case i
            Case 5
                gsSql1(3 + i) = "COND" & (i + 1)
                gsSql2(3 + i) = IIf(text(28 + i) = "LAINNYA", "L| " & Trim(bKondisi(i)), text(28 + i))
            Case Else
                gsSql1(3 + i) = "COND" & (i + 1)
                gsSql2(3 + i) = IIf(text(28 + i) = "LAINNYA", "L| " & Trim(bKondisi(i)), text(28 + i))
                
                gsSql1(9 + i) = "WAK" & (i + 1)
                gsSql2(9 + i) = text(34 + i)
                
                gsSql1(14 + i) = "WAE" & (i + 1)
                gsSql2(14 + i) = text(39 + i)
        End Select
    Next
    gsSql1(19) = "BBM"
    gsSql2(19) = IIf(opBBM(0).value, "0", IIf(opBBM(1).value, "1", IIf(opBBM(2).value, "2", IIf(opBBM(3).value, "3", "4"))))
    gsSql1(20) = "GANTIPART"
    gsSql2(20) = IIf(opGanti(0).value, "0", "1")
    gsSql1(21) = "TUNGGUSVC"
    gsSql2(21) = IIf(opSvc(0).value, "0", "1")
    gsSql1(22) = "PEMBAYARAN"
    gsSql2(22) = IIf(opByr(0).value, "CS", IIf(opByr(1).value, "CR", "T"))
    gsSql1(23) = "HEAVYUSE"
    gsSql2(23) = IIf(chUse.value = Checked, 1, 0)
    
    Set rsHlp = objO.FetchData("SELECT * FROM tempFU WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "'") ' AND no_wo = '" & Text(22) & "'
    If Not rsHlp.EOF Then
        gsSql1(24) = "TGLFU"
        gsSql2(24) = Format(rsHlp("TGLFU"), "mm/dd/yyyy")
        gsSql1(25) = "JAMFU"
        gsSql2(25) = rsHlp("JAMFU")
        gsSql1(26) = "TELPFU"
        gsSql2(26) = rsHlp("TELPFU")
        gsSql1(27) = "KETFU"
        gsSql2(27) = rsHlp("KETFU")
    Else
        gsSql1(24) = "TGLFU"
        gsSql2(24) = ""
        gsSql1(25) = "JAMFU"
        gsSql2(25) = ""
        gsSql1(26) = "TELPFU"
        gsSql2(26) = ""
        gsSql1(27) = "KETFU"
        gsSql2(27) = ""
    End If
    rsHlp.Close
    
    gsSql1(28) = "VERIFIKATOR"
    gsSql2(28) = gVerifikator
    gsSql1(29) = "DATESERAH"
    gsSql2(29) = Format(text(88), "MM/DD/YYYY")
    gsSql1(30) = "TIMESERAH"
    gsSql2(30) = Trim(text(89))
    gsSql1(31) = "DATANG"
    gsSql2(31) = Trim(text(107))
    gsSql1(32) = "DTLDATANG"
    gsSql2(32) = Trim(text(52))
    gsSql1(33) = "KARAKTER"
    gsSql2(33) = Trim(cboPelanggan)
    gsSql1(34) = "DTLKARAKTER"
    gsSql2(34) = Trim(text(75))
    gsSql1(35) = "WASHING"
    gsSql2(35) = IIf(opWash(0), 1, 0)
    
    For i = 1 To 5
        gsSql1(35 + i) = "CMPL" & i
        'gsSql2(35 + i) = text(96 + i)
        gsSql2(35 + i) = Text1(-1 + i)
    Next
    
    Set rsTemp = objO.FetchData("SELECT * FROM diagdtl WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "' AND no_diag = '" & text(22) & "'")
    If Not rsTemp.EOF Then
        sSQL = GensSql2("DIAGDTL", True) + "WHERE no_chassis = '" & text(15) & "' AND no_DIAG = '" & text(22) & "'"
    Else
        sSQL = GensSql2("DIAGDTL", False)
    End If
    Call objHT.UpdateData_Bengkel(sSQL)
    Call objHT.DeleteData_Bengkel("DELETE tempFU WHERE no_chassis = '" & text(15) & "' AND no_WO = '" & text(22) & "'")
End Sub

Private Sub simpanPreKoord()
    'ISI FIELD
    Call ClearsSql
            
    gsSql1(0) = "NO_DIAG"
    gsSql2(0) = text(22)
    gsSql1(1) = "NO_CHASSIS"
    gsSql2(1) = text(15)
    gsSql1(2) = "NO_WO"
    gsSql2(2) = Trim(text(64))
    
    Call objHT.DeleteData_Bengkel("DELETE diagKoord WHERE no_chassis = '" & text(15) & "' AND no_diag = '" & text(22) & "'")
    
    For i = 0 To nk - 1
        gsSql1(3) = "KOORDX"
        gsSql2(3) = KorX(i)
        gsSql1(4) = "KOORDY"
        gsSql2(4) = KorY(i)
        
'        Set rsTemp = objO.FetchData("SELECT * FROM DIAGKoord WHERE no_chassis = '" & text(15) & "' AND no_DIAG = '" & text(22) & "' AND KOORDX = '" & KorX(i) & "' AND KOORDY = '" & KorY(i) & "'")
'        If rsTemp.EOF Then
            sSQL = GensSql2("DIAGKOORD", False)
'        End If
        Call objHT.UpdateData_Bengkel(sSQL)
    Next
    
    ' SIMPAN STALL
    Dim NoFak As String
    Set rsTemp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE keterangan like '%" & Trim(text(14)) & " |%' ") 'AND tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "'")
    If Not rsTemp.EOF Then
        If cmdPilih.Tag = "No. Rangka" Then
            Call objHT.UpdateData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & text(22) & "|%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'")
            Call objHT.InsertData_Bengkel("INSERT INTO STALL SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [Group], '" & Trim(text(50)) & "', 'N', '" & text(14) & " |" & text(22) & "| " & text(3) & "' AS Keterangan,adjust FROM tempStall WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM STALL WITH (NOLOCK) WHERE keterangan like '%|" & Trim(text(22)) & "|%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "') AND keterangan like '%| " & text(3) & "%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'")
            Call objHT.UpdateData_Bengkel("DELETE TEMPSTALL WHERE KETERANGAN LIKE '%| " & text(3) & "%' AND tanggal = '" & Format(text(65), "mm/dd/yyyy") & "'")
        End If
    End If
End Sub

Private Sub simpanPrePart()
Dim nPart As String

    'ISI FIELD
    Call ClearsSql
            
    gsSql1(0) = "NOFAK"
    gsSql2(0) = text(78)
    gsSql1(1) = "NO_CHASSIS"
    gsSql2(1) = text(15)
    gsSql1(2) = "NO_WO"
    gsSql2(2) = Trim(text(64))
    
    ' //???// NOTE KE SPAREPART //???//
    ' HAPUS PART YG TIDAK ADA DI BOOK_PART
    Set rsTemp = objO.FetchData("SELECT * FROM DIAGPART WITH (NOLOCK) WHERE NO_DIAG = '" & Trim(text(22)) & "' ORDER BY kode_Item")
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If rsGridPart.State <> adStateClosed Then
                If Not rsGridPart.EOF Then
                    rsGridPart.MoveFirst
                    rsGridPart.Find "KODEITEM = '" & Trim(rsTemp("KODE_ITEM")) & "'", , adSearchForward, 1
                    If Not rsGridPart.EOF Then
                        Call objHT.DeleteData_Bengkel("DELETE DIAGPART WHERE NO_DIAG = '" & Trim(text(22)) & "' AND KODE_ITEM = '" & Trim(rsTemp("KODE_ITEM")) & "'")
                    End If
                End If
            Else
                Call objHT.DeleteData_Bengkel("DELETE DIAGPART WHERE NO_DIAG = '" & Trim(text(22)) & "'")
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If Not rsGridPart.State = adStateClosed Then
        If Not rsGridPart.EOF Then
            rsGridPart.MoveFirst
            i = 0
            nPart = ""
            Do While Not rsGridPart.EOF
                
                gsSql1(3) = "KODE_ITEM"
                gsSql2(3) = rsGridPart("KODEITEM")
                gsSql1(4) = "NAMA_ITEM"
                gsSql2(4) = rsGridPart("NAMAITEM")
                gsSql1(5) = "QTY"
                gsSql2(5) = rsGridPart("QTY")
                gsSql1(6) = "GANTI"
                gsSql2(6) = IIf(rsGridPart("GANTI"), "Y", "N")
                gsSql1(7) = "BAWA"
                gsSql2(7) = IIf(rsGridPart("BAWA"), "Y", "N")
                gsSql1(8) = "PARTBAWA"
                If rsGridPart("BAWA") Then
                    gsSql2(8) = text(69 + i)
                    i = i + 1
                Else
                    gsSql2(8) = ""
                End If
                gsSql1(9) = "TGLPASANG"
                gsSql2(9) = Format(text(65), "MM/DD/YYYY")
                
                ' AMBIL QTY SUPPLY DARI SPAREPART
                    Set rsTemp = objO.FetchDataACC("SELECT INTRN.KODEITEM, SUM(QTY) AS JML FROM INTRN WITH (NOLOCK) INNER JOIN INMST WITH (NOLOCK) ON INTRN.KODEITEM=INMST.KODEITEM " _
                                                & "WHERE SBRBUKTI = '" & Trim(text(0)) & "' AND INTRN.KODEITEM = '" & rsGridPart("KODEITEM") & "' AND INMST.KODEGDG = '04' GROUP BY INTRN.KODEITEM HAVING SUM(QTY) <> 0")
                    If Not rsTemp.EOF Then
                        gsSql1(10) = "QTYSUPPLY"
                        gsSql2(10) = CDbl(rsTemp("JML"))
                    Else
                        gsSql1(10) = "QTYSUPPLY"
                        gsSql2(10) = 0
                    End If
                Dim nDis As Integer
                nDis = 0
                If opTpSvc(2) And text(21) <> "" Then
                    ' UTK SSC PART TIDAK MINTA KE SPAREPART
                Else
                    If opTpSvc(1) Then
                        nDis = objO.FetchData_String("SELECT DISCPART FROM HO_PROGRAM WITH (NOLOCK) WHERE ID_PROGRAM = '" & text(21) & "'")
                    End If
                    If rsGridPart("GANTI") = "Y" And rsGridPart("BAWA") = "N" Then
                        nPart = nPart + (rsGridPart("KODEITEM") & "^" & rsGridPart("NAMAITEM") & "^" & rsGridPart("QTY") & "^" & nDis & "|")
                    End If
                End If
                
                Set rsTemp = objO.FetchData("SELECT * FROM DIAGPart WITH (NOLOCK) WHERE no_chassis = '" & text(15) & "' AND no_DIAG = '" & text(22) & "' AND KODE_ITEM = '" & rsGridPart("KODEITEM") & "' AND ADDJOB = 0")
                gsSql1(11) = "no_DIAG"
                gsSql2(11) = Trim(text(22))
                If Not rsTemp.EOF Then
                    sSQL = GensSql2("DIAGPART", True) + "WHERE no_chassis = '" & text(15) & "' AND no_DIAG = '" & text(22) & "' AND KODE_ITEM = '" & rsGridPart("KODEITEM") & "' AND ADDJOB = 0"
                Else
                    sSQL = GensSql2("DIAGPART", False)
                End If
                Call objHT.UpdateData_Bengkel(sSQL)
                rsGridPart.MoveNext
            Loop
        End If
    End If
    
    ' INPUT SUBLET
    
    Set rsTemp = objO.FetchData("SELECT * FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE no_Chassis = '" & Trim(text(15)) & "' AND NO_WO = '" & IIf(Trim(text(64)) = "", "", Trim(text(22))) & "'")
    If Not rsTemp.State Then
        gPO = IIf(gPO <> "", gPO, isiNoPoSublet)
        Call objHT.DeleteData_Bengkel("DELETE DIAGSUBLET WHERE no_DIAG = '" & Trim(text(22)) & "' AND no_Chassis = '" & Trim(text(15)) & "' AND NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
        Call objHT.InsertData_Bengkel("INSERT INTO DIAGSUBLET SELECT DISTINCT '" & gPO & "', '" & text(22) & "', '" & Trim(text(64)) & "', '" & text(67) & "', no_chassis, nofak, kodeItem, NamaBarang, kodeSatuan, Qty, Harga, nonota, keterangan,FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, '', '' " _
                     & "FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM DIAGSUBLET WITH (NOLOCK) WHERE no_DIAG = '" & Trim(text(22)) & "' AND no_Chassis = '" & Trim(text(15)) & "') AND " _
                     & "no_Chassis = '" & Trim(text(15)) & "' AND NO_WO = '' AND NO_POLISI = '" & Trim(text(14)) & "'")
        Call objHT.DeleteData_Bengkel("DELETE TEMPJOBSUBLET WHERE no_wo = '" & Trim(text(22)) & "' AND no_Chassis = '" & text(15) & "'  AND NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
    Else
        Call objHT.DeleteData_Bengkel("DELETE DIAGSUBLET WHERE no_DIAG = '" & Trim(text(22)) & "' AND no_Chassis = '" & Trim(text(15)) & "' AND NOFAK = '" & Trim(text(78)) & "' AND NO_POLISI = '" & Trim(text(14)) & "'")
    End If
End Sub

Private Sub cetakFormPrediag()
    Dim rpt As Object
    Select Case gPrintCetak
        Case "2"
            Set rpt = New RptOrder4            ' KGW + PEMUDA ' 2190
        Case "1"
            Set rpt = New RptOrder5            ' MLATI + ALL  ' 2180
        Case "3"
            Set rpt = New RptOrder6            ' SLAMET  ' LX 300+II
        Case "4"
            Set rpt = New RptOrder7            ' BLANKO UTK KERTAS KOSONG
            rpt.txtDealer.SetText support.kopOrder(0)
    End Select
    Dim rsRepo As ADODB.Recordset
    
    If MsgBox("Mencetak Order ?", vbYesNo, "Konfirmasi") = vbYes Then
        gNoOrder = text(22) & "/" & Trim(text(103))
        gKetSA = gSA
        
        NamaSAOrder = objO.FetchData_String("Select FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS ='" & Trim(gKetSA) & "'")
        
        gCust = text(3)
        gAlmt = text(4)
        gKota = text(5)
        gTelpon = text(6)
        gCatatan = Txt(53) & " " & Txt(87)
        gTahun = text(9)
        gNoPolisi = text(14)
        gModel = text(11)
        gWarna = text(13)
        gTglMasuk = text(65)
        gJamMasuk = text(66)
        gDealer = text(80)
        gDD = Replace(text(79), ">", "")
        gMDLOrder = text(10)
        gFRM = text(15)
        gEG = text(76)
        gKeluhan(1) = text(23)
        gKeluhan(2) = text(24)
        gKeluhan(3) = text(25)
        gKeluhan(4) = text(26)
        gKeluhan(5) = text(27)
        gOrder(1) = DGrid.TextMatrix(1, 1)
        gOrder(2) = DGrid.TextMatrix(2, 1)
        gOrder(3) = DGrid.TextMatrix(3, 1)
        gOrder(4) = DGrid.TextMatrix(4, 1)
        gOrder(5) = DGrid.TextMatrix(5, 1)
        gOrder(6) = DGrid.TextMatrix(6, 1)
        gOrder(7) = DGrid.TextMatrix(7, 1)
        
        Set rsHlp = objO.FetchData("SELECT * FROM DIAGSUBLET WITH (NOLOCK) WHERE NO_DIAG = '" & text(22) & "' AND NO_CHASSIS = '" & text(15) & "'")
        If Not rsHlp.EOF Then
            i = 1
            Do While Not rsHlp.EOF
                If i < 10 Then
                    gSublet(i) = rsHlp("NAMABARANG")
                    i = i + 1
                End If
                rsHlp.MoveNext
            Loop
        End If
        
        gKM = Format(text(51), "#,#0")
        gPKhusus = Txt(54)
        gSAdvisor = NamaSAOrder
        gPSelesai = text(63)
        gKodePlg = text(1)
        
        Dim sHlp As Object
        Dim sHlp2 As Object
        Dim sHlp3 As Object
        Dim sHlp4 As Object
        Dim crxSubReportHistory As CRAXDRT.Report
        Dim crxSubReportBooking As CRAXDRT.Report
        Dim crxSubReportBooking2 As CRAXDRT.Report
        Dim crxSubReportBookingFU As CRAXDRT.Report
        If gUkuranOrder = False Then
        
            Dim rRs As ADODB.Recordset
            Set rRs = objO.FetchData("SELECT A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.NO_POLISI, A.NO_ENGINE, A.NO_CHASSIS, A.MDL, A.MODEL, A.FMDL, A.COLOR, " _
                                  & "D.FS_KD_DEALER, D.FS_NM_DEALER, D.DDATE, A.DOTAM, A.CODE, A.NPWP, A.PPN, A.DISC, '0', '', A.FRONT, A.[GROUP], A.MECH, " _
                                  & "A.KM, A.DATE, A.TIME, '', '', '', '', '', A.PART, A.MTRL, A.OIL, A.MET, A.SBLT, A.LABOR, A.BODY, A.TOTAL, " _
                                  & "A.JOB, A.REQ1, A.REQ2, A.REQ3, A.REQ4, A.REQ5, '', '', A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, A.REP8, A.SLIP, " _
                                  & "A.RTE1, A.RTE2, A.RTE3, A.RTE4, A.RTE5, A.RTE6, A.RTE7, A.RTE8, A.INVDATE, A.NO_diag, A.TPSVC, A.NOINV, A.STATUS, A.PAID, A.NOFAK, " _
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
                                  & "FROM DIAGDTA AS A WITH (NOLOCK) INNER JOIN DIAGDTL AS B WITH (NOLOCK) ON A.NO_DIAG = B.NO_DIAG LEFT OUTER JOIN CUSTDAT AS D WITH (NOLOCK) ON A.NO_CHASSIS = D.NO_CHASSIS INNER JOIN " _
                                  & "T_PETUGAS AS E WITH (NOLOCK) ON A.FRONT = E.FS_KD_PETUGAS WHERE A.NO_DIAG = '" & Trim(text(22)) & "'")
            rpt.Database.SetDataSource rRs
            gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
            
            sSQL = "SELECT * FROM historyUnitFunc('" & Trim(gFRM) & "')"
            Set sHlp = objO.FetchData(sSQL)
            Set crxSubReportHistory = rpt.Subreport2.OpenSubreport
            crxSubReportHistory.Database.SetDataSource sHlp

            sSQL = "select DATE_BOOK, TIME_D, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(78)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
            Set sHlp2 = objO.FetchData(sSQL)
            Set crxSubReportBooking = rpt.Subreport1.OpenSubreport
            crxSubReportBooking.Database.SetDataSource sHlp2

            sSQL = "select DATE, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(78)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
            Set sHlp4 = objO.FetchData(sSQL)
            Set crxSubReportBooking2 = rpt.Subreport4.OpenSubreport
            crxSubReportBooking2.Database.SetDataSource sHlp4

            sSQL = "select WAKTU FROM M_FUBOOK WITH (NOLOCK) where NOFAK = '" & Trim(text(78)) & "' "
            Set sHlp3 = objO.FetchData(sSQL)
            Set crxSubReportBookingFU = rpt.Subreport3.OpenSubreport
            crxSubReportBookingFU.Database.SetDataSource sHlp3

            Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(text(22)) & "|%'") 'TANGGAL = '" & Format(Trim(text(65)), "MM/DD/YYYY") & "' AND
            If Not rsHlp.EOF Then
                rpt.txtStall.SetText Trim(rsHlp("NOSTALL"))
            End If

            rpt.mechPlan.SetText ""

            If bodyClear Then
                rpt.Text93.TextColor = &HFF&
                rpt.Text94.TextColor = &HFF&
                rpt.Text95.TextColor = &HFF&
                rpt.Text96.TextColor = &HFF&
                rpt.Text97.TextColor = &HFF&
            End If

            Dim n As Integer
            Set rsHlp = objO.FetchData("SELECT * FROM DIAGKOORD WITH (NOLOCK) WHERE NO_DIAG = '" & Trim(text(22)) & "'")
            If Not rsHlp.EOF Then
                n = 1
                Do While Not rsHlp.EOF
                    Call pasangTanda(rpt, n, CDbl(rsHlp("KOORDY")), CDbl(rsHlp("KOORDX")))
                    n = n + 1
                    rsHlp.MoveNext
                Loop
            End If

            Set rsHlp = objO.FetchData("SELECT BBM FROM DIAGDTL WITH (NOLOCK) WHERE NO_DIAG = '" & Trim(text(22)) & "'")
            If Not rsHlp.EOF Then
                Select Case rsHlp("BBM")
                    Case 0
                        ' FULL
                        rpt.Text90.TextColor = &H0&
                    Case 1
                        ' 3/4
                        rpt.Text62.TextColor = &H0&
                    Case 2
                        ' 1/2
                        rpt.Text64.TextColor = &H0&
                    Case 3
                        ' 1/4
                        rpt.Text88.TextColor = &H0&
                    Case 4
                        ' EMPTY
                        rpt.Text89.TextColor = &H0&
                End Select
            End If
            With FrmctkOrder.CRViewer1
                .ReportSource = rpt
                .ViewReport
                .Zoom 100
                .Visible = True
            End With
            FrmctkOrder.cmdTutup.Visible = True
            Screen.MousePointer = vbDefault
            FrmctkOrder.WindowState = vbMaximized
            FrmctkOrder.Show vbModal
        End If
    End If
End Sub

Private Function isOverdueAR() As Boolean
    Dim nOAR As ADODB.Recordset
    Set nOAR = objO.FetchOverdueAR(Trim(Replace(text(1).text, ">", "")))
    If nOAR.RecordCount > 0 Then
        Set nOAR = objO.FetchOverdueAR_Data(Trim(Replace(text(1).text, ">", "")))
        If nOAR.RecordCount > 0 Then
            isOverdueAR = True
            FrAR.Visible = True
            TimerAR.Enabled = True
        Else
            isOverdueAR = False
            FrAR.Visible = False
            TimerAR.Enabled = False
        End If
    Else
        isOverdueAR = False
        FrAR.Visible = False
        TimerAR.Enabled = False
    End If
End Function

Private Function isCurrentAR() As Boolean
    isCurrentAR = False
    If Trim(text(1).text) = "" Then Exit Function
    Dim aBatasKr As Double
    Dim aCurrent As Double
    aBatasKr = objO.Fetch_BatasKredit_Ar(Trim(Replace(text(1).text, ">", "")))
    isCurrentAR = False
    txtRp(201) = aBatasKr
    txtRp(1) = aBatasKr
    aCurrent = objO.Fetch_Current_Ar(Trim(text(1).text))
    txtRp(200) = aCurrent
    txtRp(0) = aCurrent
    If aCurrent <> 0 And aBatasKr <> 0 Then
        nPersenAR = Round((aCurrent / aBatasKr) * 100, 0)
    Else
        nPersenAR = 0
    End If
    nAR.Caption = nPersenAR & "%"
    nAR2.Caption = nPersenAR & "%"
    
    If aBatasKr = 0 Then GoTo TanpaBatasAR
    If nPersenAR >= 100 Then
        isCurrentAR = True
        FrAR.Visible = True
        TimerAR.Enabled = True
    ElseIf nPersenAR >= 80 Then
        isCurrentAR = True
        If objO.FetchIsValid_AR_Kabeng(Trim(text(1).text)) = True Then GoTo TanpaBatasAR
        FrAR.Visible = True
        TimerAR.Enabled = True
    ElseIf nPersenAR < 80 Then
        isCurrentAR = False
TanpaBatasAR:
        FrAR.Visible = False
        TimerAR.Enabled = False
    End If
End Function

Private Sub isCurrentAR_Info()
    If Trim(text(1).text) = "" Then Exit Sub
    Dim aBatasKr As Double
    Dim aCurrent As Double
    aBatasKr = objO.Fetch_BatasKredit_Ar(Trim(Replace(text(1).text, ">", "")))
    txtRp(201) = aBatasKr
    txtRp(1) = aBatasKr
    aCurrent = objO.Fetch_Current_Ar(Trim(text(1).text))
    txtRp(200) = aCurrent
    txtRp(0) = aCurrent
    If aCurrent <> 0 And aBatasKr <> 0 Then
        nPersenAR = Round((aCurrent / aBatasKr) * 100, 0)
    Else
        nPersenAR = 0
    End If
    nAR.Caption = nPersenAR & "%"
    nAR2.Caption = nPersenAR & "%"
End Sub

Private Sub TimerAR_Timer()
    FrAR.BackColor = RGB(Int(Rnd() * 256), Int(Rnd() * 256), Int(Rnd() * 256))
End Sub

Private Sub XPButton1_Click()
    FrAR.Visible = False
    TimerAR.Enabled = False
End Sub

Public Function GetDelIndex(ModilID As String) As String
    Dim acak As Double
    Randomize Timer
    acak = Int(Rnd * 99999)
    GetDelIndex = ModilID & acak
End Function


