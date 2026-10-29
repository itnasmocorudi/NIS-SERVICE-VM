VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBtlOrd 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pembatalan Order"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13305
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   12600
      TabIndex        =   196
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.PictureBox frmBtlOrder 
      BackColor       =   &H00FFFFFF&
      Height          =   2275
      Left            =   3360
      ScaleHeight     =   2220
      ScaleWidth      =   4590
      TabIndex        =   191
      Top             =   3840
      Visible         =   0   'False
      Width           =   4645
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1095
         Index           =   0
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   4335
      End
      Begin VB.Line Line5 
         BorderWidth     =   3
         X1              =   4560
         X2              =   0
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   5040
         X2              =   480
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   2280
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2280
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   120
         Picture         =   "frmBtlOrd.frx":0000
         Top             =   120
         Width           =   4380
      End
      Begin VB.Image cmdOk 
         Height          =   345
         Left            =   840
         Picture         =   "frmBtlOrd.frx":0DB2
         Top             =   1680
         Width           =   1515
      End
      Begin VB.Image cmdCancel 
         Height          =   345
         Left            =   2400
         Picture         =   "frmBtlOrd.frx":3C7A
         Top             =   1680
         Width           =   1470
      End
   End
   Begin VB.CheckBox ChkReturnOrder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10440
      TabIndex        =   182
      Top             =   120
      Width           =   1800
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
      Height          =   2490
      Left            =   8640
      TabIndex        =   171
      Top             =   480
      Width           =   3615
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
         Left            =   -120
         Picture         =   "frmBtlOrd.frx":6D93
         Style           =   1  'Graphical
         TabIndex        =   220
         Top             =   1800
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
         Left            =   1680
         Picture         =   "frmBtlOrd.frx":7550
         Style           =   1  'Graphical
         TabIndex        =   221
         Top             =   1800
         Width           =   1935
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   1080
         TabIndex        =   172
         Top             =   555
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
         Index           =   20
         Left            =   1080
         TabIndex        =   173
         Top             =   1440
         Width           =   1245
         _ExtentX        =   2196
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
         Index           =   19
         Left            =   1080
         TabIndex        =   174
         Top             =   1140
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
         Index           =   18
         Left            =   1080
         TabIndex        =   175
         Top             =   840
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
      Begin VB.Label Label61 
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
         Left            =   2160
         TabIndex        =   223
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label64 
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
         TabIndex        =   222
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   240
         Picture         =   "frmBtlOrd.frx":7D0D
         Top             =   0
         Width           =   3375
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
         Left            =   420
         TabIndex        =   181
         Top             =   1155
         Width           =   585
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
         Left            =   390
         TabIndex        =   180
         Top             =   1485
         Width           =   615
      End
      Begin VB.Label Label7 
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
         Left            =   435
         TabIndex        =   179
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         TabIndex        =   178
         Top             =   855
         Width           =   1170
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
         Left            =   420
         TabIndex        =   177
         Top             =   885
         Width           =   585
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2400
         TabIndex        =   176
         Top             =   570
         Width           =   1170
      End
   End
   Begin VB.CheckBox chTHS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unit THS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9000
      TabIndex        =   170
      Top             =   120
      Width           =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   8640
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
      Height          =   2490
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   8085
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   810
         TabIndex        =   3
         Top             =   750
         Width           =   3405
         _ExtentX        =   6006
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
         Index           =   4
         Left            =   810
         TabIndex        =   4
         Top             =   1035
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
         Index           =   5
         Left            =   810
         TabIndex        =   5
         Top             =   1350
         Width           =   3405
         _ExtentX        =   6006
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
         Index           =   6
         Left            =   810
         TabIndex        =   6
         Top             =   1650
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
         Left            =   5070
         TabIndex        =   7
         Top             =   435
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
         Index           =   10
         Left            =   5070
         TabIndex        =   8
         Top             =   750
         Width           =   2055
         _ExtentX        =   3625
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
         Index           =   13
         Left            =   5070
         TabIndex        =   9
         Top             =   1350
         Width           =   2895
         _ExtentX        =   5106
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
         Index           =   12
         Left            =   5070
         TabIndex        =   10
         Top             =   1050
         Width           =   2895
         _ExtentX        =   5106
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
         Left            =   7170
         TabIndex        =   11
         Top             =   750
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
         Index           =   2
         Left            =   2520
         TabIndex        =   12
         Top             =   450
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
         Index           =   1
         Left            =   810
         TabIndex        =   13
         Top             =   450
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
         Index           =   14
         Left            =   5070
         TabIndex        =   14
         Top             =   1650
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   "CC-9999-CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   5070
         TabIndex        =   15
         Top             =   1935
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   6750
         TabIndex        =   16
         Top             =   435
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
         Index           =   7
         Left            =   810
         TabIndex        =   17
         Top             =   1935
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
         Index           =   8
         Left            =   810
         TabIndex        =   18
         Top             =   2235
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
         Index           =   76
         Left            =   5070
         TabIndex        =   19
         Top             =   2235
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   79
         Left            =   1095
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
      Begin VB.Image Image4 
         Height          =   390
         Left            =   0
         Picture         =   "frmBtlOrd.frx":880D
         Top             =   0
         Width           =   9270
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
         Left            =   255
         TabIndex        =   39
         Top             =   795
         Width           =   405
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
         Left            =   165
         TabIndex        =   38
         Top             =   1095
         Width           =   495
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
         Left            =   330
         TabIndex        =   37
         Top             =   1395
         Width           =   330
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
         Left            =   285
         TabIndex        =   36
         Top             =   1695
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
         Left            =   4275
         TabIndex        =   35
         Top             =   465
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
         Left            =   4500
         TabIndex        =   34
         Top             =   765
         Width           =   420
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
         Left            =   4440
         TabIndex        =   33
         Top             =   1365
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
         Height          =   195
         Left            =   15
         TabIndex        =   32
         Top             =   480
         Width           =   660
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
         Left            =   4275
         TabIndex        =   31
         Top             =   1695
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
         Left            =   4410
         TabIndex        =   30
         Top             =   1980
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
         Left            =   6120
         TabIndex        =   29
         Top             =   480
         Width           =   585
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
         Left            =   345
         TabIndex        =   28
         Top             =   1980
         Width           =   315
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
         Left            =   225
         TabIndex        =   27
         Top             =   2280
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
         Left            =   4380
         TabIndex        =   26
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label SgnRekom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rekom"
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
         Left            =   6600
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   1290
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
         TabIndex        =   24
         Top             =   3165
         Width           =   450
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
         TabIndex        =   23
         Top             =   3015
         Width           =   585
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   480
      TabIndex        =   40
      Top             =   3000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12577002
      TabCaption(0)   =   "."
      TabPicture(0)   =   "frmBtlOrd.frx":9DDB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label62"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "text(87)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "text(86)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "text(85)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "text(84)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "text(83)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "text(74)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "text(27)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "text(26)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "text(25)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "text(24)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text(23)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame9"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "."
      TabPicture(1)   =   "frmBtlOrd.frx":AAC0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "text(34)"
      Tab(1).Control(2)=   "text(35)"
      Tab(1).Control(3)=   "text(36)"
      Tab(1).Control(4)=   "text(37)"
      Tab(1).Control(5)=   "text(38)"
      Tab(1).Control(6)=   "text(39)"
      Tab(1).Control(7)=   "text(40)"
      Tab(1).Control(8)=   "text(41)"
      Tab(1).Control(9)=   "text(42)"
      Tab(1).Control(10)=   "text(43)"
      Tab(1).Control(11)=   "imgTanda(0)"
      Tab(1).Control(12)=   "imgTanda(1)"
      Tab(1).Control(13)=   "ImgCrop"
      Tab(1).Control(14)=   "imgTanda(2)"
      Tab(1).Control(15)=   "imgTanda(3)"
      Tab(1).Control(16)=   "imgTanda(4)"
      Tab(1).Control(17)=   "imgTanda(5)"
      Tab(1).Control(18)=   "imgTanda(6)"
      Tab(1).Control(19)=   "imgTanda(7)"
      Tab(1).Control(20)=   "imgTanda(8)"
      Tab(1).Control(21)=   "imgTanda(9)"
      Tab(1).Control(22)=   "imgTanda(10)"
      Tab(1).Control(23)=   "imgTanda(11)"
      Tab(1).Control(24)=   "imgTanda(12)"
      Tab(1).Control(25)=   "imgTanda(13)"
      Tab(1).Control(26)=   "imgTanda(14)"
      Tab(1).Control(27)=   "imgTanda(15)"
      Tab(1).Control(28)=   "imgTanda(16)"
      Tab(1).Control(29)=   "imgTanda(17)"
      Tab(1).Control(30)=   "imgTanda(18)"
      Tab(1).Control(31)=   "imgTanda(19)"
      Tab(1).Control(32)=   "imgTanda(20)"
      Tab(1).Control(33)=   "imgTanda(21)"
      Tab(1).Control(34)=   "imgTanda(22)"
      Tab(1).Control(35)=   "imgTanda(23)"
      Tab(1).Control(36)=   "imgTanda(24)"
      Tab(1).Control(37)=   "imgTanda(25)"
      Tab(1).Control(38)=   "imgTanda(26)"
      Tab(1).Control(39)=   "imgTanda(27)"
      Tab(1).Control(40)=   "imgTanda(28)"
      Tab(1).Control(41)=   "imgTanda(29)"
      Tab(1).Control(42)=   "imgTanda(30)"
      Tab(1).Control(43)=   "imgTanda(31)"
      Tab(1).Control(44)=   "imgTanda(32)"
      Tab(1).Control(45)=   "imgTanda(33)"
      Tab(1).Control(46)=   "imgTanda(34)"
      Tab(1).Control(47)=   "imgTanda(35)"
      Tab(1).Control(48)=   "imgTanda(36)"
      Tab(1).Control(49)=   "imgTanda(37)"
      Tab(1).Control(50)=   "imgTanda(38)"
      Tab(1).Control(51)=   "imgTanda(39)"
      Tab(1).Control(52)=   "imgTanda(40)"
      Tab(1).Control(53)=   "imgTanda(41)"
      Tab(1).Control(54)=   "imgTanda(42)"
      Tab(1).Control(55)=   "imgTanda(43)"
      Tab(1).Control(56)=   "imgTanda(44)"
      Tab(1).Control(57)=   "imgTanda(45)"
      Tab(1).Control(58)=   "imgTanda(46)"
      Tab(1).Control(59)=   "imgTanda(47)"
      Tab(1).Control(60)=   "imgTanda(48)"
      Tab(1).Control(61)=   "imgTanda(49)"
      Tab(1).Control(62)=   "imgTanda(50)"
      Tab(1).Control(63)=   "imgTanda(51)"
      Tab(1).Control(64)=   "imgTanda(52)"
      Tab(1).Control(65)=   "imgTanda(53)"
      Tab(1).Control(66)=   "Label50"
      Tab(1).Control(67)=   "Label53"
      Tab(1).Control(68)=   "Image1"
      Tab(1).Control(69)=   "Shape5"
      Tab(1).ControlCount=   70
      TabCaption(2)   =   "."
      TabPicture(2)   =   "frmBtlOrd.frx":B5E0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "cmdWaktu"
      Tab(2).Control(3)=   "Frame12"
      Tab(2).Control(4)=   "cmdSublet"
      Tab(2).Control(5)=   "frmeWaktu"
      Tab(2).Control(6)=   "text(51)"
      Tab(2).Control(7)=   "text(78)"
      Tab(2).Control(8)=   "Label27"
      Tab(2).Control(9)=   "Label56"
      Tab(2).Control(10)=   "Shape6"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "."
      TabPicture(3)   =   "frmBtlOrd.frx":C486
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdFU"
      Tab(3).Control(1)=   "cboAlasan"
      Tab(3).Control(2)=   "Frame14"
      Tab(3).Control(3)=   "chUse"
      Tab(3).Control(4)=   "Frame8"
      Tab(3).Control(5)=   "Frame4"
      Tab(3).Control(6)=   "Frame2"
      Tab(3).Control(7)=   "cboStatus"
      Tab(3).Control(8)=   "cboAlasan1"
      Tab(3).Control(9)=   "Frame3"
      Tab(3).Control(10)=   "Frame13"
      Tab(3).Control(11)=   "cboPelanggan"
      Tab(3).Control(12)=   "text(52)"
      Tab(3).Control(13)=   "text(53)"
      Tab(3).Control(14)=   "text(54)"
      Tab(3).Control(15)=   "text(55)"
      Tab(3).Control(16)=   "text(75)"
      Tab(3).Control(17)=   "text(88)"
      Tab(3).Control(18)=   "text(89)"
      Tab(3).Control(19)=   "text(102)"
      Tab(3).Control(20)=   "text(90)"
      Tab(3).Control(21)=   "Label67"
      Tab(3).Control(22)=   "Label63"
      Tab(3).Control(23)=   "Label10"
      Tab(3).Control(24)=   "Label52"
      Tab(3).Control(25)=   "Label45"
      Tab(3).Control(26)=   "Label36"
      Tab(3).Control(27)=   "Label35"
      Tab(3).Control(28)=   "Label34"
      Tab(3).Control(29)=   "Shape3"
      Tab(3).ControlCount=   30
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
         Left            =   -68160
         Picture         =   "frmBtlOrd.frx":D295
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   4200
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         MaxLength       =   500
         TabIndex        =   216
         Top             =   2640
         Width           =   11535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         MaxLength       =   500
         TabIndex        =   215
         Top             =   2880
         Width           =   11535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         MaxLength       =   500
         TabIndex        =   214
         Top             =   3120
         Width           =   11535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         MaxLength       =   500
         TabIndex        =   213
         Top             =   3315
         Width           =   11535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
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
         Index           =   5
         Left            =   120
         MaxLength       =   500
         TabIndex        =   212
         Top             =   3540
         Width           =   11535
      End
      Begin VB.TextBox cboAlasan 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   -73800
         TabIndex        =   211
         Top             =   960
         Width           =   1935
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -68280
         TabIndex        =   207
         Top             =   1560
         Width           =   1815
         Begin VB.OptionButton opWash 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ya"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   209
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton opWash 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tidak"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   208
            Top             =   480
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WASHING"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   210
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.CheckBox chUse 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   -72480
         TabIndex        =   203
         Top             =   630
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "                           Kondisi ketika terjadi Keluhan                         "
         Height          =   2655
         Left            =   12600
         TabIndex        =   124
         Top             =   2520
         Width           =   4695
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   28
            Left            =   1485
            TabIndex        =   125
            Top             =   360
            Width           =   2970
            _ExtentX        =   5239
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
            Index           =   29
            Left            =   1485
            TabIndex        =   126
            Top             =   720
            Width           =   2970
            _ExtentX        =   5239
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
            Index           =   30
            Left            =   1485
            TabIndex        =   127
            Top             =   1080
            Width           =   2970
            _ExtentX        =   5239
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
            Index           =   31
            Left            =   1485
            TabIndex        =   128
            Top             =   1440
            Width           =   2970
            _ExtentX        =   5239
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
            Index           =   32
            Left            =   1485
            TabIndex        =   129
            Top             =   1800
            Width           =   2970
            _ExtentX        =   5239
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
            Index           =   33
            Left            =   1485
            TabIndex        =   130
            Top             =   2160
            Width           =   2970
            _ExtentX        =   5239
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Terjadi"
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
            TabIndex        =   135
            Top             =   405
            Width           =   1005
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
            Left            =   120
            TabIndex        =   134
            Top             =   765
            Width           =   690
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
            Left            =   120
            TabIndex        =   133
            Top             =   1125
            Width           =   435
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   120
            TabIndex        =   132
            Top             =   1485
            Width           =   1290
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
            Left            =   120
            TabIndex        =   131
            Top             =   1845
            Width           =   555
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "           WO          "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -68280
         TabIndex        =   116
         Top             =   2400
         Width           =   2175
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   64
            Left            =   840
            TabIndex        =   117
            Top             =   360
            Width           =   1200
            _ExtentX        =   2117
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
            Index           =   65
            Left            =   840
            TabIndex        =   118
            Top             =   645
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
            Index           =   66
            Left            =   840
            TabIndex        =   119
            Top             =   960
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
            Index           =   67
            Left            =   840
            TabIndex        =   120
            Top             =   1305
            Width           =   1200
            _ExtentX        =   2117
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
            Left            =   210
            TabIndex        =   123
            Top             =   1350
            Width           =   525
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
            Left            =   135
            TabIndex        =   122
            Top             =   390
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
            Left            =   135
            TabIndex        =   121
            Top             =   690
            Width           =   570
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "                     Estimasi                     "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   -66040
         TabIndex        =   98
         Top             =   1560
         Width           =   2685
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   56
            Left            =   900
            TabIndex        =   99
            Top             =   240
            Width           =   1620
            _ExtentX        =   2858
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
            Index           =   57
            Left            =   900
            TabIndex        =   100
            Top             =   555
            Width           =   1620
            _ExtentX        =   2858
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
            Height          =   405
            Index           =   61
            Left            =   900
            TabIndex        =   101
            Top             =   2280
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
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
            Index           =   58
            Left            =   900
            TabIndex        =   102
            Top             =   1455
            Width           =   1620
            _ExtentX        =   2858
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
            Index           =   59
            Left            =   900
            TabIndex        =   103
            Top             =   840
            Width           =   1620
            _ExtentX        =   2858
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
            Index           =   60
            Left            =   900
            TabIndex        =   104
            Top             =   1155
            Width           =   1620
            _ExtentX        =   2858
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
            Index           =   62
            Left            =   960
            TabIndex        =   105
            Top             =   3000
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
            Index           =   63
            Left            =   960
            TabIndex        =   106
            Top             =   3360
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
            Index           =   77
            Left            =   900
            TabIndex        =   107
            Top             =   1755
            Width           =   1620
            _ExtentX        =   2858
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
            TabIndex        =   115
            Top             =   900
            Width           =   180
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
            TabIndex        =   114
            Top             =   1200
            Width           =   450
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
            TabIndex        =   113
            Top             =   2325
            Width           =   645
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
            TabIndex        =   112
            Top             =   600
            Width           =   300
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
            Left            =   225
            TabIndex        =   111
            Top             =   285
            Width           =   495
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
            TabIndex        =   110
            Top             =   1500
            Width           =   570
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Selesai"
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
            Left            =   180
            TabIndex        =   109
            Top             =   3000
            Width           =   660
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
            TabIndex        =   108
            Top             =   1845
            Width           =   660
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   2520
            Y1              =   2160
            Y2              =   2160
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -66480
         TabIndex        =   94
         Top             =   480
         Width           =   1335
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   96
            Top             =   675
            Width           =   975
         End
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   95
            Top             =   420
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SERVICE"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   97
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frmBtlOrd.frx":D9A9
         Left            =   -73800
         List            =   "frmBtlOrd.frx":D9B3
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboAlasan1 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frmBtlOrd.frx":D9C1
         Left            =   -71880
         List            =   "frmBtlOrd.frx":D9C3
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   90
         Top             =   3240
         Width           =   8055
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   1455
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2566
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "ITEM PART"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA PART"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "KODE SATUAN"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "QTY"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   4
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "GANTI"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "BAWA SENDIRI"
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "OLI"
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
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
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
            _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
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
            Index           =   48
            Left            =   960
            TabIndex        =   187
            Top             =   195
            Visible         =   0   'False
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
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   270
            Left            =   3900
            TabIndex        =   188
            Top             =   195
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
            Left            =   4005
            TabIndex        =   189
            Top             =   195
            Visible         =   0   'False
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
            TabIndex        =   190
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   2955
         Left            =   -74880
         TabIndex        =   76
         Top             =   360
         Width           =   8040
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   46
            Left            =   960
            TabIndex        =   77
            Top             =   555
            Visible         =   0   'False
            Width           =   2850
            _ExtentX        =   5027
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
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   270
            Left            =   3840
            TabIndex        =   78
            Top             =   555
            Visible         =   0   'False
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   2025
            Left            =   120
            TabIndex        =   79
            Top             =   840
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3572
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            HighLight       =   2
            ScrollBars      =   0
            SelectionMode   =   1
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   44
            Left            =   960
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
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
            Left            =   1485
            TabIndex        =   81
            Top             =   240
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
            Index           =   45
            Left            =   1635
            TabIndex        =   82
            Top             =   240
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   47
            Left            =   3945
            TabIndex        =   83
            Top             =   555
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
            Left            =   5040
            TabIndex        =   84
            Top             =   555
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
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
            Mask            =   ">CCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   270
            Left            =   5850
            TabIndex        =   85
            Top             =   555
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
            Index           =   82
            Left            =   6000
            TabIndex        =   86
            Top             =   555
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
            Left            =   495
            TabIndex        =   89
            Top             =   285
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Order "
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
            TabIndex        =   88
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MCH"
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
            TabIndex        =   87
            Top             =   600
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0FF&
         Caption         =   " ^ Meter Bahan Bakar ^ "
         Height          =   615
         Left            =   -71040
         TabIndex        =   70
         Top             =   3480
         Width           =   4695
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Penuh"
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
            Left            =   3600
            TabIndex        =   75
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00C0C0FF&
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
            Left            =   2760
            TabIndex        =   74
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00C0C0FF&
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
            Left            =   1920
            TabIndex        =   73
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00C0C0FF&
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
            Left            =   1080
            TabIndex        =   72
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opBBM 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Kosong"
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
            TabIndex        =   71
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -65170
         TabIndex        =   65
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton opByr 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opByr 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   580
            Width           =   1335
         End
         Begin VB.OptionButton opByr 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   66
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PEMBAYARAN"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   300
         TabIndex        =   58
         Top             =   720
         Width           =   4695
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
            TabIndex        =   61
            Top             =   960
            Width           =   1215
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
            TabIndex        =   60
            Top             =   720
            Width           =   975
         End
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
            TabIndex        =   59
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   21
            Left            =   1560
            TabIndex        =   62
            Top             =   480
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
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
            Index           =   22
            Left            =   1560
            TabIndex        =   63
            Top             =   840
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
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
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIPE SERVICE"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1560
            TabIndex        =   64
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdWaktu 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Waktu Terpakai"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66800
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   480
         Width           =   3450
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "               BAWA PART            "
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
         Left            =   -66840
         TabIndex        =   51
         Top             =   3240
         Width           =   3495
         Begin MSMask.MaskEdBox text 
            Height          =   240
            Index           =   69
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   423
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
            Height          =   240
            Index           =   70
            Left            =   120
            TabIndex        =   53
            Top             =   705
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   423
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
            Height          =   240
            Index           =   71
            Left            =   120
            TabIndex        =   54
            Top             =   930
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   423
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
            Height          =   240
            Index           =   72
            Left            =   120
            TabIndex        =   55
            Top             =   1155
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   423
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
            Height          =   240
            Index           =   73
            Left            =   120
            TabIndex        =   56
            Top             =   1380
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   423
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
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   -68280
         TabIndex        =   47
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton opGanti 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton opGanti 
            BackColor       =   &H00FFFFFF&
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GANTI PART"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   1485
         End
      End
      Begin VB.CommandButton cmdSublet 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Sublet"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -65640
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1920
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.ComboBox cboPelanggan 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frmBtlOrd.frx":D9C5
         Left            =   -73800
         List            =   "frmBtlOrd.frx":D9C7
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Frame frmeWaktu 
         BackColor       =   &H00C0FFFF&
         Height          =   3375
         Left            =   -63240
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   6135
         Begin TrueOleDBGrid70.TDBGrid GridWaktu1 
            Height          =   3855
            Left            =   8640
            TabIndex        =   42
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
            TabIndex        =   43
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
            TabIndex        =   44
            Top             =   2880
            Width           =   4425
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   240
         Index           =   23
         Left            =   7020
         TabIndex        =   136
         Top             =   840
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   24
         Left            =   7020
         TabIndex        =   137
         Top             =   1065
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   25
         Left            =   7020
         TabIndex        =   138
         Top             =   1290
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   26
         Left            =   7020
         TabIndex        =   139
         Top             =   1515
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   27
         Left            =   7020
         TabIndex        =   140
         Top             =   1740
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Height          =   270
         Index           =   52
         Left            =   -71880
         TabIndex        =   141
         Top             =   960
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   476
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
         Height          =   270
         Index           =   53
         Left            =   -73680
         TabIndex        =   142
         Top             =   4080
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   476
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   54
         Left            =   -73680
         TabIndex        =   143
         Top             =   3495
         Width           =   5235
         _ExtentX        =   9234
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
         Index           =   55
         Left            =   -73680
         TabIndex        =   144
         Top             =   3795
         Width           =   5235
         _ExtentX        =   9234
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
         Height          =   240
         Index           =   34
         Left            =   -70080
         TabIndex        =   145
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   35
         Left            =   -70080
         TabIndex        =   146
         Top             =   825
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   36
         Left            =   -70080
         TabIndex        =   147
         Top             =   1050
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   37
         Left            =   -70080
         TabIndex        =   148
         Top             =   1275
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   38
         Left            =   -70080
         TabIndex        =   149
         Top             =   1500
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   39
         Left            =   -70080
         TabIndex        =   150
         Top             =   1920
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   40
         Left            =   -70080
         TabIndex        =   151
         Top             =   2145
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   41
         Left            =   -70080
         TabIndex        =   152
         Top             =   2370
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   42
         Left            =   -70080
         TabIndex        =   153
         Top             =   2595
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Height          =   240
         Index           =   43
         Left            =   -70080
         TabIndex        =   154
         Top             =   2820
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   423
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
         Index           =   51
         Left            =   -65610
         TabIndex        =   155
         Top             =   1320
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   74
         Left            =   12240
         TabIndex        =   156
         Top             =   1920
         Visible         =   0   'False
         Width           =   2970
         _ExtentX        =   5239
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
         Height          =   270
         Index           =   75
         Left            =   -71880
         TabIndex        =   157
         Top             =   1245
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   476
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
         Index           =   78
         Left            =   -65610
         TabIndex        =   158
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
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
         Index           =   88
         Left            =   -73680
         TabIndex        =   192
         Top             =   2400
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
         Index           =   89
         Left            =   -73680
         TabIndex        =   193
         Top             =   2760
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
         Height          =   240
         Index           =   83
         Left            =   360
         TabIndex        =   197
         Top             =   2640
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   84
         Left            =   360
         TabIndex        =   198
         Top             =   2865
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   85
         Left            =   360
         TabIndex        =   199
         Top             =   3090
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   86
         Left            =   360
         TabIndex        =   200
         Top             =   3315
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   87
         Left            =   360
         TabIndex        =   201
         Top             =   3540
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   423
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   102
         Left            =   -73680
         TabIndex        =   204
         Top             =   4635
         Width           =   2835
         _ExtentX        =   5001
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
         Height          =   270
         Index           =   90
         Left            =   -73680
         TabIndex        =   206
         Top             =   4360
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   476
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
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   0
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":D9C9
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   1
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":E163
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image ImgCrop 
         Height          =   2100
         Left            =   -73920
         Top             =   1440
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   2
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":E8FD
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   3
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":F097
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   4
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":F831
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   5
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":FFCB
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   6
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":10765
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   7
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":10EFF
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   8
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":11699
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   9
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":11E33
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   10
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":125CD
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   11
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":12D67
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   12
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":13501
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   13
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":13C9B
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   14
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":14435
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   15
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":14BCF
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   16
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":15369
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   17
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":15B03
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   18
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":1629D
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   19
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":16A37
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   20
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":171D1
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   21
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":1796B
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   22
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":18105
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   23
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":1889F
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   24
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":19039
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   25
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":197D3
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   26
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":19F6D
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   27
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":1A707
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   28
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":1AEA1
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   29
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":1B63B
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   30
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":1BDD5
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   31
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":1C56F
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   32
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":1CD09
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   33
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":1D4A3
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   34
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":1DC3D
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   35
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":1E3D7
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   36
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":1EB71
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   37
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":1F30B
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   38
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":1FAA5
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   39
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":2023F
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   40
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":209D9
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   41
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":21173
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   42
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":2190D
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   43
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":220A7
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   44
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":22841
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   45
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":22FDB
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   46
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":23775
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   47
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":23F0F
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   48
         Left            =   -73920
         Picture         =   "frmBtlOrd.frx":246A9
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   49
         Left            =   -73680
         Picture         =   "frmBtlOrd.frx":24E43
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   50
         Left            =   -73440
         Picture         =   "frmBtlOrd.frx":255DD
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   51
         Left            =   -73200
         Picture         =   "frmBtlOrd.frx":25D77
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   52
         Left            =   -72960
         Picture         =   "frmBtlOrd.frx":26511
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   210
      End
      Begin VB.Image imgTanda 
         Height          =   195
         Index           =   53
         Left            =   -72720
         Picture         =   "frmBtlOrd.frx":26CAB
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   210
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
         Height          =   255
         Left            =   -74880
         TabIndex        =   205
         Top             =   4650
         Width           =   1035
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   240
         TabIndex        =   202
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu Penyerahan"
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
         Left            =   -74700
         TabIndex        =   194
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   7080
         TabIndex        =   169
         Top             =   600
         Width           =   1545
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
         Left            =   -74385
         TabIndex        =   168
         Top             =   4140
         Width           =   585
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -74970
         TabIndex        =   167
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
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
         Left            =   -74385
         TabIndex        =   166
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -74625
         TabIndex        =   165
         Top             =   3810
         Width           =   825
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
         Height          =   255
         Left            =   -74835
         TabIndex        =   164
         Top             =   3510
         Width           =   1035
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kondisi Kendaraan"
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
         Left            =   -71040
         TabIndex        =   163
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barang Berharga Tertinggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -71040
         TabIndex        =   162
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "K. Meter"
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
         Left            =   -66375
         TabIndex        =   161
         Top             =   1365
         Width           =   615
      End
      Begin VB.Label Label34 
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
         Left            =   -74520
         TabIndex        =   160
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label56 
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
         Left            =   -66615
         TabIndex        =   159
         Top             =   1005
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   2970
         Left            =   -74280
         Picture         =   "frmBtlOrd.frx":27445
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2055
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   -74980
         Top             =   360
         Width           =   11755
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   -74980
         Top             =   360
         Width           =   11760
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   20
         Top             =   360
         Width           =   11760
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   -74980
         Top             =   360
         Width           =   11760
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   4100
      TabIndex        =   183
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
      Index           =   68
      Left            =   2640
      TabIndex        =   184
      Top             =   8700
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CommandButton TblBatal 
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
      Height          =   400
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmBtlOrd.frx":29E12
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   8640
      Width           =   1245
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
      Left            =   10080
      Picture         =   "frmBtlOrd.frx":2A2DA
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   8640
      Width           =   2010
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8895
      Left            =   0
      Picture         =   "frmBtlOrd.frx":2A9A9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   400
   End
   Begin VB.Image tndBatal 
      Height          =   345
      Left            =   4680
      Picture         =   "frmBtlOrd.frx":2D71F
      Top             =   75
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label60 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Order"
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
      TabIndex        =   186
      Top             =   158
      Width           =   720
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
      Left            =   495
      TabIndex        =   185
      Top             =   8700
      Visible         =   0   'False
      Width           =   2145
   End
End
Attribute VB_Name = "frmBtlOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 18 April 2009 / 21 Feb 2012
'Created Coding  : Thania / Acep
'Tgl. Modifikasi : 08 Febuary 2012 / Acep
'========================================

Dim rsTemp As New ADODB.Recordset
Dim rsTempStall As New ADODB.Recordset
Dim rsGridPart As New ADODB.Recordset
Dim rsOli As New ADODB.Recordset
Dim pesan As String
Dim baris As Integer
Dim ntRate As Double
Dim nk As Integer
Dim KorX(100) As Integer
Dim KorY(100) As Integer
Dim bisaBatal As Boolean

Private Sub hapusForm()
    For i = 1 To 81
        Select Case i
            Case 9, 20, 49, 51, 56, 57, 58, 59, 60, 61, 77
                Text(i) = 0
            Case 6, 7, 8
                Text(i).Mask = ""
                Text(i) = ">>>>->>>>>>>>>>"
                Text(i).Mask = "9999-9999999999"
            Case 14
                Text(i).Mask = ""
                Text(i) = ">>->>>>->>>"
                Text(i).Mask = "??-9999-???"
            Case 16, 17, 18, 62, 65, 79
                Text(i).Mask = ""
                Text(i) = ">>->>->>>>"
                Text(i).Mask = "99-99-9999"
            Case 47
                Text(i).Mask = ""
                Text(i) = "00,00"
                Text(i).Mask = "99,99"
            Case 63, 66
                Text(i).Mask = ""
                Text(i) = ">>:>>"
                Text(i).Mask = "99:99"
            Case Else
                Text(i) = ""
        End Select
    Next
    For i = 0 To 53
        imgTanda(i).Visible = False
    Next
    For i = 0 To 5
        Text1(i) = ""
    Next
    chUse.value = Unchecked
    opBBM(4).value = True
    CboStatus.AddItem "EXT"
    CboStatus.AddItem "INT"
    CboStatus.AddItem "AFFPPN"
    
    Call aturGridJob
    For i = 0 To 7
        dGrid.TextMatrix(i, 1) = ""
        dGrid.TextMatrix(i, 2) = ""
        dGrid.TextMatrix(i, 3) = ""
        dGrid.TextMatrix(i, 4) = ""
        dGrid.TextMatrix(i, 5) = ""
        dGrid.TextMatrix(i, 6) = ""
    Next
    
    If rsGridPart.State = adStateOpen Then
        rsGridPart.Close
        Set rsGridPart = Nothing
    End If

    
    ' ISI KARAKTER PELANGGAN
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = objO.FetchData("SELECT * FROM T_Karakter WITH(NOLOCK) ORDER BY KODE")
    If Not rsTemp.EOF Then
        cboPelanggan.Clear
        Do While Not rsTemp.EOF
            cboPelanggan.AddItem (rsTemp("KODE"))
            rsTemp.MoveNext
        Loop
    End If
    cboPelanggan.ListIndex = 0
    Set rsGridPart = Nothing
    dGrid.Enabled = False
    partGrid.Enabled = False
End Sub

Private Sub aturGridJob()

    baris = 1
    dGrid.Rows = 11
    dGrid.cols = 7
    dGrid.ColWidth(0) = 0
    dGrid.ColWidth(1) = 3500
    dGrid.ColWidth(2) = 800
    dGrid.ColWidth(3) = 0
    dGrid.ColWidth(4) = 0
    dGrid.ColWidth(5) = 0
    dGrid.ColWidth(6) = 0
    dGrid.ColAlignment(1) = flexAlignLeftCenter
    dGrid.TextMatrix(0, 1) = "Keluhan & Order Booking"
End Sub

Private Sub PanggilDataOrder(nRgk As String, nWO As String)
Dim rsTmp As New ADODB.Recordset
Dim rsOli As New ADODB.Recordset
Dim n As Integer
        
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = objO.FetchData("SELECT a.*,c.*,d.*,b.gsm,b.cdma,b.tgl_rep,b.rcdate,b.fs_km_akhir, b.Ddate, b.FS_kd_dealer, b.Fs_nm_dealer " _
                             & "FROM jobdta a WITH(NOLOCK) INNER JOIN custdat b WITH(NOLOCK) ON a.no_chassis = b.no_chassis AND a.no_polisi = b.no_polisi " _
                             & "LEFT JOIN jobpart c WITH(NOLOCK) on a.no_chassis = c.no_chassis and a.no_wo = c.no_wo " _
                             & "LEFT JOIN jobdtl d WITH(NOLOCK) on a.no_chassis = d.no_chassis and a.no_wo = d.no_wo " _
                             & "WHERE a.no_chassis = '" & nRgk & "' AND a.NO_WO = '" & nWO & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' ") 'AND TIPEWO <> 'W'")
                             
    ''& "LEFT join jobKoord e on a.no_chassis = e.no_chassis and a.no_wo = e.no_wo"
    If Not rsTemp.EOF Then
        Text(1) = rsTemp("FS_KD_PLG")
        Text(2) = rsTemp("FS_KD_WIL")
        Text(3) = rsTemp("CUSTOMER")
        Text(4) = Left(rsTemp("CUST_ADRES"), 60)
        Text(5) = rsTemp("KOTA")
        Text(6) = rsTemp("TELPON")
        Text(7) = rsTemp("GSM")
        Text(8) = rsTemp("CDMA")
        Text(9) = rsTemp("FS_TH_MOBIL")
        Text(10) = rsTemp("MDL")
        Text(11) = rsTemp("MODEL")
        Text(12) = rsTemp("FMDL")
        Text(13) = rsTemp("COLOR")
        Text(14).Mask = ""
        Text(14) = rsTemp("NO_POLISI")
        Text(15) = Nul(rsTemp("NO_CHASSIS"))
        Text(16).Mask = ""
        Text(16) = rsTemp("DOTAM")
        Text(20) = rsTemp("FS_KM_AKHIR")
        
        ' PERMINTAAN CUSTOMER
        Dim hlp As String
        For i = 1 To 5
            hlp = "REQ" & i
            Text(22 + i) = rsTemp(hlp)
        Next
        
        ' KELUHAN CUSTOMER
        For i = 1 To 5
            hlp = "CMPL" & i
            Text1(0 + i) = rsTemp(hlp)
        Next
        
        ' KONDISI KETIKA TERJADI KERUSAKAN
        For i = 1 To 6
            hlp = "COND" & i
            Text(27 + i) = Nul(rsTemp(hlp))
        Next
        
        ' WALK AROUND CHECK
        ' %%%% KONDISI KENDARAAN %%%%
        For i = 1 To 5
            hlp = "WAK" & i
            Text(33 + i) = Nul(rsTemp(hlp))
        Next
        
        ' %%%% BARANG BERHARGA %%%%
        For i = 1 To 5
            hlp = "WAE" & i
            Text(38 + i) = Nul(rsTemp(hlp))
        Next
        
        chUse.value = IIf(rsTemp("HEAVYUSE") = 1, Checked, Unchecked)
        
        ' %%%% KOORDINAT %%%%
        nk = 0
        If rsTmp.State Then rsTmp.Close
        Set rsTmp = objO.FetchData("SELECT * FROM jobKoord WITH(NOLOCK) WHERE no_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                imgTanda(nk).Top = Image1.Top + rsTmp("KOORDY") - (imgTanda(nk).Height / 2)
                imgTanda(nk).Left = Image1.Left + rsTmp("KOORDX") - (imgTanda(nk).Width / 2)
                imgTanda(nk).Visible = True
                KorX(nk) = rsTmp("KOORDX")
                KorY(nk) = rsTmp("KOORDY")
                nk = nk + 1
                rsTmp.MoveNext
            Loop
        End If
        
        ' BBM
        n = IIf(Nul(rsTemp("BBM")) = "", 4, Nul(rsTemp("BBM")))
        opBBM(n).value = True
        
        ' ISI JOBORDER
        ntRate = 0
        myChassis = Trim(Text(15))
        For i = 1 To 8
            tmp1 = "REP" & i
            tmp2 = "RTE" & i
            tmp3 = "MEKANIK" & i
            dGrid.TextMatrix(i, 1) = rsTemp(tmp1)
            If rsTemp(tmp1) <> "" Then
                baris = baris + 1
                If rsTmp.State Then rsTmp.Close
                Set rsTmp = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE FS_ITEM = '" & rsTemp(tmp1) & "' AND MDL = '" & Trim(Text(10)) & "'")
                If Not rsTmp.EOF Then
                    dGrid.TextMatrix(i, 3) = False '""
                    dGrid.TextMatrix(i, 4) = rsTmp("JOB")
                Else
                    dGrid.TextMatrix(i, 3) = True '"fOther"
                    'DGrid.TextMatrix(i, 4) = RStmp("JOB")
                End If
                dGrid.TextMatrix(i, 2) = rsTemp(tmp2)
                dGrid.TextMatrix(i, 5) = hitungLabour(CDbl(rsTemp(tmp2)), Text(11))
                dGrid.TextMatrix(i, 6) = rsTemp(tmp3)
                ntRate = ntRate + CDbl(rsTemp(tmp2))
            End If
        Next
        
        ' ISI FU
        If rsTmp.State Then rsTmp.Close
        Set rsTmp = objO.FetchData("SELECT * FROM JOBDTL WITH(NOLOCK) WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
        If Not rsTmp.EOF Then
            If Trim(rsTmp("TELPFU")) <> "" And Not IsNull(rsTemp("TELPFU")) Then
                Call objO.InsertData("INSERT INTO TEMPFU SELECT NO_WO, NO_CHASSIS, TGLFU, JAMFU, TELPFU, KETFU,psfu_owner_flag FROM jobDtl WITH(NOLOCK) " _
                             & "WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' AND " _
                             & "NOT EXISTS (SELECT * FROM TEMPFU WITH(NOLOCK) WHERE NO_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "')")
            End If
        End If
        
        ' ISI PART
        If rsGridPart.State <> adStateClosed Then rsGridPart.Close
        Set rsGridPart = Nothing
        
        If rsTmp.State Then rsTmp.Close
        Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM jobPart A WITH(NOLOCK) LEFT JOIN T_OLI B WITH(NOLOCK) ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & rsTemp("NO_CHASSIS") & "' AND no_wo = '" & rsTemp("NO_WO") & "'")
        If Not rsTmp.EOF Then
            If rsGridPart.State = adStateClosed Then createTblPart
            Do While Not rsTmp.EOF
                With rsGridPart
                    .AddNew
                    !KodeItem = rsTmp("KODE_ITEM")
                    !NAMAITEM = rsTmp("NAMA_ITEM")
                    !SATITEM = rsTemp("SAT_ITEM")
                    !qty = CDbl(rsTmp("QTY"))
                    !ganti = IIf(rsTmp("GANTI") <> "N", True, False)
                    !BAWA = IIf(rsTmp("BAWA") <> "N", True, False)
                    !oli = IIf(rsTmp("KDOLI") <> "", True, False)
                    .Update
                End With
                n = 69
                If rsTmp("PARTBAWA") <> "" Then
                    Text(n) = rsTmp("PARTBAWA")
                    n = n + 1
                End If
            rsTmp.MoveNext
            Loop
            Call HeadColumn
            Set partGrid.DataSource = rsGridPart
        End If
        If rsTmp.State Then rsTmp.Close
        
        opTpSvc(2).Enabled = False
        'cmdWarranty.Enabled = False
        If rsTemp("TPSVC") = "W" Then
            opTpSvc(2).value = True
            opTpSvc(2).Enabled = True
            'cmdWarranty.Enabled = True
        ElseIf rsTemp("TPSVC") = "P" Then
            opTpSvc(1).value = True
        ElseIf rsTemp("TPSVC") = "" Then
            opTpSvc(0).value = True
        End If
            
        ' FIELD _ FIELD LAIN
        n = IIf(Nul(rsTemp("GANTIPART")) = "", 0, Nul(rsTemp("GANTIPART")))
        opGanti(n).value = True
        
        n = IIf(Nul(rsTemp("TUNGGUSVC")) = "", 0, Nul(rsTemp("TUNGGUSVC")))
        opSvc(n).value = True
        
        n = IIf(Nul(rsTemp("PEMBAYARAN")) = "CS", 0, IIf(rsTemp("PEMBAYARAN") = "CR", 1, 2))
        opByr(n).value = True
        
        n = IIf(Nul(rsTemp("WASHING")) = "0", 1, 0)
        opWash(n).value = True
        
        Text(50) = rsTemp("MECH")
        Text(51) = rsTemp("KM")
        Text(53) = rsTemp("CATAT")
        Text(54) = rsTemp("REMARK1")
        Text(55) = rsTemp("SUGGEST")
        Text(56) = rsTemp("LABOR")
        Text(57) = rsTemp("PART")
        Text(58) = rsTemp("MTRL")
        Text(59) = rsTemp("OIL")
        Text(60) = rsTemp("SBLT")
        Text(61) = rsTemp("TOTAL")
        Text(61) = CDbl(Text(56)) + CDbl(Text(57)) + CDbl(Text(58)) + CDbl(Text(59)) + CDbl(Text(60)) + CDbl(Text(77))
        
        ' %%%% KOORDINAT %%%%
        nk = 0
        If rsTmp.State Then rsTmp.Close
        Set rsTmp = objO.FetchData("SELECT * FROM jobKoord WITH(NOLOCK) WHERE no_WO = '" & rsTemp("NO_WO") & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "'")
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                imgTanda(nk).Top = Image1.Top + rsTmp("KOORDY") - (imgTanda(nk).Height / 2)
                imgTanda(nk).Left = Image1.Left + rsTmp("KOORDX") - (imgTanda(nk).Width / 2)
                imgTanda(nk).Visible = True
                KorX(nk) = rsTmp("KOORDX")
                KorY(nk) = rsTmp("KOORDY")
                nk = nk + 1
                rsTmp.MoveNext
            Loop
        End If
        rsTmp.Close

        
        ' HITUNG NILAI PART & OLI
        If rsGridPart.State <> adStateClosed Then
            Text(59) = 0
            Text(57) = 0
            If Not rsGridPart.EOF Then
            rsGridPart.MoveFirst
            Do While Not rsGridPart.EOF
                    If rsOli.State Then rsOli.Close
                    Set rsOli = objO.FetchDataACC("SELECT hargaJualA FROM inmst WITH(NOLOCK) WHERE kodeItem = '" & rsGridPart("KODEITEM") & "'")
                    If Not rsOli.EOF Then
                        If rsGridPart("OLI") Then
                            Text(59) = Text(59) + CDbl(rsOli("HARGAJUALA"))
                        Else
                            Text(57) = Text(57) + CDbl(rsOli("HARGAJUALA"))
                        End If
                    End If
                    rsOli.Close
                rsGridPart.MoveNext
            Loop
            End If
        End If
        
        CboStatus = rsTemp("STATUS")
        Text(65) = TglDDMMYYYY(rsTemp("DATE"))
        Text(66) = rsTemp("TIME")
        If Text(66) <> ">>:>>" And Text(65) <> ">>->>->>>>" Then
            nWKT = hitungWaktu(CDbl(ntRate), Text(66), Text(65))
            Text(63) = Left(nWKT, InStr(nWKT, "|") - 1)
            Text(62) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
        End If
        Text(88) = TglDDMMYYYY(rsTemp("DATESERAH"))
        Text(89) = rsTemp("TIMESERAH")
        Text(64) = Nul(rsTemp("NO_WO"))
        Text(67) = rsTemp("NOINV")
        Text(78) = rsTemp("NOFAK")
        ChkReturnOrder = IIf(rsTemp("FN_RETURN") = 1, Checked, Unchecked)
        chTHS = IIf(rsTemp("THS") = 1, Checked, Unchecked)
        cboPelanggan = Trim(rsTemp("KARAKTER"))
        Text(75) = Trim(rsTemp("DTLKARAKTER"))
        cboAlasan = Trim(rsTemp("DATANG"))
        Text(52) = Trim(rsTemp("DTLDATANG"))
        
        ' CHECK STALL
        If rsTmp.State Then rsTmp.Close
        Set rsTmp = objO.FetchData("SELECT KETERANGAN FROM stall WITH(NOLOCK) WHERE KETERANGAN like '%|" & nWO & "|%'")
        If Not rsTmp.EOF Then
            'cmdStall.Tag = ""
            Call objO.InsertData("INSERT INTO tempStall SELECT noStall, tanggal, JamMulai, JamSelesai, rate, [Group], Mech, Flag, Keterangan,adjust FROM STALL WITH(NOLOCK) " _
                         & "WHERE KETERANGAN like '%|" & nWO & "|%' AND Tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "' AND NOT EXISTS " _
                         & "(SELECT * FROM TEMPSTALL WITH(NOLOCK) WHERE KETERANGAN like '%|" & nWO & "|%' AND Tanggal = '" & Format(Text(65), "mm/dd/yyyy") & "')")
        End If
    End If
End Sub

Private Sub panggilHistory(noRgk As String, noPols As String)
    Dim rsHist As ADODB.Recordset
    Set rsHist = objO.FetchData("SELECT a.*, rcdate FROM History a WITH(NOLOCK) INNER JOIN custdat b WITH(NOLOCK) ON a.no_chassis = b.no_chassis WHERE a.no_chassis = '" & Trim(noRgk) & "' ORDER BY [date] DESC")
    If Not rsHist.EOF Then
        Text(17) = Format(Trim(rsHist("DATE")), "dd-mm-yyyy")
        Text(18) = Format(Trim(rsHist("RCDATE")), "dd-mm-yyyy")
        Text(19) = Nul(Trim(rsHist("FS_NM_DEALER")))
        Text(20) = Format(CDbl(rsHist("KM")), "0#,#")
        If DateDiff("d", Format(rsHist("DATE"), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) > 15 Then
            ChkReturnOrder.Enabled = False
        Else
            ChkReturnOrder.Enabled = True
        End If
        cmdHistory.Enabled = True
        Command1.Enabled = True
    Else
        ChkReturnOrder.Enabled = False
        'cmdHistory.Enabled = False
    End If
    cmdHistory.Enabled = True
    Command1.Enabled = True
End Sub


Function createTblPart() As Boolean
1         On Error GoTo vb_error
          
2         If rsGridPart.State Then rsGridPart.Close
3         Set rsGridPart = Nothing
4         With rsGridPart
5             .Fields.Append "KODEITEM", adVarChar, 50
6             .Fields.Append "NAMAITEM", adVarChar, 200
7             .Fields.Append "SATITEM", adVarChar, 30
8             .Fields.Append "QTY", adInteger
9             .Fields.Append "GANTI", adBoolean
10            .Fields.Append "BAWA", adBoolean
11            .Fields.Append "OLI", adBoolean
12            .Open
13        End With
14        If rsGridPart.State Then createTblPart = True
15        Exit Function
          
vb_error:
16        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub HeadColumn()
    partGrid.Columns(0).Caption = "KODE PART"
    partGrid.Columns(0).DataField = "KODEITEM"
    partGrid.Columns(0).Width = 1500
    partGrid.Columns(1).Caption = "NAMA PART"
    partGrid.Columns(1).DataField = "NAMAITEM"
    partGrid.Columns(1).Width = 2500
    partGrid.Columns(2).Caption = "SATUAN PART"
    partGrid.Columns(2).DataField = "SATITEM"
    partGrid.Columns(2).Width = 800
    partGrid.Columns(3).Caption = "QTY"
    partGrid.Columns(3).DataField = "QTY"
    partGrid.Columns(3).Width = 800
    partGrid.Columns(4).Caption = "GANTI"
    partGrid.Columns(4).DataField = "GANTI"
    partGrid.Columns(4).Alignment = dbgCenter
    partGrid.Columns(5).Caption = "PART SENDIRI"
    partGrid.Columns(5).DataField = "BAWA"
    partGrid.Columns(5).Alignment = dbgCenter
    partGrid.Columns(6).Caption = "OLI"
    partGrid.Columns(6).DataField = "OLI"
    partGrid.Columns(6).Width = 0
End Sub

Private Sub cmdCancel_Click()
    frmBtlOrder.Visible = False
End Sub

Private Sub cmdHistory_Click()
'    gTag = "BTLORDER"
'    frmHistory.text(0) = text(15)
'    frmHistory.text(1) = text(3)
'    frmHistory.text(2) = Replace(Replace(text(14), ">", ""), " ", "")
'    frmHistory.Show (vbModal)
    
    If Replace(Replace(Text(15), ">", ""), " ", "") = "" Then Exit Sub
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
    dRangka = Text(15)
    nData(0) = Text(15)
    nData(1) = Text(3)
    nData(2) = Replace(Replace(Text(14), ">", ""), " ", "")
    Dim fH As New FrmHistory_new
    fH.show (vbModal)
End Sub
Private Sub Command1_Click()
    If Replace(Replace(Text(15), ">", ""), " ", "") = "" Then Exit Sub
    gTag = "LOCAL"
    dRangka = Text(15)
    nData(0) = Text(15)
    nData(1) = Text(3)
    nData(2) = Replace(Replace(Text(14), ">", ""), " ", "")
    Dim fH As New FrmHistory_new
    fH.show (vbModal)
End Sub

Private Sub cmdOK_Click()

1         On Error GoTo vbError
2         If Text1(0) <> "" Then
3             Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlOrd", "cmdOk", 0)
                  
4             If rsTemp.State Then rsTemp.Close
5             Set rsTemp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Text(0) & "'")
                  
              ' KURANG JAM MEKANIK
              Dim hlp As String
              Dim rsAct As ADODB.Recordset
              Dim nTotAct As Integer
              
6             nTotAct = 0
7             For i = 1 To 7
8                 hlp = "REP" & i
9                 If rsTemp(hlp) <> "" Then
                      
10                  '  If Log.RPLHistori = 0 Then
11                        Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(rsTemp(hlp)) & "' AND MDL = '" & Trim(rsTemp("MDL")) & "'")
12                 '   Else
13                  '      Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM [" & dbReplikasiHO & "]..ACTRATE WITH(NOLOCK) WHERE FS_ITEM = '" & Trim(rsTemp(hlp)) & "' AND MDL = '" & Trim(rsTemp("MDL")) & "'")
14                    'End If
15                    If Not rsAct.EOF Then
16                        nTotAct = nTotAct + CDbl(rsAct("RATE"))
17                    End If
18                    rsAct.Close
19                End If
20            Next

21            erTranAll = False
22            onTR = True
23            Call objHT.CloseTransactionAll_Err
24            Call objHT.OpenTransactionAll

25            If Not rsTemp.EOF Then
                  ' PROSES PEMBATALAN ORDER
26                If erTranAll = False Then Call simpanAntrian(2, "", 0, Format(gTglServer, "dd-mm-yyyy"), Trim(Text(0)), Trim(Text(14)), Format(Text(88) & " " & Text(89), "yyyy-mm-dd hh:mm"))
27                sSQLHO = "UPDATE JOBDTA SET DELIV = 'BTL', ALASAN = '" & Text1(0) & "', FS_LOG_DEL = '" & gSA & "', FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY") & "' WHERE NO_WO = '" & Text(0) & "'"
28                If erTranAll = False Then Call objHT.UpdateData_Bengkel(sSQLHO)
                  Dim keterangan As String
29                keterangan = ""
30                If rsTemp("NOFAK") <> "" Then
                      ' ORDER DARI BOOKING
31                    Call objHT.UpdateData_Bengkel("UPDATE BOOK SET DELIV = 'BOK' WHERE NOFAK = '" & rsTemp("NOFAK") & "'")
32                    Call objHT.UpdateData_Bengkel("UPDATE STALL SET KETERANGAN = LEFT(KETERANGAN, CHARINDEX('|', KETERANGAN)) + '" & Trim(rsTemp("NOFAK")) & "' + RIGHT(RTRIM(KETERANGAN), LEN(KETERANGAN) - (CHARINDEX('|',KETERANGAN)+ " & Len(Trim(Text(0))) & ")), FLAG = '' WHERE KETERANGAN LIKE '%|" & Text(0) & "|%' AND TANGGAL = '" & Format(TglYYYYMMDD(Format(isnow, "DD/MM/YYYY")), "MM/DD/YYYY") & "'")
33                    Call removeNextDayBooking(Trim(rsTemp("NOFAK")), False)
34                    keterangan = "KETERANGAN LIKE '%|" & Text(0) & "|%' AND TANGGAL = '" & Format(TglYYYYMMDD(Format(isnow, "DD/MM/YYYY")), "MM/DD/YYYY") & "'"
35                Else
                      ' ORDER NON BOOKING
36                    Call objHT.UpdateData_Bengkel("UPDATE STALL SET FLAG = 'X' WHERE KETERANGAN LIKE '%|" & Text(0) & "|%' ") 'AND TANGGAL = '" & Format(TglYYYYMMDD(Format(isnow, "DD/MM/YYYY")), "MM/DD/YYYY") & "'")
37                    Call createMekanik
38                    keterangan = "KETERANGAN LIKE '%|" & Text(0) & "|%'"
                      'Call decresRestTime(rsTemp("MECH"), CDbl(nTotAct), 1, Format(rsTemp("DATE"), "MM/DD/YYYY"))
39                End If
40                If Log.helpJpcb = 1 Then
                      Dim rsHlpJPCB As ADODB.Recordset
                      Dim rsLogJPCB As ADODB.Recordset
                      'Dim rsTemp As ADODB.Recordset
41                    Set rsTempStall = objHT.FetchDataBengkel("SELECT * FROM STALL WHERE " & keterangan)
42                    If Not rsTempStall.EOF Then
43                        Set rsLogJPCB = objHT.FetchDataBengkel("SELECT * FROM PROFILBENGKEL.DBO.LOGJPCB")
44                        Set rsHlpJPCB = objHT.FetchDataBengkel("SELECT TOP 1 * FROM HELPJPCB WHERE " & keterangan)
45                        If rsHlpJPCB.EOF Then
46                             If rsLogJPCB.RecordCount > 0 Then
47                                rsLogJPCB.MoveFirst
48                                Do While Not rsLogJPCB.EOF
                                      'lakukan pengisian data help jpcb berdasarkan mech
49                                   Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(Text(0)) & "','T','S','" & UCase(Trim(rsTempStall("nostall"))) & "','" & Format(rsTempStall("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsTempStall("JAMMULAI")) & "','" & Trim(rsTempStall("JAMSELESAI")) & "'," & Replace(rsTempStall("rate"), ",", ".") & ",'" & Trim(rsTempStall("mech")) & "','','" & Trim(rsTempStall!keterangan) & "','" & Trim(rsTempStall("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
50                                   rsLogJPCB.MoveNext
51                                Loop
52                             End If
53                         Else
54                            If rsLogJPCB.RecordCount > 0 Then
55                                rsLogJPCB.MoveFirst
56                                Do While Not rsLogJPCB.EOF
                                      'lakukan pengisian data help jpcb berdasarkan mech
57                                   Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(Text(0)) & "','T','S','" & UCase(Trim(rsTempStall("nostall"))) & "','" & Format(rsTempStall("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsTempStall("JAMMULAI")) & "','" & Trim(rsTempStall("JAMSELESAI")) & "'," & Replace(rsTempStall("rate"), ",", ".") & ",'" & Trim(rsTempStall("mech")) & "','','" & Trim(rsTempStall!keterangan) & "','" & Trim(rsTempStall("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
58                                   rsLogJPCB.MoveNext
59                                Loop
60                             End If
          '                    rsHlpJPCB.MoveFirst
          '                    If rsHlpJPCB("STATUS") = "F" Or rsHlpJPCB!tipeUpdate = "DELETE" Then
          '                        If rsLogJPCB.RecordCount > 0 Then
          '                            rsLogJPCB.MoveFirst
          '                            Do While Not rsLogJPCB.EOF
          '                               Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry]) VALUES ('" & Trim(text(0)) & "','T','S','" & UCase(Trim(rsTempStall("nostall"))) & "','" & Format(rsTempStall("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsTempStall("JAMMULAI")) & "','" & Trim(rsTempStall("JAMSELESAI")) & "'," & Replace(rsTempStall("rate"), ",", ".") & ",'" & Trim(rsTempStall("mech")) & "','','" & Trim(rsTempStall!keterangan) & "','" & Trim(rsTempStall("adjust")) & "','DELETE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE())")
          '                               rsLogJPCB.MoveNext
          '                            Loop
          '                        End If
          '                    End If
61                        End If
62                    End If
63                End If
                  
64                Call objHT.InsertData_Bengkel("INSERT INTO NOTEPART (NOSO, TIPE) VALUES ('" & Trim(Text(0)) & "', 'BO')")
                  
65                Call objHT.DeleteData_Bengkel("DELETE TEMPJOBSUBLET WHERE NO_CHASSIS = '" & Text(15) & "'")     'NO_WO = '" & Text(64) & "' AND
66                Call objHT.DeleteData_Bengkel("DELETE TEMPFU WHERE NO_CHASSIS = '" & Text(15) & "'")   'NO_WO = '" & Text(64) & "' AND
67                Call objHT.DeleteData_Bengkel("DELETE TEMPSTALL WHERE KETERANGAN LIKE '" & Text(14) & " |%'")
                  
                  ' LEPASKAN PREDIAG DARI WO
68                Call objHT.UpdateData_Bengkel("UPDATE DIAGDTA SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
69                Call objHT.UpdateData_Bengkel("UPDATE DIAGDTL SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
70                Call objHT.UpdateData_Bengkel("UPDATE DIAGKOORD SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
71                Call objHT.UpdateData_Bengkel("UPDATE DIAGLENGKAP SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
72                Call objHT.UpdateData_Bengkel("UPDATE DIAGPART SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
73                Call objHT.UpdateData_Bengkel("UPDATE DIAGSUBLET SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
74                Call objHT.UpdateData_Bengkel("UPDATE DIAGAPPROVAL SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
75                Call objHT.UpdateData_Bengkel("UPDATE DIAGADD2 SET NO_WO = '' WHERE NO_WO = '" & Trim(Text(0)) & "' ")
76                Call objHT.Update_LogAR_BatalWO(Trim(Text(0)))
77                If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
78                    If rsTemp("NOFAK") = "" Then
79                        Call objHT.UpdateData_ACC("UPDATE SOHDR SET TUTUPORDER = 1 WHERE NOSO = '" & Trim(Text(0)) & "' ")
80                    Else
81                        Call objHT.UpdateData_ACC("UPDATE SOHDR SET NOSO = NOSPK WHERE NOSO = '" & Trim(Text(0)) & "' ")
82                        Call objHT.UpdateData_ACC("UPDATE SOORD SET NOSO = NOSPK WHERE NOSO = '" & Trim(Text(0)) & "' ")
83                    End If
84                End If
85            End If
              ''-----------------------------------------------------
              ''20180214 replikasi untuk batal order
              ''-----------------------------------------------------
86            If Log.RPLHistori = 1 Then
87                sSQL = sSQLHO & " And FS_KD_Dealer = '" & gKdCbg & "'"
88                sSQL = Replace(sSQL, "JOBDTA", "[" & dbReplikasiHO & "]..JOBDTA")
89                Call objHT.UpdateData_Bengkel(sSQL)
                  
90            Else
91                SQLHO = sSQLHO & " And FS_KD_Dealer = '" & gKdCbg & "'"
92                If Log.LogHOConn = 1 Then
93                    If ConHO = False Then
94                        Call Saving_Log_HO("Delete", sSQLHO)
95                    Else
96                        If Get_ConnectionProfil_Bengkel99 = False Then
97                            Call Saving_Log_HO("Delete", sSQLHO)
98                        Else
99                            Call objHO.OpenTransactionLog_NRM
100                           Call objHO.AddNew_WO(sSQLHO)
101                           Call objHO.CloseTransactionLog_NRM
102                           If errHO = True Then Call Saving_Log_HO("Delete", sSQLHO)
103                       End If
104                   End If
105               Else
106                   Call Saving_Log_HO("Delete", sSQLHO)
107               End If
108           End If
                  
                          
109               If erTranAll = False Then
110                   Call objHT.CloseTransactionAll
111                   onTR = False
112                   Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlOrd", "cmdOk", 1)
113               Else
114                   Call objHT.CloseTransactionAll_Err
115                   onTR = False
116                   Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlOrd", "cmdOk", 2)
117                   MsgBox "Order tidak berhasil dihapus !!" & vbCr & "Ulangi proses Penghapusan....", vbCritical, "Hapus Order"
118                   Exit Sub
119               End If
120               If Not rsTemp.EOF Then Call cetakPassBatal(Trim(Text(0)))
121       Else
122           MsgBox "Isi Alasan Pembatalan !!! "
123           Text1(0).SetFocus
124           Exit Sub
125       End If
126       frmBtlOrder.Visible = False
127       Call hapusForm
128       Exit Sub

vbError:
129       Call objHT.CloseTransactionAll_Err
130       If erTranAll = True Then
              
131           Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlOrd", "cmdOk", 2)
132           MsgBox "Order tidak berhasil dihapus !!" & vbCr & "Ulangi proses Penghapusan....", vbCritical, "Hapus Order"
133       End If
134       Exit Sub
End Sub

Private Sub Saving_Log_HO(nStatus As String, isSQl As String)
    If isSQl <> "" Then
        Call objHO.AddNew_LogHO(Trim(Text(0)), nStatus, "WO", isSQl)
    End If
End Sub

Private Sub cetakPassBatal(nWO As String)
    Dim Report As New RptGatePassBatal
    Dim rsRepo As New ADODB.Recordset
    
    If rsRepo.State Then rsRepo.Close
    sSQL = "SELECT NO_WO, NOINV, NO_POLISI, NO_CHASSIS, MDL, FMDL, CUSTOMER, CUST_ADRES, ALASAN FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Text(0) & "' AND DELIV = 'BTL'"
    Set rsRepo = objO.FetchData(sSQL)
    If Not rsRepo.EOF Then
        Report.database.SetDataSource rsRepo
        Report.txtKabeng.SetText gSecmanager
    
        CRViewer1.ReportSource = Report
        CRViewer1.ViewReport
        CRViewer1.Zoom 100
        CRViewer1.Visible = True
        cmdTutup.Visible = True
        Screen.MousePointer = vbDefault
    Else
        MsgBox "Data belum disimpan"
    End If
    'rsRepo.Close
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    CRViewer1.Width = 0
    CRViewer1.Height = 0
    cmdTutup.Visible = False
    Call hapusForm
End Sub

Private Sub Form_Load()
    Call hapusForm
    SSTab1.Tab = 0
End Sub

Private Sub Form_Resize()
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub tblBatal_Click()
Dim rsSearch As New ADODB.Recordset

    If bisaBatal Then
        If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
            Dim rsPSO As ADODB.Recordset
            'Set rsPSO = objO.FetchDataACC("Select NOSO FROM SOORD Where NoSO = '" & Trim(Text(64)) & "' and ((IsNull(QtyKirimBL,0) <> 0 OR IsNull(QtyKirimBI,0) <> 0 OR [TO] <> 0) Or BO not in ('','CANCEL') )")
            Set rsPSO = objO.FetchDataACC("Select NOSO FROM SOORD WITH(NOLOCK) Where NoSO = '" & Trim(Text(64)) & "' and ((IsNull(QtyKirimBL,0) <> 0 OR IsNull(QtyKirimBI,0) <> 0 OR [TO] <> 0) Or BO not in ('','CANCEL') ) and ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0 )")
            If rsPSO.RecordCount > 0 Then
                MsgBox "Part sudah di proses oleh partman !," & Chr(13) & "Work Order tidak dapat di batalkan" & Chr(13) & Chr(13) & "Lakukan pembatalan proses part oleh Partman, untuk membatalkan Work Order"
                Exit Sub
            End If
        End If
            'Set rsSearch = objO.FetchDataACC("SELECT ISNULL(SUM(qtysaldoawal+qtyterimabln+qtykeluarbln+abs(qtyjualbln)),0) AS JML FROM INMST WHERE NORACK = '" & Trim(text(0)) & "' AND (qtysaldoawal+qtyterimabln+qtykeluarbln+abs(qtyjualbln)) > 0")
            If AdaPemakaianPart(Trim(Text(0))) Then
            'If Not rsSearch.EOF Then
                'If CDbl(rsSearch("JML")) > 0 Then
                    MsgBox "Masih ada Penggunaan part. Pemakaian Part dibatalkan dahulu. "
                    Exit Sub
                'Else
                '    frmBtlOrder.Visible = True
                'End If
            Else
                If AdaPemakaianSublet(Trim(Text(0))) Then
                    MsgBox "Masih ada Penggunaan Sublet. Validasi Sublet diretur dahulu. "
                    Exit Sub
                Else
                    frmBtlOrder.Visible = True
                End If
            End If
    Else
        MsgBox pesan & vbCr & vbCr & "Tentukan Order yang akan dibatalkan !", vbCritical, "Pembatalan Order"
        Text(0).SelLength = Len(Text(0))
        Text(0).SetFocus
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
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
                gSearchOption = "ORDERBATAL"
                frmCari.show (vbModal)
        End Select
    End If
    If KeyCode = vbKeyEscape Then
        Call hapusForm
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL' AND NOINV = '' AND NO_WO = '" & Trim(Text(Index)) & "' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '') ")
                If rsSearch.EOF Then
                    pesan = "WO belum ada atau sudah divalidasi. "
                    bisaBatal = False
                    gMainKeyCari = Text(Index)
                    gSearchOption = "ORDERBATAL"
                    frmCari.show (vbModal)
                Else
                    If Not AdaInvoicePart(Trim(Text(Index))) Then
                        If Not AdaPemakaianPart(Trim(Text(Index))) Then
                            If Not AdaPemakaianSublet(Trim(Text(Index))) Then
                                Call PanggilDataOrder(rsSearch("NO_CHASSIS"), Trim(Text(Index)))
                                Call panggilHistory(rsSearch("NO_CHASSIS"), Trim(Text(14)))
                                bisaBatal = True
                            Else
                                pesan = "Masih ada Pemakaian Sublet !!" & vbCr & "Tidak bisa dibatalkan .."
                                bisaBatal = False
                            End If
                        Else
                            pesan = "Masih ada Pemakaian Part !!" & vbCr & "Tidak bisa dibatalkan .."
                            bisaBatal = False
                        End If
                    Else
                        pesan = "Masih ada Invoice Part !!" & vbCr & "Tidak bisa dibatalkan .."
                        bisaBatal = False
                    End If
                End If
            End If
    End Select
    
    If Trim(pesan) <> "" Then
        MsgBox pesan, vbCritical
        pesan = ""
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
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
