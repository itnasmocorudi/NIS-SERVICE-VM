VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmRepoJmlInv 
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   12945
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
      Height          =   270
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   4800
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
      Height          =   270
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   14
      Top             =   4440
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
   Begin VB.CommandButton tblRefresh 
      Caption         =   "&Refresh"
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
      Left            =   6840
      TabIndex        =   41
      Top             =   3840
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   480
      TabIndex        =   15
      Top             =   480
      Width           =   12375
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   68026369
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2400
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
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
         Index           =   2
         Left            =   1530
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   5400
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
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
         Index           =   3
         Left            =   1530
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
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
         Index           =   1
         Left            =   2520
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1530
         TabIndex        =   5
         Top             =   1920
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   2400
         TabIndex        =   21
         Top             =   1920
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
         Index           =   7
         Left            =   1530
         TabIndex        =   6
         Top             =   2280
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
         Left            =   3480
         TabIndex        =   22
         Top             =   2280
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
         Index           =   6
         Left            =   2520
         TabIndex        =   23
         Top             =   1920
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   1530
         TabIndex        =   24
         Top             =   2640
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
         Index           =   9
         Left            =   7290
         TabIndex        =   7
         Top             =   840
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   8160
         TabIndex        =   25
         Top             =   847
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
         Index           =   10
         Left            =   8280
         TabIndex        =   26
         Top             =   840
         Width           =   3195
         _ExtentX        =   5636
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   7290
         TabIndex        =   8
         Top             =   1200
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   9135
         TabIndex        =   27
         Top             =   1207
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
         Index           =   12
         Left            =   7290
         TabIndex        =   9
         Top             =   1560
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   9135
         TabIndex        =   28
         Top             =   1567
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
         Index           =   13
         Left            =   7290
         TabIndex        =   10
         Top             =   1920
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   9135
         TabIndex        =   29
         Top             =   1927
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   68026369
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   7290
         TabIndex        =   42
         Top             =   2280
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox10 
         Height          =   270
         Left            =   9135
         TabIndex        =   43
         Top             =   2280
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
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
         Left            =   6735
         TabIndex        =   44
         Top             =   2325
         Width           =   465
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Job Order"
         BeginProperty Font 
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
         TabIndex        =   40
         Top             =   885
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   39
         Top             =   1245
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mechanic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   780
         TabIndex        =   38
         Top             =   1605
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   1245
         TabIndex        =   37
         Top             =   1965
         Width           =   195
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan"
         BeginProperty Font 
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
         TabIndex        =   36
         Top             =   2325
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Penjual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6165
         TabIndex        =   35
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Service"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6330
         TabIndex        =   34
         Top             =   1245
         Width           =   870
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Order"
         BeginProperty Font 
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
         TabIndex        =   33
         Top             =   1605
         Width           =   765
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proses"
         BeginProperty Font 
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
         TabIndex        =   32
         Top             =   1965
         Width           =   480
      End
      Begin VB.Label Label10 
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
         Left            =   870
         TabIndex        =   31
         Top             =   240
         Width           =   570
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
         Left            =   3240
         TabIndex        =   30
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   12360
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
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
      Left            =   10920
      TabIndex        =   12
      Top             =   3840
      Width           =   1890
   End
   Begin VB.CommandButton TblSimpan 
      Caption         =   "&Cetak"
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
      Left            =   9000
      TabIndex        =   11
      Top             =   3840
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Laporan Jumlah Invoice Masuk"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
End
Attribute VB_Name = "frmRepoJmlInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : -
'=======================================

Dim Report As Object
Dim rsTemp As New ADODB.Recordset
Dim vNamaTabel As String
Dim objBilling As New BClsBilling

Private Sub hapusForm()
    For i = 0 To 14
        text(i) = ""
    Next
    tgl1 = gTglServer
    tgl2 = gTglServer
End Sub

Private Sub BuatTabelRepoInv(vNamaTabel As String)
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
         & " drop table [dbo].[" & vNamaTabel & "]"
    oConn.Execute sSQL
    
    sSQL = "CREATE TABLE [dbo].[" & vNamaTabel & "] (" _
         & "    [SA][char](5) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [JOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [JRTE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PDI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [WAS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [BOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [MATERIAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PART][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [OLI][decimal](18,2) NOT NULL DEFAULT 0, "
         
    sSQL = sSQL + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
         & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
         & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
         & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000') ON [PRIMARY]"
    oConn.Execute sSQL
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As New ADODB.Recordset

    createTemp = False
    If rsHlp.State Then rsHlp.Close
    Set rsHlp = oConn.Execute("select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepoInv(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(ket As String, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
Dim hlp6 As String
Dim hlp7 As String
Dim rsHlp As New ADODB.Recordset
Dim rsPro As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim nNilai As Double
Dim nDuoInv As Boolean
Dim subletOnly As Boolean
Dim noinv As String
Dim nmCust As String
Dim nStatus As String
'Dim vNamaTabel As  String
    
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                             & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                             & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                             & "FROM T_INVOICE A INNER JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                             & "WHERE DELIV <> 'BTL' AND INVDATE BETWEEN '" & Format(tgl1, "MM/DD/YYYY") & "' AND '" & Format(tgl2, "MM/DD/YYYY") & "' " & ket & " ORDER BY NOINV")
    If Not rsTemp.EOF Then
        Call BuatTabelRepoInv(vNamaTabel)
        Do While Not rsTemp.EOF
            subletOnly = True
            For i = 1 To 8
                hlp = "REP" & i
                If rsTemp(hlp) <> "" Then
                    subletOnly = False
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "RTE" & i    '"ACTRATE" & i
                    hlp4 = "MEKANIK" & i
                    hlp5 = "OFF" & i
                    hlp6 = "DISC" & i
                    hlp7 = "FREEC" & i
                    If createTemp(vNamaTabel) Then
                        nDuoInv = False
                        If rsTemp("NDUOINV") = 1 Then
                            nDuoInv = True
                        End If
                        If nDuoInv Then
                            If rsTemp(hlp7) = 0 Then
                                nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
                                nNilai = CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And Trim(rsTemp(hlp2)) = "SBE") Or (Trim(rsTemp(hlp2) = "SBI")), 0, CDbl(rsTemp(hlp6)) / 100))
                                noinv = Trim(rsTemp("NOINV")) & "-B"
                                nmCust = Trim(rsTemp("NMPENJUAL"))
                            Else
                                nStatus = rsTemp("STATUS")
                                nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
                                noinv = Trim(rsTemp("NOINV")) & "-A"
                                nmCust = Trim(rsTemp("CUSTOMER"))
                            End If
                        Else
                            If rsTemp("FS_RC") = "R001" Then
                                nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
                                nNilai = CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And Trim(rsTemp(hlp2)) = "SBE") Or (Trim(rsTemp(hlp2) = "SBI")), 0, CDbl(rsTemp(hlp6)) / 100))
                                noinv = Trim(rsTemp("NOINV"))
                                nmCust = Trim(rsTemp("NMPENJUAL"))
                            Else
                                nStatus = rsTemp("STATUS")
                                nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
                                noinv = Trim(rsTemp("NOINV"))
                                nmCust = Trim(rsTemp("CUSTOMER"))
                            End If
                        End If
                        
                        If rsHlp.State Then rsHlp.Close
                        Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                        If Not rsHlp.EOF Then
                            oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "BOR = '" & cekDuo(CDbl(rsTemp("BODY")), noinv) & "', MATERIAL = '" & cekDuo(CDbl(rsTemp("MTRL")), noinv) & "', PART = '" & cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noinv) & "', " _
                                         & "OLI = '" & cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noinv) & "', SUBLET = '" & cekDuo(CDbl(rsTemp("SBLT")), noinv) & "', SUBTOT = '" & cekDuo(CDbl(rsTemp("SUBTOT")), noinv) & "', " _
                                         & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & cekDuo(CDbl(rsTemp("HPP")), noinv) & "' " _
                                         & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                        Else
                            oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noinv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                         & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                         & "'" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "'" & cekDuo(CDbl(rsTemp("BODY")), noinv) & "', '" & cekDuo(CDbl(rsTemp("MTRL")), noinv) & "', '" & cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noinv) & "', '" & cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noinv) & "', '" & cekDuo(CDbl(rsTemp("SBLT")), noinv) & "', '" & cekDuo(CDbl(rsTemp("SUBTOT")), noinv) & "', " _
                                         & " '" & cekDuo(CDbl(rsTemp("MET")), noinv, 1) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', '" & cekDuo(CDbl(rsTemp("HPP")), noinv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                        End If
                        If rsHlp.State <> adStateClosed Then rsHlp.Close
                        
                        If nDuoInv Then
                            nStatus = rsTemp("STATUS")
                            nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100))
                            noinv = Trim(rsTemp("NOINV")) & "-A"
                            nmCust = Trim(rsTemp("CUSTOMER"))
                            If rsHlp.State Then rsHlp.Close
                            Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NOINV = '" & Trim(noinv) & "'")
                            If Not rsHlp.EOF Then
                                oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                             & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                             & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP", setKoma(CDbl(nNilai)), 0) & "', " _
                                             & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                             & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                             & "BOR = '" & CDbl(rsTemp("BODY")) & "', MATERIAL = '" & CDbl(rsTemp("MTRL")) & "', PART = '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', " _
                                             & "OLI = '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', SUBLET = '" & CDbl(rsTemp("SBLT")) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                             & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & CDbl(rsTemp("HPP")) & "' " _
                                             & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                            Else
                                oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noinv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                             & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                             & "'" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                             & "'" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                             & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                             & " '" & cekDuo(CDbl(rsTemp("MET")), noinv, 2) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "', '" & CDbl(rsTemp("HPP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                            End If
                            If rsHlp.State <> adStateClosed Then rsHlp.Close
                        End If
                    End If
                End If
            Next
            If subletOnly Then
                If rsHlp.State Then rsHlp.Close
                Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                If Not rsHlp.EOF Then
                    oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                 & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & CDbl(rsTemp("BODY")) & "', MATERIAL = '" & CDbl(rsTemp("MTRL")) & "', " _
                                 & "PART = '" & CDbl(rsTemp("PART_G") + rsTemp("PART_NG")) & "', OLI = '" & CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG")) & "', SUBLET = '" & CDbl(rsTemp("SBLT")) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                 & "LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & CDbl(rsTemp("HPP")) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                Else
                    oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                 & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                 & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                 & " '" & CDbl(rsTemp("MET")) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(rsTemp("STATUS")) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & CDbl(rsTemp("HPP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                End If
            End If
            
            ' DARI ADD JOB
            If rsAdd.State Then rsAdd.Close
            Set rsAdd = oConn.Execute("SELECT * FROM JOBADD2 WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
            If Not rsAdd.EOF Then
                Do While Not rsAdd.EOF
                    hlp2 = "KLPKERJA"
                    hlp3 = "RTE"
                    hlp4 = "MEKANIK"
                    hlp5 = "OFFN"
                    hlp6 = "FREEC"
                    
                    nDuoInv = False
                    If rsTemp("NDUOINV") = 1 Then
                        nDuoInv = True
                    End If
                    
                    If nDuoInv Then
                        If rsAdd(hlp6) = 0 Then
                            nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
                            nNilai = CDbl(rsAdd(hlp5)) * (1 - IIf((nDuoInv And rsAdd(hlp2) = "SBE"), 0, CDbl(rsAdd("DISC")) / 100))
                            noinv = Trim(rsTemp("NOINV")) & "-B"
                            nmCust = Trim(rsTemp("NMPENJUAL"))
                        Else
                            nStatus = rsTemp("STATUS")
                            nNilai = CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100))
                            noinv = Trim(rsTemp("NOINV")) & "-A"
                            nmCust = Trim(rsTemp("CUSTOMER"))
                        End If
                    Else
                        nStatus = rsTemp("STATUS")
                        nNilai = CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100))
                        noinv = Trim(rsTemp("NOINV"))
                        nmCust = Trim(rsTemp("CUSTOMER"))
                    End If
                    
                    
                    If rsHlp.State Then rsHlp.Close
                    Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                    If Not rsHlp.EOF Then
                        oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                     & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                     & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP", setKoma(CDbl(nNilai)), 0) & "', " _
                                     & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                     & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                     & "BOR = '" & cekDuo(CDbl(rsTemp("BODY")), noinv) & "', MATERIAL = '" & cekDuo(CDbl(rsTemp("MTRL")), noinv) & "', PART = '" & cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noinv) & "', " _
                                     & "OLI = '" & cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noinv) & "', SUBLET = '" & cekDuo(CDbl(rsTemp("SBLT")), noinv) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "', HPP = '" & cekDuo(CDbl(rsTemp("HPP")), noinv) & "', " _
                                     & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "' " _
                                     & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                    Else
                        oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                     & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                     & "'" & cekDuo(CDbl(rsTemp("BODY")), noinv) & "', '" & cekDuo(CDbl(rsTemp("MTRL")), noinv) & "', '" & cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noinv) & "', '" & cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noinv) & "', '" & cekDuo(CDbl(rsTemp("SBLT")), noinv) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                     & " '" & cekDuo(CDbl(rsTemp("MET")), noinv, 1) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "','" & cekDuo(CDbl(rsTemp("HPP")), noinv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                    End If
                    If rsHlp.State <> adStateClosed Then rsHlp.Close
                    
                    If nDuoInv Then
                        nStatus = rsTemp("STATUS")
                        nNilai = CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100))
                        noinv = Trim(rsTemp("NOINV")) & "-A"
                        nmCust = Trim(rsTemp("CUSTOMER"))
                        If rsHlp.State Then rsHlp.Close
                        Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                        If Not rsHlp.EOF Then
                            oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                         & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                         & "BOR = '" & CDbl(rsTemp("BODY")) & "', MATERIAL = '" & CDbl(rsTemp("MTRL")) & "', PART = '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', " _
                                         & "OLI = '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', SUBLET = '" & CDbl(rsTemp("SBLT")) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "', HPP = '" & CDbl(rsTemp("HPP")) & "', " _
                                         & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "' " _
                                         & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                        Else
                            oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                         & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                         & "'" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                         & "'" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                         & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                         & " '" & cekDuo(CDbl(rsTemp("MET")), noinv, 2) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & CDbl(rsTemp("HPP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "')")
                        End If
                        If rsHlp.State <> adStateClosed Then rsHlp.Close
                    End If
                    rsAdd.MoveNext
                Loop
            End If
            If rsAdd.State <> adStateClosed Then rsAdd.Close
            rsTemp.MoveNext
        Loop
        Call Cetak(vNamaTabel, txt)
    Else
        MsgBox "Tidak ada data !!"
        Exit Sub
    End If
End Sub

Private Function statusDealer(codeStatus As Integer)
    
    Select Case codeStatus
        Case 1
            statusDealer = "AFFNONPPN"
        Case 2
            statusDealer = "AFFPPN"
        Case 3
            statusDealer = "EXT"
    End Select
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 2 Then
            cekDuo = 0
        End If
    End If
End Function

Private Sub IsiTempRepo2(ket As String, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
Dim hlp6 As String
Dim rsHlp As New ADODB.Recordset
Dim rsPro As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim nNilai As Double
Dim nDuoInv As Boolean
'Dim vNamaTabel As  String
    
    If rsTemp.State Then rsTemp.Close
    'A LEFT JOIN STALL B ON REPLACE(SUBSTRING(B.KETERANGAN, CHARINDEX('|',B.KETERANGAN)+1,8),'|','') = A.NO_WO AND B.TANGGAL = A.[DATE]
    Set rsTemp = oConn.Execute("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, * FROM T_INVOICE " _
                             & "WHERE DELIV <> 'BTL' AND INVDATE BETWEEN '" & Format(tgl1, "MM/DD/YYYY") & "' AND '" & Format(tgl2, "MM/DD/YYYY") & "' " & ket & " ORDER BY NOINV")
    If Not rsTemp.EOF Then
'        Call BuatTabelRepoInv(vNamaTabel)
'        Do While Not rsTemp.EOF
'            For i = 1 To 8
'                hlp = "REP" & i
'                If rsTemp(hlp) <> "" Then
'                    hlp2 = "KLPKERJA" & i
'                    hlp3 = "RTE" & i    '"ACTRATE" & i
'                    hlp4 = "MEKANIK" & i
'                    hlp5 = "OFF" & i
'                    If createTemp(vNamaTabel) Then
'                        nDuoInv = False
'                        If Trim(rsTemp("TPSVC")) = "P" Then
'                            If rsPro.State Then rsPro.Close
'                            Set rsPro = oConn.Execute("SELECT * FROM HO_PROGRAM WHERE ID_PROGRAM = '" & Trim(rsTemp("KODETIPE")) & "'")
'                            If Not rsPro.EOF Then
'                                If Trim(rsPro("DOUBLEINV")) = 1 Then
'                                    nDuoInv = True
'                                End If
'                            End If
'                            If rsPro.State <> adStateClosed Then rsPro.Close
'                        End If
'                        nNilai = CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And rsTemp(hlp2) = "SBE"), 0, CDbl(rsTemp("DISC")) / 100))
'                        If rsHlp.State Then rsHlp.Close
'                        Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
'                        If Not rsHlp.EOF Then
'                            oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
'                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
'                                         & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP", CDbl(nNilai), 0) & "', " _
'                                         & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) & "', " _
'                                         & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) & "', TIPEINV = '" & Trim(rsTemp("STATUS")) & "' " _
'                                         & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
'                        Else
'                            oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
'                                         & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
'                                         & "'" & IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsTemp(hlp2) = "GRP", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
'                                         & "'" & IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
'                                         & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
'                                         & " '" & CDbl(rsTemp("MET")) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(rsTemp("STATUS")) & "')")
'                        End If
'                    End If
'                End If
'            Next
'            rsTemp.MoveNext
'        Loop
        Call BuatTabelRepoInv(vNamaTabel)
        Do While Not rsTemp.EOF
            subletOnly = True
            For i = 1 To 8
                hlp = "REP" & i
                If rsTemp(hlp) <> "" Then
                    subletOnly = False
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "RTE" & i    '"ACTRATE" & i
                    hlp4 = "MEKANIK" & i
                    hlp5 = "OFF" & i
                    hlp6 = "DISC" & i
                    If createTemp(vNamaTabel) Then
                        nDuoInv = False
                        If Trim(rsTemp("TPSVC")) = "P" Then
                            If rsPro.State Then rsPro.Close
                            Set rsPro = oConn.Execute("SELECT * FROM HO_PROGRAM WHERE ID_PROGRAM = '" & Trim(rsTemp("KODETIPE")) & "'")
                            If Not rsPro.EOF Then
                                If Trim(rsPro("DOUBLEINV")) = 1 Then
                                    nDuoInv = True
                                End If
                            End If
                            If rsPro.State <> adStateClosed Then rsPro.Close
                        End If
                        nNilai = Replace(CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And rsTemp(hlp2) = "SBE") Or (Trim(rsTemp(hlp2) = "SBI")), 0, CDbl(rsTemp(hlp6)) / 100)), ",", ".")
                        'nNilai = CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp("DISC")) / 100))
                        If rsHlp.State Then rsHlp.Close
                        Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                        If Not rsHlp.EOF Then
                            oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                         & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP", CDbl(nNilai), 0) & "', " _
                                         & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) & "', " _
                                         & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) & "', " _
                                         & "BOR = '" & CDbl(rsTemp("BODY")) & "', MATERIAL = '" & CDbl(rsTemp("MTRL")) & "', PART = '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', " _
                                         & "OLI = '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', SUBLET = '" & CDbl(rsTemp("SBLT")) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                         & "TIPEINV = '" & Trim(rsTemp("STATUS")) & "'" _
                                         & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                        Else
                            oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                         & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                         & "'" & IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsTemp(hlp2) = "GRP", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                         & "'" & IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                         & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                         & " '" & CDbl(rsTemp("MET")) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(rsTemp("STATUS")) & "')")
                        End If
                    End If
                End If
            Next
            If subletOnly Then
                If rsHlp.State Then rsHlp.Close
                Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                If Not rsHlp.EOF Then
                    oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                 & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & CDbl(rsTemp("BODY")) & "', MATERIAL = '" & CDbl(rsTemp("MTRL")) & "', " _
                                 & "PART = '" & CDbl(rsTemp("PART_G") + rsTemp("PART_NG")) & "', OLI = '" & CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG")) & "', SUBLET = '" & CDbl(rsTemp("SBLT")) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "'WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                Else
                    oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                 & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                 & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                 & " '" & CDbl(rsTemp("MET")) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(rsTemp("STATUS")) & "')")
                End If
            End If
            
            ' DARI ADD JOB
            If rsAdd.State Then rsAdd.Close
            Set rsAdd = oConn.Execute("SELECT * FROM JOBADD2 WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
            If Not rsAdd.EOF Then
                Do While Not rsAdd.EOF
                    hlp2 = "KLPKERJA"
                    hlp3 = "RTE"
                    hlp4 = "MEKANIK"
                    hlp5 = "OFFN"
                    
                    nDuoInv = False
                    If Trim(rsTemp("TPSVC")) = "P" Then
                        If rsPro.State Then rsPro.Close
                        Set rsPro = oConn.Execute("SELECT * FROM HO_PROGRAM WHERE ID_PROGRAM = '" & Trim(rsTemp("KODETIPE")) & "'")
                        If Not rsPro.EOF Then
                            If Trim(rsPro("DOUBLEINV")) = 1 Then
                                nDuoInv = True
                            End If
                        End If
                        If rsPro.State <> adStateClosed Then rsPro.Close
                    End If
                    nNilai = Replace(CDbl(rsAdd(hlp5)) * (1 - IIf((nDuoInv And rsAdd(hlp2) = "SBE") Or (Trim(rsAdd(hlp2) = "SBI")), 0, CDbl(rsAdd("DISC")) / 100)), ",", ".")
                    If rsHlp.State Then rsHlp.Close
                    Set rsHlp = oConn.Execute("SELECT * FROM [" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                    If Not rsHlp.EOF Then
                        oConn.Execute ("UPDATE [" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & rsTemp("NOINV") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                     & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                     & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP", CDbl(nNilai), 0) & "', " _
                                     & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', " _
                                     & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "', " _
                                     & "BOR = '" & CDbl(rsTemp("BODY")) & "', MATERIAL = '" & CDbl(rsTemp("MTRL")) & "', PART = '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', " _
                                     & "OLI = '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', SUBLET = '" & CDbl(rsTemp("SBLT")) & "', SUBTOT = '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                     & "TIPEINV = '" & Trim(rsTemp("STATUS")) & "' " _
                                     & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "'")
                    Else
                        oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & rsTemp("NOINV") & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                     & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                     & "'" & CDbl(rsTemp("BODY")) & "', '" & CDbl(rsTemp("MTRL")) & "', '" & CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")) & "', '" & CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")) & "', '" & CDbl(rsTemp("SBLT")) & "', '" & CDbl(rsTemp("SUBTOT")) & "', " _
                                     & " '" & CDbl(rsTemp("MET")) & "', '" & CDbl(rsTemp("TOTALL")) & "', '" & Trim(rsTemp("STATUS")) & "')")
                    End If
                    rsAdd.MoveNext
                Loop
            End If
            If rsAdd.State <> adStateClosed Then rsAdd.Close
            rsTemp.MoveNext
        Loop
        Call Cetak(vNamaTabel, txt)
    Else
        MsgBox "Tidak ada data !!"
        Exit Sub
    End If
End Sub

Private Sub Cetak(vNamaTabel As String, txt As String)
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptRepoInvSA
'    If txt <> "" Then
'        txt = " WHERE " & txt
'    End If
    sSQL = "SELECT * FROM [" & vNamaTabel & "] ORDER BY NOINV"     '" & txt & "
    Set oRs = oConn.Execute(sSQL)
    Report.Database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(tgl1, "DD-MM-YYYY") & " s/d " & Format(tgl2, "DD-MM-YYYY")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    sSQL = "DELETE [" & vNamaTabel & "]"
    oConn.Execute sSQL
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    Call hapusForm
    vNamaTabel = "JmlInv" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
         & " drop table [dbo].[" & vNamaTabel & "]"
    oConn.Execute sSQL
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub TblSimpan_Click()
Dim boleh As Boolean
Dim txtTemp As String
Dim txt As String
    
    boleh = False
    txtTemp = ""
    txt = ""
    If text(0) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "KLPKERJA = '" & text(0) & "' "
    End If
    If text(2) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "REPS = '" & text(2) & "' "
    End If
    If text(3) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "MECHS = '" & text(3) & "' "
    End If
    If text(5) <> "" Then
        txtTemp = txtTemp & "AND FRONT = '" & text(5) & "' "
    End If
    If text(7) <> "" Then
        txtTemp = txtTemp & "AND MDL = '" & text(7) & "' "
    End If
    If text(9) <> "" Then
        txtTemp = txtTemp & "AND FS_KD_DEALER = '" & text(9) & "' "
    End If
    If text(11) <> "" Then
        txtTemp = txtTemp & "AND TPSVC = '" & IIf(text(11) = "REGULAR", "", IIf(text(11) = "WARRANTY", "W", "P")) & "' "
    End If
    If text(12) <> "" Then
        If text(12) = "BOOKING" Then
            txtTemp = txtTemp & "AND NOFAK <> ''"
        ElseIf text(12) = "NON BOOKING" Then
            txtTemp = txtTemp & "AND NOFAK = ''"
        End If
    End If
    If text(13) <> "" Then
        If text(13) = "PROSES" Then
            txtTemp = txtTemp & "AND RPR_COND = 'P'"
        ElseIf text(13) = "WAITING" Then
            txtTemp = txtTemp & "AND RPR_COND = '' AND NOINV = ''"
        ElseIf text(13) = "FINISH" Then
            txtTemp = txtTemp & "AND RPR_COND = 'S'"
        ElseIf text(13) = "CLOSING" Then
            txtTemp = txtTemp & "AND RPR_COND = 'K'"
        ElseIf text(13) = "PAUSED" Then
            txtTemp = txtTemp & "AND RPR_COND = 'T'"
        End If
    End If
    If text(14) <> "" Then
        txtTemp = txtTemp & "AND STATUS = '" & Trim(text(14)) & "'"
    End If
    
    Call IsiTempRepo(txtTemp, vNamaTabel, txt)
End Sub

Private Sub text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
'    Text(Index).SelLength = Len(Text(Index))
'    Text(Index).SetFocus
End Sub

Private Sub text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.Show (vbModal)
            Case 2
                If text(0) <> "" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = text(0)
                    frmCari.Show (vbModal)
                Else
                    MsgBox "Group Job Order Harus diisi"
                    text(0).SetFocus
                    Exit Sub
                End If
            Case 3
                gMainKeyCari = text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = ""
                frmCari.Show (vbModal)
            Case 5
                gMainKeyCari = text(Index)
                gSearchOption = "SA"
                frmCari.Show (vbModal)
            Case 7
                gMainKeyCari = text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
            Case 9
                gMainKeyCari = text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.Show (vbModal)
            Case 11
                gMainKeyCari = text(Index)
                gSearchOption = "TPSVC"
                frmCari.Show (vbModal)
            Case 12
                gMainKeyCari = text(Index)
                gSearchOption = "TP_ORD"
                frmCari.Show (vbModal)
            Case 13
                gMainKeyCari = text(Index)
                gSearchOption = "PROSES"
                frmCari.Show (vbModal)
            Case 14
                gMainKeyCari = text(Index)
                gSearchOption = "STATUSAKUN"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
'    If Text(Index) <> "" Then
        Select Case Index
            Case 0
                If text(Index) <> "" Then
                    sSQL = "SELECT * FROM HO_FLATE WHERE JOB = '" & text(Index) & "' and FS_LOG_DEL = ''"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "JENIS_KERJA"
                        frmCari.Show (vbModal)
                    Else
                        text(Index + 1) = oRs("REMARK")
                    End If
                    oRs.Close
                Else
                    text(1) = ""
                End If
            Case 2
                If text(0) <> "" Then
                    If text(Index) <> "" Then
                        sSQL = "SELECT * FROM HO_FLATE WHERE JOB = '" & text(0) & "' AND FS_ITEM = '" & text(Index) & "' and FS_LOG_DEL = ''"
                        oRs.Open sSQL, oConn, adOpenStatic
                        If oRs.EOF Then
                            gMainKeyCari = text(Index)
                            gSearchOption = "JOB_ORDER"
                            gBantuan = text(0)
                            frmCari.Show (vbModal)
                        Else
                            text(Index + 1) = oRs("KETERANGAN")
                        End If
                        oRs.Close
                    Else
                        text(Index + 1) = ""
                    End If
                Else
                    MsgBox "Group Job Order Harus diisi"
                    text(0).SetFocus
                    Exit Sub
                End If
            Case 3
                If text(Index) <> "" Then
                    sSQL = "SELECT * FROM MEKANIK WHERE MECH = '" & text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "MEKANIK"
                        gBantuan = "ALL"
                        gBantuan4 = ""
                        frmCari.Show (vbModal)
                    Else
                        text(Index + 1) = oRs("NAME")
                    End If
                    oRs.Close
                Else
                    text(Index + 1) = ""
                End If
            Case 5
                If text(Index) <> "" Then
                    sSQL = "SELECT * FROM T_PETUGAS WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "SA"
                        frmCari.Show (vbModal)
                    Else
                        text(Index + 1) = oRs("FS_NM_PETUGAS")
                    End If
                    oRs.Close
                Else
                    text(Index + 1) = ""
                End If
            Case 7
                If text(Index) <> "" Then
                    sSQL = "SELECT * FROM MODEL WHERE MDL = '" & text(Index) & "' and FS_LOG_DEL = ''"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "MDL"
                        frmCari.Show (vbModal)
                    Else
                        text(Index + 1) = oRs("DESC")
                    End If
                    oRs.Close
                Else
                    text(Index + 1) = ""
                End If
            Case 9
                If text(Index) <> "" Then
                    sSQL = "SELECT * FROM T_DEALER WHERE FS_KD_DEALER = '" & text(Index) & "' and FS_LOG_DEL = ''"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gBantuan = "ALL"
                        gSearchOption = "DEALER"
                        frmCari.Show (vbModal)
                    Else
                        text(Index + 1) = oRs("FS_NM_DEALER")
                    End If
                    oRs.Close
                Else
                    text(Index + 1) = ""
                End If
            Case 11
                If text(Index) <> "" Then
                    sSQL = "SELECT TPSERVICE FROM T_REPO WHERE TPSERVICE = '" & text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "TPSVC"
                        frmCari.Show (vbModal)
                    End If
                    oRs.Close
                End If
            Case 12
                If text(Index) <> "" Then
                    sSQL = "SELECT TPORDER FROM T_REPO WHERE TPORDER = '" & text(Index) & "' AND TPORDER <> ''"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "TP_ORD"
                        frmCari.Show (vbModal)
                    End If
                    oRs.Close
                End If
            Case 13
                If text(Index) <> "" Then
                    sSQL = "SELECT PROSES FROM T_REPO WHERE PROSES = '" & text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "PROSES"
                        frmCari.Show (vbModal)
                    End If
                    oRs.Close
                End If
            Case 14
                If text(Index) <> "" Then
                Dim rsSearch As New ADODB.Recordset
                If rsSearch.State Then rsSearch.Close
                Set rsSearch = oConn.Execute("SELECT * FROM T_STATUSAKUN WHERE TIPE = '" & Trim(text(Index)) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "STATUSAKUN"
                    frmCari.Show (vbModal)
                End If
                rsSearch.Close
            End If
        End Select
'    End If
End Sub


