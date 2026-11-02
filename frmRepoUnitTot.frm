VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmRepoUnitTot 
   Caption         =   "Laporan Jumlah Total Unit Masuk"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   12630
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
      Left            =   6600
      TabIndex        =   41
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   480
      TabIndex        =   15
      Top             =   480
      Width           =   12015
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   22937601
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   2
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2400
         TabIndex        =   16
         Top             =   840
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
         Format          =   22937601
         CurrentDate     =   39852
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   12000
         Y1              =   600
         Y2              =   600
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
         TabIndex        =   40
         Top             =   240
         Width           =   60
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
         TabIndex        =   39
         Top             =   240
         Width           =   570
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
         TabIndex        =   38
         Top             =   1965
         Width           =   480
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
         TabIndex        =   37
         Top             =   1605
         Width           =   765
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
         TabIndex        =   36
         Top             =   1245
         Width           =   870
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
         TabIndex        =   34
         Top             =   2325
         Width           =   1245
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
         TabIndex        =   33
         Top             =   1965
         Width           =   195
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
         TabIndex        =   32
         Top             =   1605
         Width           =   660
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
         TabIndex        =   31
         Top             =   1245
         Width           =   720
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
         TabIndex        =   30
         Top             =   885
         Width           =   1200
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
      Left            =   10560
      TabIndex        =   12
      Top             =   3720
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
      Left            =   8640
      TabIndex        =   11
      Top             =   3720
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
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
      FormatString    =   "Laporan Jumlah Total Unit Masuk"
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
Attribute VB_Name = "frmRepoUnitTot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : -
'=======================================

Dim rsTemp As New ADODB.Recordset
Dim vNamaTabel As String

Private Sub hapusForm()
    For i = 0 To 13
        text(i) = ""
    Next
    tgl1 = gTglServer
    tgl2 = gTglServer
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As New ADODB.Recordset

    createTemp = False
    If rsHlp.State Then rsHlp.Close
    Set rsHlp = oConn.Execute("select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepo(vNamaTabel)
    End If
End Function

Private Sub IsiTempRepo(ket As String, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
'Dim vNamaTabel As String
    
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT * FROM JOBDTA A LEFT JOIN STALL B ON REPLACE(SUBSTRING(B.KETERANGAN, CHARINDEX('|',B.KETERANGAN)+1,8),'|','') = A.NO_WO AND B.TANGGAL = A.[DATE] WHERE DELIV <> 'BTL' AND [DATE] BETWEEN '" & Format(tgl1, "MM/DD/YYYY") & "' AND '" & Format(tgl2, "MM/DD/YYYY") & "' " & ket & "")
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            For i = 1 To 8
                hlp = "REP" & i
                If rsTemp(hlp) <> "" Then
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "ACTRATE" & i
                    hlp4 = "MEKANIK" & i
                    If createTemp(vNamaTabel) Then
                        oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & rsTemp("NO_CHASSIS") & "', '" & rsTemp("NO_POLISI") & "', '" & rsTemp("FS_KD_PLG") & "', '" & rsTemp("FS_KD_WIL") & "', " _
                                     & "'" & rsTemp("CUSTOMER") & "', '" & rsTemp("MDL") & "', '" & rsTemp("FMDL") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & rsTemp("NO_WO") & "', '" & rsTemp("FRONT") & "', '" & rsTemp(hlp2) & "', " _
                                     & "'" & rsTemp(hlp) & "', '" & rsTemp(hlp4) & "', '" & Replace(rsTemp(hlp3), ",", ".") & "', '0', '0', '" & Trim(rsTemp("TPSVC")) & "', '" & rsTemp("THS") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("RPR_COND") & "', " _
                                     & "'" & rsTemp("NOFAK") & "', '" & rsTemp("NOSTALL") & "', '" & rsTemp("FS_KD_DEALER") & "', '" & rsTemp("FS_NM_DEALER") & "')")
                    End If
                End If
            Next
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
Dim Report As New RptRepoUnitTot
Screen.MousePointer = vbHourglass
    
    sSQL = "SELECT  KLPKERJA1 AS JOB, REP1, A.MECH AS MEKANIK, FRONT, MDL, " _
         & "(CASE TPSVC WHEN '' THEN 'REGULAR' WHEN 'P' THEN 'PROGRAM' WHEN 'W' THEN 'WARRANTY' END) AS TIPE, " _
         & "(CASE NOFAK WHEN '' THEN 'NONBOOKING' ELSE 'BOOKING' END) AS BOOKING, " _
         & "(CASE RPR_COND WHEN 'K' THEN 'CLOSING' WHEN 'S' THEN 'FINISH' WHEN 'P' THEN 'PROCCESS' Else 'WAITING' END) AS PROSES, COUNT(*) AS TOT " _
         & "FROM JOBDTA A INNER JOIN (SELECT JOB,REMARK FROM HO_FLATE WHERE FS_LOG_DEL = '' GROUP BY JOB,REMARK) B ON A.KLPKERJA1 = B.JOB INNER JOIN MEKANIK C ON A.MECH = C.MECH " _
         & "WHERE DELIV <> 'BTL' AND FS_LOG_DEL = '' AND [DATE] BETWEEN '" & Format(tgl1, "MM/DD/YYYY") & "' AND '" & Format(tgl2, "MM/DD/YYYY") & "' " & txt & "" _
         & "GROUP BY KLPKERJA1, REMARK, REP1, A.MECH, [NAME], FRONT, MDL, TPSVC, NOFAK, RPR_COND " _
         & "ORDER BY KLPKERJA1, REMARK, REP1, A.MECH, [NAME], FRONT, MDL, TPSVC, NOFAK, RPR_COND "
         
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
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTutup_Click()
'    sSQL = "DELETE [" & vNamaTabel & "]"
'    oConn.Execute sSQL
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    Call hapusForm
'    vNamaTabel = "JmlUnit" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub TblSimpan_Click()
Dim boleh As Boolean
Dim txt As String
    
    boleh = False
    txt = ""
    If text(0) <> "" Then
        boleh = True
        txt = txt & "AND KLPKERJA1 = '" & text(0) & "' "
    End If
    If text(2) <> "" Then
        boleh = True
        txt = txt & "AND REP1 = '" & text(2) & "' "
    End If
    If text(3) <> "" Then
        boleh = True
        txt = txt & "AND MEKANIK1 = '" & text(3) & "' "
    End If
    If text(5) <> "" Then
        boleh = True
        txt = txt & "AND FRONT = '" & text(5) & "' "
    End If
    If text(7) <> "" Then
        txt = txt & "AND MDL = '" & text(7) & "' "
    End If
    If text(9) <> "" Then
        txt = txt & "AND FS_KD_DEALER = '" & text(9) & "' "
    End If
    If text(11) <> "" Then
        txt = txt & "AND TPSVC = '" & IIf(text(11) = "REGULAR", "", IIf(text(11) = "WARRANTY", "W", "P")) & "' "
    End If
    If text(12) <> "" Then
        If text(12) = "BOOKING" Then
            txt = txt & "AND NOFAK <> ''"
        ElseIf text(12) = "NON BOOKING" Then
            txt = txt & "AND NOFAK = ''"
        End If
    End If
    If text(13) <> "" Then
        If text(13) = "PROSES" Then
            txt = txt & "AND (CASE RPR_COND WHEN 'K' THEN 'CLOSING' WHEN 'S' THEN 'FINISH' WHEN 'P' THEN 'PROCCESS' WHEN 'T' THEN 'PAUSED' Else 'WAITING' END) = 'PROCCESS'"
        ElseIf text(13) = "WAITING" Then
            txt = txt & "AND (CASE RPR_COND WHEN 'K' THEN 'CLOSING' WHEN 'S' THEN 'FINISH' WHEN 'P' THEN 'PROCCESS' WHEN 'T' THEN 'PAUSED' Else 'WAITING' END) = 'WAITING'"
        ElseIf text(13) = "FINISH" Then
            txt = txt & "AND (CASE RPR_COND WHEN 'K' THEN 'CLOSING' WHEN 'S' THEN 'FINISH' WHEN 'P' THEN 'PROCCESS' WHEN 'T' THEN 'PAUSED' Else 'WAITING' END) = 'FINISH'"
        ElseIf text(13) = "CLOSING" Then
            txt = txt & "AND (CASE RPR_COND WHEN 'K' THEN 'CLOSING' WHEN 'S' THEN 'FINISH' WHEN 'P' THEN 'PROCCESS' WHEN 'T' THEN 'PAUSED' Else 'WAITING' END) = 'CLOSING'"
        ElseIf text(13) = "PAUSED" Then
            txt = txt & "AND (CASE RPR_COND WHEN 'K' THEN 'CLOSING' WHEN 'S' THEN 'FINISH' WHEN 'P' THEN 'PROCCESS' WHEN 'T' THEN 'PAUSED' Else 'WAITING' END) = 'PAUSED'"
        End If
    End If
    
    Call Cetak(vNamaTabel, txt)
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
                    text(Index + 1) = ""
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
'                        Else
'                            Text(Index + 1) = oRs("KETERANGAN")
                        End If
                        oRs.Close
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
        End Select
'    End If
End Sub


