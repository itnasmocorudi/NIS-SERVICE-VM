VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoMap 
   BackColor       =   &H00FFFFFF&
   Caption         =   "TPSS"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoMap.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   39
      Top             =   0
      Width           =   445
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
      Height          =   270
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4800
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
      Height          =   270
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   7
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
      EnableSearchControl=   -1  'True
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   12375
      Begin VB.CheckBox chLS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label Sent"
         Height          =   255
         Left            =   7320
         TabIndex        =   38
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.OptionButton opRtn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tanpa Return Order"
         Height          =   255
         Index           =   1
         Left            =   9240
         TabIndex        =   36
         Top             =   2040
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton opRtn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dengan Return Order"
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   35
         Top             =   2040
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1530
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   111214593
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         Index           =   1
         Left            =   2520
         TabIndex        =   11
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
         Index           =   5
         Left            =   7290
         TabIndex        =   5
         Top             =   1200
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
         Left            =   9240
         TabIndex        =   12
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
         Index           =   6
         Left            =   7290
         TabIndex        =   13
         Top             =   1560
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
         Index           =   3
         Left            =   7290
         TabIndex        =   4
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
         TabIndex        =   14
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
         Index           =   4
         Left            =   8280
         TabIndex        =   15
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
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   111280129
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1530
         TabIndex        =   22
         Top             =   1920
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   2880
         TabIndex        =   24
         Top             =   1920
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   1515
         TabIndex        =   26
         Top             =   2280
         Visible         =   0   'False
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   2865
         TabIndex        =   28
         Top             =   2280
         Visible         =   0   'False
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   1530
         TabIndex        =   30
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
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   2355
         TabIndex        =   31
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
         Index           =   12
         Left            =   2475
         TabIndex        =   32
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   2775
         Left            =   0
         Top             =   0
         Width           =   12375
      End
      Begin VB.Label Label12 
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
         Left            =   6000
         TabIndex        =   34
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label Label9 
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
         Top             =   1605
         Width           =   195
      End
      Begin VB.Label Label8 
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
         Left            =   2745
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Delivery"
         BeginProperty Font 
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
         Top             =   2325
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label3 
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
         Left            =   2760
         TabIndex        =   25
         Top             =   1920
         Width           =   60
      End
      Begin VB.Label Label2 
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
         Left            =   855
         TabIndex        =   23
         Top             =   1965
         Width           =   585
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
         TabIndex        =   21
         Top             =   885
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
         TabIndex        =   20
         Top             =   1245
         Width           =   720
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
         Left            =   5955
         TabIndex        =   19
         Top             =   1245
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
         TabIndex        =   18
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Service"
         BeginProperty Font 
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
         TabIndex        =   17
         Top             =   240
         Width           =   1140
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
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   0
         X2              =   12360
         Y1              =   600
         Y2              =   600
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin ucXPButton.XPButton tblRefresh 
      Height          =   495
      Left            =   9120
      TabIndex        =   40
      Top             =   3240
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
      MICON           =   "frmRepoMap.frx":21DE
      PICN            =   "frmRepoMap.frx":21FA
      PICH            =   "frmRepoMap.frx":2776
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
      Left            =   10320
      TabIndex        =   41
      Top             =   3240
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
      MICON           =   "frmRepoMap.frx":2DAD
      PICN            =   "frmRepoMap.frx":2DC9
      PICH            =   "frmRepoMap.frx":3220
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
      Left            =   11640
      TabIndex        =   42
      Top             =   3240
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
      MICON           =   "frmRepoMap.frx":370A
      PICN            =   "frmRepoMap.frx":3726
      PICH            =   "frmRepoMap.frx":3BFA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   12840
      X2              =   480
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8655
      Left            =   0
      Picture         =   "frmRepoMap.frx":4145
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 13 November 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Sejak Januari 2012 - Acep
'=======================================

'Dim rsTemp As New ADODB.Recordset
Dim vNamaTabel As String
Dim Report As Object

Private Sub hapusForm()
    For i = 0 To 10
        Select Case i
            Case 7, 8, 9, 10
                Text(i) = ">>->>->>>>"
            Case Else
                Text(i) = ""
        End Select
    Next
    Tgl1 = isnow
    Tgl2 = isnow
End Sub

Private Sub Cetak(vNamaTabel As String, Txt As String)
Dim oRs As Object

Screen.MousePointer = vbHourglass
Set Report = New rptMapCust
    
    If chLS.value = Unchecked Then
        
        If Log.RPLHistori = 0 Then
            sSQL = "SELECT DATE, A.FS_TH_MOBIL, CASE WHEN B.CUSTOMER = B.USENAME THEN '' ELSE B.CUSTOMER END AS PERUSAHAAN, ANCUST AS PEMILIK, " _
             & "TELPANCUST AS SELULAR,'' AS TELPRUMAH, '' AS TELPKANTOR, REPLACE(A.NO_POLISI,'-','') AS NO_POLISI, " _
             & "CASE WHEN FS_RC = 'R001' THEN '1000' WHEN FS_RC <> '' THEN 'Servis Berkala' ELSE (CASE TPSVC WHEN 'W' THEN 'Perbaikan Umum (Warranty)' ELSE 'Perbaikan Umum (Non Warranty)' END) END AS TIPESVC, " _
             & "FRONT , a.NO_CHASSIS, a.MDL, a.FMDL, B.DOTAM, B.ddate, ANCUST, TELPANCUST, USEADDR, USECITY, (SELECT top 1 FS_KETERANGAN FROM T_REGULAR WITH(NOLOCK) WHERE FS_RC = A.FS_RC) AS SVC, B.UseEmail as EMAIL FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS WHERE FS_LOG_DEL = '' AND A.DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'" & Txt
        Else
            sSQL = "SELECT DATE, A.FS_TH_MOBIL, CASE WHEN B.CUSTOMER = B.USENAME THEN '' ELSE B.CUSTOMER END AS PERUSAHAAN, ANCUST AS PEMILIK, " _
             & "TELPANCUST AS SELULAR,'' AS TELPRUMAH, '' AS TELPKANTOR, REPLACE(A.NO_POLISI,'-','') AS NO_POLISI, " _
             & "CASE WHEN FS_RC = 'R001' THEN '1000' WHEN FS_RC <> '' THEN 'Servis Berkala' ELSE (CASE TPSVC WHEN 'W' THEN 'Perbaikan Umum (Warranty)' ELSE 'Perbaikan Umum (Non Warranty)' END) END AS TIPESVC, " _
             & "FRONT , a.NO_CHASSIS, a.MDL, a.FMDL, B.DOTAM, B.ddate, ANCUST, TELPANCUST, USEADDR, USECITY, (SELECT top 1 FS_KETERANGAN FROM [" & dbReplikasiHO & "]..T_REGULAR WITH(NOLOCK) WHERE FS_RC = A.FS_RC) AS SVC, B.UseEmail as EMAIL FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS WHERE FS_LOG_DEL = '' AND A.DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'" & Txt
        End If
        
    Else
        
        If Log.RPLHistori = 0 Then
            sSQL = "SELECT DATE, A.FS_TH_MOBIL, CASE WHEN B.CUSTOMER = B.USENAME THEN '' ELSE B.CUSTOMER END AS PERUSAHAAN, ANCUST AS PEMILIK, " _
             & "TELPANCUST AS SELULAR,'' AS TELPRUMAH, '' AS TELPKANTOR, REPLACE(A.NO_POLISI,'-','') AS NO_POLISI, " _
             & "CASE WHEN FS_RC = 'R001' THEN '1000' WHEN FS_RC <> '' THEN 'Servis Berkala' ELSE (CASE TPSVC WHEN 'W' THEN 'Perbaikan Umum (Warranty)' ELSE 'Perbaikan Umum (Non Warranty)' END) END AS TIPESVC, " _
             & "FRONT , a.NO_CHASSIS, a.MDL, a.FMDL, B.DOTAM, B.ddate, ANCUST, TELPANCUST, USEADDR, USECITY, (SELECT top 1 FS_KETERANGAN FROM T_REGULAR WITH(NOLOCK) WHERE FS_RC = A.FS_RC) AS SVC, B.UseEmail as EMAIL FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN DTLFU C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND B.NO_CHASSIS = C.NO_CHASSIS WHERE FS_LOG_DEL = '' AND RESPON = 'K' AND A.DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'" & Txt
        Else
            sSQL = "SELECT DATE, A.FS_TH_MOBIL, CASE WHEN B.CUSTOMER = B.USENAME THEN '' ELSE B.CUSTOMER END AS PERUSAHAAN, ANCUST AS PEMILIK, " _
             & "TELPANCUST AS SELULAR,'' AS TELPRUMAH, '' AS TELPKANTOR, REPLACE(A.NO_POLISI,'-','') AS NO_POLISI, " _
             & "CASE WHEN FS_RC = 'R001' THEN '1000' WHEN FS_RC <> '' THEN 'Servis Berkala' ELSE (CASE TPSVC WHEN 'W' THEN 'Perbaikan Umum (Warranty)' ELSE 'Perbaikan Umum (Non Warranty)' END) END AS TIPESVC, " _
             & "FRONT , a.NO_CHASSIS, a.MDL, a.FMDL, B.DOTAM, B.ddate, ANCUST, TELPANCUST, USEADDR, USECITY, (SELECT top 1 FS_KETERANGAN FROM [" & dbReplikasiHO & "]..T_REGULAR WITH(NOLOCK) WHERE FS_RC = A.FS_RC) AS SVC, B.UseEmail as EMAIL FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN DTLFU C WITH(NOLOCK) ON A.NO_WO = C.NO_WO AND B.NO_CHASSIS = C.NO_CHASSIS WHERE FS_LOG_DEL = '' AND RESPON = 'K' AND A.DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'" & Txt
        End If
        
    End If
    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    
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
    cmdSettingPrinter.Visible = False
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height - 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width - 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = CRViewer1.Top + 400
    cmdSettingPrinter.Left = CRViewer1.Left + (50 + 1080)
    cmdSettingPrinter.Visible = False
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    Call hapusForm
    vNamaTabel = "MapCust" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub tblSimpan_Click()
Dim txtTemp As String
    
    txtTemp = ""
    If Text(0) <> "" Then
        txtTemp = txtTemp & " AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7+KLPKERJA8 LIKE '%" & Text(0) & "%' "
    End If
    If Text(2) <> "" Then
        txtTemp = txtTemp & " AND REP1+REP2+REP3+REP4+REP5+REP6+REP7+REP8 LIKE '%" & Text(2) & "%' "
    End If
    If Text(5) <> "" Then
        txtTemp = txtTemp & "AND A.MDL = '" & Text(5) & "' "
    End If
    If Text(11) <> "" Then
        txtTemp = txtTemp & "AND A.FRONT = '" & Text(11) & "' "
    End If
    If Text(3) <> "" Then
        txtTemp = txtTemp & "AND B.FS_KD_DEALER = '" & Text(3) & "' "
    End If
    If Text(7) <> ">>->>->>>>" And Text(8) <> ">>->>->>>>" Then
        txtTemp = txtTemp & "AND B.DOTAM BETWEEN '" & Format(Trim(Text(7)), "MM/DD/YYYY") & "' AND '" & Format(Trim(Text(8)), "MM/DD/YYYY") & "'"
    End If
    If opRtn(1).value Then
        txtTemp = txtTemp & "AND A.FN_RETURN = '0'"
    End If
    Call Cetak(vNamaTabel, txtTemp)
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
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
                gSearchOption = "JENIS_KERJA"
                frmCari.Show (vbModal)
            Case 2
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(0)
                    frmCari.Show (vbModal)
                Else
                    MsgBox "Group Job Order Harus diisi"
                    Text(0).SetFocus
                    Exit Sub
                End If
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
            Case 3
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.Show (vbModal)
            Case 11
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT DISTINCT JOB, REMARK FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = oRs("REMARK")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 2
            If Text(0) <> "" Then
                If Text(Index) <> "" Then
                    Set oRs = objO.FetchData("SELECT DISTINCT FS_ITEM, KETERANGAN FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(0) & "' AND FS_ITEM = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JOB_ORDER"
                        gBantuan = Text(0)
                        frmCari.Show (vbModal)
                    Else
                        Text(Index + 1) = oRs("KETERANGAN")
                    End If
                    oRs.Close
                End If
            Else
                MsgBox "Group Job Order Harus diisi"
                Text(0).SetFocus
                Exit Sub
            End If
        Case 5
            If Text(Index) <> "" Then
                If Log.RPLHistori = 0 Then
                    Set oRs = objO.FetchData("SELECT DISTINCT MDL, [DESC] FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Else
                    Set oRs = objO.FetchData("SELECT DISTINCT MDL, [DESC] FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                End If
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = oRs("DESC")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 3
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "ALL"
                    gSearchOption = "DEALER"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_DEALER")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 7
            If Not CekTgl(Text(Index)) Then
                If Text(Index) <> ">>->>->>>>" Then
                    MsgBox "Format tanggal salah"
                End If
                Text(Index) = ">>->>->>>>"
            End If
        Case 8
            If Not CekTgl(Text(Index)) Then
                If Text(Index) <> ">>->>->>>>" Then
                    MsgBox "Format tanggal salah"
                End If
                Text(Index) = ">>->>->>>>"
            End If
        Case 11
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_PETUGAS")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
End Sub
