VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form frmRepoInvoiceDiskon 
   Caption         =   "Lap Revenue Service Detail Discount"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5640
      TabIndex        =   25
      Top             =   6990
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrintOut 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6960
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3075
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6570
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1590
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6570
      Width           =   1440
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6570
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   3
      Top             =   6480
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   11535
      Begin VB.CheckBox chInvDealer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hanya Invoice B"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chTHS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THS"
         Height          =   255
         Left            =   6960
         TabIndex        =   27
         Top             =   2520
         Width           =   255
      End
      Begin ARProgBarCtrl.ARProgressBar ARProgressBar1 
         Height          =   375
         Left            =   3720
         TabIndex        =   26
         Top             =   4080
         Visible         =   0   'False
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
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
      Begin VB.CommandButton tblRefresh 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Refresh"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4560
         Width           =   1890
      End
      Begin VB.CommandButton TblSimpan 
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
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4560
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
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4560
         Width           =   1890
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1605
         TabIndex        =   5
         Top             =   720
         Width           =   3855
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1 Model"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Beberapa Model"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   6
            Top             =   0
            Width           =   1695
         End
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   100204545
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   100204545
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1575
         TabIndex        =   10
         Top             =   1065
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
         Left            =   3525
         TabIndex        =   11
         Top             =   1065
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
         Index           =   8
         Left            =   1575
         TabIndex        =   12
         Top             =   1425
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
         Index           =   14
         Left            =   1965
         TabIndex        =   13
         Top             =   1425
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   20
         Top             =   1800
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
         Top             =   1800
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
         TabIndex        =   22
         Top             =   1800
         Width           =   3150
         _ExtentX        =   5556
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
         Index           =   9
         Left            =   1560
         TabIndex        =   29
         Top             =   2160
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
         Left            =   2370
         TabIndex        =   30
         Top             =   2160
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
         Left            =   2520
         TabIndex        =   31
         Top             =   2160
         Width           =   3150
         _ExtentX        =   5556
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
         Left            =   1560
         TabIndex        =   32
         Top             =   2880
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   2790
         TabIndex        =   33
         Top             =   2880
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
         Left            =   6930
         TabIndex        =   34
         Top             =   720
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
         Left            =   8775
         TabIndex        =   35
         Top             =   720
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
         Left            =   6930
         TabIndex        =   36
         Top             =   1080
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
         Left            =   8775
         TabIndex        =   37
         Top             =   1080
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
         Index           =   0
         Left            =   6930
         TabIndex        =   38
         Top             =   1440
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
         Left            =   8775
         TabIndex        =   39
         Top             =   1440
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
         Index           =   15
         Left            =   6930
         TabIndex        =   40
         Top             =   1800
         Width           =   1755
         _ExtentX        =   3096
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   8775
         TabIndex        =   41
         Top             =   1800
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
         Index           =   16
         Left            =   6930
         TabIndex        =   42
         Top             =   2160
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   503
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   4125
         TabIndex        =   43
         Top             =   2880
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox MaskEdBox11 
         Height          =   270
         Left            =   5520
         TabIndex        =   44
         Top             =   2880
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
         Index           =   25
         Left            =   1560
         TabIndex        =   45
         Top             =   3240
         Visible         =   0   'False
         Width           =   4155
         _ExtentX        =   7329
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
         Index           =   1
         Left            =   8880
         TabIndex        =   54
         Top             =   1440
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
         Left            =   6360
         TabIndex        =   53
         Top             =   1125
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
         Left            =   6075
         TabIndex        =   52
         Top             =   765
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
         Left            =   420
         TabIndex        =   51
         Top             =   2925
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
         Left            =   255
         TabIndex        =   50
         Top             =   2205
         Width           =   1035
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
         Left            =   6360
         TabIndex        =   49
         Top             =   1500
         Width           =   465
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
         Left            =   6165
         TabIndex        =   48
         Top             =   1830
         Width           =   660
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   6480
         TabIndex        =   47
         Top             =   2520
         Width           =   285
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Program"
         BeginProperty Font 
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
         TabIndex        =   46
         Top             =   2925
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   1200
         TabIndex        =   23
         Top             =   1800
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Invoice"
         BeginProperty Font 
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
         TabIndex        =   16
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
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   60
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   6960
         Y1              =   600
         Y2              =   600
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
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1245
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8175
      Left            =   0
      Picture         =   "frmRepoInvoiceDiskon.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoInvoiceDiskon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim Report As Object
 Private rsCetak As ADODB.Recordset

Private Sub createTable()
1     Set rsCetak = New ADODB.Recordset
2         If rsCetak.State Then rsCetak.Close
          'rsLaporan.Close
          
          
3         With rsCetak
4             .Fields.Append "NoInv", adVarChar, 50
5             .Fields.Append "tglInv", adDate
6             .Fields.Append "NoWo", adVarChar, 50
7             .Fields.Append "NoPolisi", adVarChar, 50
8             .Fields.Append "Customer", adVarChar, 200
9             .Fields.Append "SA", adVarChar, 50
10            .Fields.Append "TipeSV", adVarChar, 200
11            .Fields.Append "NamaProgram", adVarChar, 200
12            .Fields.Append "HargaAJasa", adDouble
13            .Fields.Append "HargaAPart", adDouble
14            .Fields.Append "HargaASublet", adDouble
15            .Fields.Append "HargaDJasa", adDouble
16            .Fields.Append "HargaDPart", adDouble
17            .Fields.Append "HargaDSublet", adDouble
18            .Fields.Append "HargaTJasa", adDouble
19            .Fields.Append "HargaTPart", adDouble
20            .Fields.Append "HargaTSublet", adDouble
21            .Fields.Append "HargaTMaterai", adDouble
22            .Fields.Append "HargaTTotal", adDouble
23            .Fields.Append "ValDiscJasa", adVarChar, 50
24            .Fields.Append "ValDiscPart", adVarChar, 50
25            .Fields.Append "ValDiscSublet", adVarChar, 50
26            .Open
27        End With
End Sub


Private Sub Form_Load()
    Call bersih
End Sub

Private Sub opVehicle_Click(Index As Integer)
    If opVehicle(0) Then
        Text(7).Visible = True
        Text(8).Visible = True
        MaskEdBox4.Visible = True
        Text(14).Visible = False
    Else
        Text(7).Visible = False
        Text(8).Visible = False
        MaskEdBox4.Visible = False
        Text(14).Visible = True
        Text(14).Left = Text(7).Left
        Text(14).SelLength = Len(Text(14))
        Text(14).SetFocus
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call bersih
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdPrintOut.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
     
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = cmdExport.Left + cmdExport.Width + 30
    
    cmdPrintOut.Top = CRViewer1.Top + 345
    cmdPrintOut.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub bersih()
    Tgl1 = Format(isnow, "DD/MM/YYYY")
    Tgl2 = Format(isnow, "DD/MM/YYYY")
    opVehicle(0).value = True
    Text(5).Text = "": Text(6).Text = "": Text(7).Text = ""
    Text(8).Text = "": Text(14).Text = "": Text(9).Text = ""
    Text(10).Text = "": Text(11).Text = "": Text(24).Text = ""
    Text(25).Text = "": Text(12).Text = "": Text(13).Text = ""
    Text(0).Text = "": Text(1).Text = "": Text(15).Text = "": Text(16).Text = ""
    
    txtTemp = ""
    
    
    chTHS.value = Unchecked
    chInvDealer.value = Unchecked
    
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
    
End Sub


Private Sub tblSimpan_Click()
          Dim boleh As Boolean
          Dim txtTemp As String
          Dim txt As String
          Dim rsTemp As ADODB.Recordset
          Dim TipeSV As String
          Dim discJasa As String
          Dim discSublet As String
          Dim isRetur As Boolean
          Dim pengkali  As Integer
1         Screen.MousePointer = vbHourglas
2         txtTemp = ""
3         If NulSID(Text(9)) <> "" Then
4             txtTemp = txtTemp & "AND B.FS_KD_DEALER = '" & NulSID(Text(9)) & "' "
5         End If
          
6         If NulSID(Text(5)) <> "" Then
7             txtTemp = txtTemp & "AND a.FRONT = '" & Text(5) & "' "
8         End If
9         If NulSID(Text(7)) <> "" Or NulSID(Text(14)) <> "" Then
10            If opVehicle(0) Then
11                txtTemp = txtTemp & "AND A.MDL = '" & Trim(Text(7)) & "' "
12            Else
13                txtTemp = txtTemp & "AND A.FMDL LIKE '%" & Trim(Text(14)) & "%' "
14            End If
15        End If
          
16         If Text(11) <> "" Then
17            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
18                txtTemp = txtTemp & "AND a.TPSVC = '" & IIf(Text(11) = "WARRANTY", "W", "") & "' "
19            Else
20                txtTemp = txtTemp & "AND a.TPSVC = '" & IIf(Text(11) = "REGULAR", "", IIf(Text(11) = "WARRANTY", "W", "P")) & "' "
21            End If
22            If Trim(Text(11)) = "PROGRAM" Then
23                If Trim(Text(24)) <> "" Then
24                    If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
25                        txtTemp = txtTemp & "AND (a.KODETIPE = '" & Trim(Text(24)) & "' OR A.no_wo in (select no_wo from T_invoiceJobDetail where FS_LOG_DEL = '' and no_wo = A.no_wo and id_program = '" & Trim(Text(24)) & "' )) "
26                    Else
27                        txtTemp = txtTemp & "AND a.KODETIPE = '" & Trim(Text(24)) & "' "
28                    End If
29                End If
30            End If
31        End If
          
32        If Text(12) <> "" Then
33            If Text(12) = "BOOKING" Then
34                txtTemp = txtTemp & "AND b.NOFAK <> ''"
35            ElseIf Text(12) = "NON BOOKING" Then
36                txtTemp = txtTemp & "AND b.NOFAK = ''"
37            End If
38        End If
          
39        If Text(13) <> "" Then
40            If Text(13) = "PROSES" Then
41                txtTemp = txtTemp & "AND a.RPR_COND = 'P'"
42            ElseIf Text(13) = "WAITING" Then
43                txtTemp = txtTemp & "AND a.RPR_COND = '' AND b.NOINV = ''"
44            ElseIf Text(13) = "FINISH" Then
45                txtTemp = txtTemp & "AND a.RPR_COND = 'S'"
46            ElseIf Text(13) = "CLOSING" Then
47                txtTemp = txtTemp & "AND a.RPR_COND = 'K'"
48            ElseIf Text(13) = "PAUSED" Then
49                txtTemp = txtTemp & "AND a.RPR_COND = 'T'"
50            End If
51        End If
          
52        If NulSID(Text(0)) <> "" Then
53             txtTemp = txtTemp & "AND (a.klpkerja1+a.klpkerja2+a.klpkerja3+a.klpkerja4+a.klpkerja5+a.klpkerja6+a.klpkerja7+a.klpkerja8) like '%" & NulSID(Text(0)) & "%'"
54        End If
          
55        If chInvDealer = Checked Then
56            txtTemp = txtTemp & " and RIGHT(A.NOINV,1) = 'B'"
57        End If
58        If chTHS = Checked Then
59            txtTemp = txtTemp & " and a.THS = '1'"
60        End If
          
61        sSQL = "select * from ( select a.no_chassis,a.TPSVC,a.PDS,a.FS_RC,a.KLPKERJA1+a.KLPKERJA1+a.KLPKERJA2+a.KLPKERJA3+a.KLPKERJA4+a.KLPKERJA5+a.KLPKERJA6+a.KLPKERJA7+a.KLPKERJA8 as pekerjaan, a.INVDATE as invdate,a.NO_WO, a.NOINV as NOINV,a.NO_POLISI,"
62        sSQL = sSQL & " a.CUSTOMER,a.FRONT,b.nofk,b.dbasefk,"
63        sSQL = sSQL & " a.OFF1+a.off2+a.OFF3+a.off4+a.OFF5+a.off6+a.OFF7+a.OFF8+ISNULL((select SUM(offn) from HIST_JOBADD where NO_WO = a.no_wo and NO_CHASSIS = a.no_chassis ),0)  as jasa,"
64        sSQL = sSQL & " isnull((select sum(round(bb.HARGAJUAL * " & k2t(profilACC.nilaitotal / profilACC.nilaidpp) & ",0)*aa.qty) from HIST_SUBLET aa inner join [ReplBengkel99]..MSTSUBLET bb on aa.KODEITEM = bb.KDSUBLET where aa.noinv = a.noinv and aa.NO_WO = a.NO_WO and aa.NO_CHASSIS = a.NO_CHASSIS ),0) as sublet,"
65        sSQL = sSQL & " a.LABOR as laborinv,a.PART_G+a.OIL_G as invgenuine , a.PART_NG + a.OIL_NG as invnong,a.SBLT,a.TOTAL,a.MET"
66        sSQL = sSQL & " from T_INVOICE A"
67        sSQL = sSQL & " inner join JOBDTA b on a.NO_WO = b.NO_WO and a.NO_CHASSIS = b.NO_CHASSIS and b.FS_LOG_DEL = ''"
68        sSQL = sSQL & " where (a.INVDATE between '" & Format(Tgl1, "yyyy-mm-dd") & "' and '" & Format(Tgl2, "yyyy-mm-dd") & "') "
69        sSQL = sSQL & " " & txtTemp & " union "
70        sSQL = sSQL & " select a.no_chassis,a.TPSVC,a.PDS,a.FS_RC,a.KLPKERJA1+a.KLPKERJA1+a.KLPKERJA2+a.KLPKERJA3+a.KLPKERJA4+a.KLPKERJA5+a.KLPKERJA6+a.KLPKERJA7+a.KLPKERJA8 as pekerjaan, a.FD_RETUR as invdate,a.NO_WO, a.NOINV+'R' as NOINV,a.NO_POLISI,"
71        sSQL = sSQL & " a.CUSTOMER,a.FRONT,b.nofk,b.dbasefk,"
72        sSQL = sSQL & " a.OFF1+a.off2+a.OFF3+a.off4+a.OFF5+a.off6+a.OFF7+a.OFF8+ISNULL((select SUM(offn) from HIST_JOBADD where NO_WO = a.no_wo and NO_CHASSIS = a.no_chassis ),0)  as jasa,"
73        sSQL = sSQL & " isnull((select sum(round(bb.HARGAJUAL * " & k2t(profilACC.nilaitotal / profilACC.nilaidpp) & ",0)*aa.qty) from HIST_SUBLET aa inner join [ReplBengkel99]..MSTSUBLET bb on aa.KODEITEM = bb.KDSUBLET where aa.noinv = a.noinv and aa.NO_WO = a.NO_WO and aa.NO_CHASSIS = a.NO_CHASSIS ),0) as sublet,"
74        sSQL = sSQL & " a.LABOR as laborinv,a.PART_G+a.OIL_G as invgenuine , a.PART_NG + a.OIL_NG as invnong,a.SBLT,a.TOTAL,a.MET"
75        sSQL = sSQL & " from T_INVOICE A"
76        sSQL = sSQL & " inner join JOBDTA b on a.NO_WO = b.NO_WO and a.NO_CHASSIS = b.NO_CHASSIS and b.FS_LOG_DEL = ''"
77        sSQL = sSQL & " where (a.FD_RETUR between '" & Format(Tgl1, "yyyy-mm-dd") & "' and '" & Format(Tgl2, "yyyy-mm-dd") & " 23:59:59')"
78        sSQL = sSQL & " " & txtTemp & " ) a order by a.noinv" ''sSQL = sSQL & " and a.NOINV not like '%b' " & txtTemp
79        Set rsTemp = objO.FetchData(sSQL)
80        Call createTable
81        If Not rsTemp.EOF Then
82            ARProgressBar1.Visible = True
83            ARProgressBar1.Max = rsTemp.RecordCount + 1
84            ARProgressBar1.value = 0
85            rsTemp.MoveFirst
86            Do While Not rsTemp.EOF
87                isRetur = False
88                pengkali = 1
89                If Right(NulSID(rsTemp!noInv), 1) = "R" Then
90                    isRetur = True
91                    pengkali = -1
92                End If
93                TipeSV = ""
94                discJasa = ""
95                discSublet = ""
96                sSQL = "SELECT * FROM RequestDisc WITH (NOLOCK) WHERE (userValid <> 'BTL' AND uservalid <> '') and no_wo = '" & NulSID(rsTemp!no_wo) & "'"
97                If Log.multipleMarketing = 1 Then
98                    sSQL1 = "select b.PROGRAM from jobdetail a inner join [ReplBengkel99]..HO_PROGRAM b on a.id_Program = b.id_Program where NO_WO = '" & NulSID(rsTemp!no_wo) & "' and b.PROGRAM <> ''"
99                Else
100                   sSQL1 = "select top 1 * from jobdta where 1=0"
101               End If
102               If NulSID(rsTemp!TPSVC) = "P" Then
103                   discJasa = "Program"
104               ElseIf objO.FetchData_Count(sSQL1) > 0 Then
105                   discJasa = "Program"
106               ElseIf objO.FetchData_Count(sSQL) > 0 Then
107                   discJasa = "Validasi Kacab"
108               Else
109                   discJasa = "Disc SA"
110               End If
                  
111               sSQL = "select * from HIST_SUBLET aa inner join [ReplBengkel99]..MSTSUBLET bb on aa.KODEITEM = bb.KDSUBLET where aa.NO_WO = '" & NulSID(rsTemp!no_wo) & "' and aa.NO_CHASSIS = '" & NulSID(rsTemp!NO_CHASSIS) & "' and aa.kdpromo <> ''"
112               If objO.FetchData_Count(sSQL) > 0 Then
113                   discSublet = "Program"
114               Else
115                   discSublet = "Disc SA"
116               End If
117               If NulSID(rsTemp!TPSVC) = "W" Then
118                   TipeSV = "Warranty"
119               ElseIf NulSID(rsTemp!pds) = "1" Then
120                   TipeSV = "PDS"
121               ElseIf NulSID(rsTemp!FS_RC) = "R001" Then
122                   TipeSV = "SBI"
123               ElseIf NulSID(rsTemp!FS_RC) <> "" Then
124                   TipeSV = "SBE"
125               ElseIf InStr(NulSID(rsTemp!pekerjaan), "GRP") > 0 Then
126                   TipeSV = "GRP"
127               ElseIf InStr(NulSID(rsTemp!pekerjaan), "LUB") > 0 Then
128                   TipeSV = "LUB"
129               End If
130               With rsCetak
131                   .AddNew
132                   !noInv = NulSID(rsTemp!noInv)
133                   !tglInv = Format(NulSID(rsTemp!InvDate), "yyyy-mm-dd")
134                   !noWO = NulSID(rsTemp!no_wo)
135                   !NoPolisi = NulSID(rsTemp!NO_POLISI)
136                   !customer = NulSID(rsTemp!customer)
137                   !sa = NulSID(rsTemp!FRONT)
138                   !TipeSV = NulSID(TipeSV)
139                   !namaProgram = getNamaProgram(NulSID(rsTemp!no_wo))
140                   !HargaAJasa = NulSID(rsTemp!jasa) * pengkali
141                   If NulSID(rsTemp!NOFK) <> "" And NulSID(rsTemp!dBaseFK) <> "" Then
142                       sSQL = "select isnull(SUM(aa.hargajual*aa.qty) * " & k2t(profilACC.nilaitotal / profilACC.nilaidpp) & " ,0) as total"
143                       sSQL = sSQL & " from [" & NulSID(rsTemp!dBaseFK) & "].dbo.Sisal aa"
144                       sSQL = sSQL & " inner join [" & NulSID(rsTemp!dBaseFK) & "].dbo.Inmst bb on aa.KodeItem = bb.KodeItem"
145                       sSQL = sSQL & " where aa.NoFaktur = '" & NulSID(rsTemp!NOFK) & "'"
146                       !HargaAPart = objO.FetchData_Double(sSQL) * pengkali
147                   Else
148                       !HargaAPart = 0 * pengkali
149                   End If
                      
150                   !HargaASublet = CDbl(NulSID(rsTemp!SUBLET)) * pengkali
151                   !HargaTJasa = CDbl(NulSID(rsTemp!laborinv)) * pengkali
152                   !HargaTPart = (CDbl(NulSID(rsTemp!invgenuine)) + CDbl(NulSID(rsTemp!invnong))) * pengkali
153                   !HargaTSublet = CDbl(NulSID(rsTemp!sblt)) * pengkali
154                   !HargaTMaterai = CDbl(NulSID(rsTemp!met)) * pengkali
155                   !HargaTTotal = CDbl(NulSID(rsTemp!total)) * pengkali
                      
156                   !HargaDJasa = CDbl(!HargaAJasa) - CDbl(!HargaTJasa)
157                   If (CDbl(!HargaAPart) - CDbl(!HargaTPart)) < 0 And pengkali > 0 Then
158                       !HargaDPart = 0 * pengkali
159                   ElseIf (CDbl(!HargaAPart) + CDbl(!HargaTPart)) > 0 And pengkali < 0 Then
160                       !HargaDPart = 0 * pengkali
161                   Else
162                       !HargaDPart = (CDbl(!HargaAPart) - CDbl(!HargaTPart)) * pengkali
163                   End If
164                   !HargaDSublet = (CDbl(!HargaASublet) - CDbl(!HargaTSublet)) * pengkali
                      
165                   !ValDiscJasa = discJasa
166                   !ValDiscPart = discJasa
167                   !ValDiscSublet = discSublet
168                   .Update
169               End With
170               ARProgressBar1.value = ARProgressBar1.value + 1
171               DoEvents
                  
172               rsTemp.MoveNext
173           Loop
174       End If
175       ARProgressBar1.Visible = False
176       Set Report = New RptRepoInvoiceDiskon
177       Report.database.SetDataSource rsCetak
178       Report.txtPerusahaan.SetText gPerush
179       Report.txtAlamat.SetText gAlamat
180       Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
181       With CRViewer1
182           .ReportSource = Report
183           .ViewReport
184           .Zoom 100
185           .Visible = True
186       End With
187       cmdTutup.Visible = True
188       cmdSettingPrinter.Visible = True
189       Screen.MousePointer = vbDefault
190       Call DisableExport
191       Call setLokasiTombol
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
    asxPowerBanner1.Height = Me.Height - 500
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

Private Function getNamaProgram(noWO As String) As String
    Dim hasil As String
    Dim rsNamaProgram As ADODB.Recordset
    
    sSQL = "select b.PROGRAM"
    sSQL = sSQL & " from T_INVOICE a left join [ReplBengkel99]..HO_PROGRAM b on a.KODETIPE = b.id_Program where NO_WO = '" & noWO & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select b.PROGRAM"
    sSQL = sSQL & " from jobdetail a left join [ReplBengkel99]..HO_PROGRAM b on a.id_Program = b.id_Program where NO_WO = '" & noWO & "'"
    Set rsNamaProgram = objO.FetchData(sSQL)
    If Not rsNamaProgram.EOF Then
        rsNamaProgram.MoveFirst
        Do While Not rsNamaProgram.EOF
            If NulSID(rsNamaProgram!program) <> "" Then
                hasil = hasil & ", " & NulSID(rsNamaProgram!program)
            End If
            rsNamaProgram.MoveNext
        Loop
    End If
    getNamaProgram = hasil
End Function

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.show (vbModal)
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 9
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 11
                gMainKeyCari = Text(Index)
                gSearchOption = "TPSVC"
                frmCari.show (vbModal)
            Case 12
                gMainKeyCari = Text(Index)
                gSearchOption = "TP_ORD"
                frmCari.show (vbModal)
            Case 13
                gMainKeyCari = Text(Index)
                gSearchOption = "PROSES"
                frmCari.show (vbModal)
            Case 15
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.show (vbModal)
            Case 21
                gMainKeyCari = Text(Index)
                gSearchOption = "DLRSHIFT"
                gBantuan = "WHERE BERLAKU = 1 "
                frmCari.show (vbModal)
            Case 24
                ' PROGRAM
                gMainKeyCari = Text(Index)
                gSearchOption = "PROGRAM_ALL"
                gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset
Dim rsOpsi As ADODB.Recordset

    Text(Index) = Replace(Trim(Text(Index)), "'", "")
    Select Case Index
        
        Case 5
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("FS_NM_PETUGAS")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 7
            If Text(Index) <> "" Then
                If Log.RPLHistori = 0 Then
                    Set rsSearch = objO.FetchData("SELECT * FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Else
                    Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                End If
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("DESC")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 0
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE JOB = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("REMARK")
                End If
            Else
                Text(1) = ""
            End If
        Case 9
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "ALL"
                    gSearchOption = "DEALER"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("FS_NM_DEALER")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 11
            chInvDealer.Visible = False
            chInvDealer = Unchecked
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT TPSERVICE FROM T_REPO WITH(NOLOCK) WHERE TPSERVICE = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "TPSVC"
                    frmCari.show (vbModal)
                End If
                
                chInvDealer.Visible = False
                Label19.Visible = False
                Text(24).Visible = False
                Text(25).Visible = False
                MaskEdBox11.Visible = False
                If Text(Index) = "PROGRAM" Then
                    chInvDealer.Visible = True
                    Label19.Visible = True
                    Text(24).Visible = True
                    Text(25).Visible = True
                    MaskEdBox11.Visible = True
                End If
            End If
        Case 12
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT TPORDER FROM T_REPO WITH(NOLOCK) WHERE TPORDER = '" & Text(Index) & "' AND TPORDER <> ''")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "TP_ORD"
                    frmCari.show (vbModal)
                End If
            End If
        Case 13
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT PROSES FROM T_REPO WITH(NOLOCK) WHERE PROSES = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PROSES"
                    frmCari.show (vbModal)
                End If
            End If
        Case 15
            If Text(Index) <> "" Then
                    Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & Text(Index) & "'")
                    If rsSearch.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "KODEPLG"
                        frmCariAcess.show (vbModal)
                    Else
                        Text(16) = rsSearch("NAMALGN")
                   End If
            Else
                Text(16) = ""
            End If
        Case 21
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT SHIFT, JAMMULAI, JAMSELESAI FROM DEALERSHIFT WITH(NOLOCK) WHERE shift = '" & Trim(Text(Index)) & "' AND BERLAKU = 1")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "DLRSHIFT"
                    gBantuan = "WHERE BERLAKU = 1 "
                    frmCari.show (vbModal)
                Else
                    Text(22) = Trim(rsSearch("JAMMULAI"))
                    Text(23) = Trim(rsSearch("JAMSELESAI"))
                End If
            End If
        Case 24
            If Text(Index) <> "" Then
                
                If Log.RPLHistori = 0 Then
                    Set rsOpsi = objO.FetchData("SELECT * FROM HO_program A WITH(NOLOCK) LEFT JOIN HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                Else
                    If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) >= 0 Then
                        sSQL = "SELECT PROGRAM FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
                        sSQL = sSQL & " UNION "
                        sSQL = sSQL & " SELECT NAMAFP AS PROGRAM FROM [" & dbReplikasiHO & "]..HO_FIXPRICE A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_FIXPRICECBG B WITH(NOLOCK) ON A.id_FP = B.id_FP WHERE A.id_FP = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''"
                        Set rsOpsi = objO.FetchData(sSQL)
                    Else
                        Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                    End If
                    ''Set rsOpsi = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_program A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(Text(Index)) & "' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END and FS_LOG_DEL = ''")
                End If
                If rsOpsi.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PROGRAM_ALL"
                    gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
                    frmCari.show (vbModal)
                Else
                    Text(25) = rsOpsi("PROGRAM")
                End If
            End If
    End Select
End Sub

