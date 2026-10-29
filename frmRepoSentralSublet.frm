VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRepoSentralSublet 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Pencapaian Sublet"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   3870
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   476
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoSentralSublet.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   4920
      TabIndex        =   18
      Top             =   30
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdCetak 
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
      Left            =   8040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
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
      Left            =   2595
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
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
      Left            =   1110
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1050
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoSentralSublet.frx":001C
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   9
      Top             =   0
      Width           =   445
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   4920
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
      EnableRefreshButton=   -1  'True
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
      Height          =   3735
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   7455
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   960
         Width           =   2895
         Begin VB.OptionButton optCreator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optCreator 
            BackColor       =   &H00FFFFFF&
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
            Height          =   375
            Index           =   1
            Left            =   630
            TabIndex        =   31
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optCreator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1470
            TabIndex        =   30
            Top             =   0
            Width           =   615
         End
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   450
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Visible         =   0   'False
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   794
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rekap"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton opCetak 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detail"
         Height          =   255
         Index           =   0
         Left            =   1650
         TabIndex        =   6
         Top             =   2040
         Value           =   -1  'True
         Width           =   945
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1650
         TabIndex        =   0
         Top             =   1650
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   123142145
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   1650
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   123142145
         CurrentDate     =   39852
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbPromo 
         Height          =   285
         Left            =   1650
         TabIndex        =   20
         Top             =   675
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbPekerjaan 
         Height          =   285
         Left            =   1650
         TabIndex        =   21
         Top             =   45
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbSupplier 
         Height          =   285
         Left            =   1650
         TabIndex        =   24
         Top             =   360
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   6000
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   6120
         Top             =   600
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbDealer 
         Height          =   285
         Left            =   1650
         TabIndex        =   27
         Top             =   1320
         Width           =   2415
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   375
         Left            =   5880
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   6000
         Top             =   1680
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label3 
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
         Left            =   945
         TabIndex        =   28
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Creator Master"
         BeginProperty Font 
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
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         BeginProperty Font 
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
         TabIndex        =   25
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label4 
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
         Left            =   990
         TabIndex        =   23
         Top             =   45
         Width           =   450
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promo"
         BeginProperty Font 
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
         TabIndex        =   22
         Top             =   675
         Width           =   450
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   2655
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
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
         TabIndex        =   5
         Top             =   1680
         Width           =   540
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
         TabIndex        =   4
         Top             =   240
         Width           =   60
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
   Begin ucXPButton.XPButton tblSimpan 
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   4200
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
      MICON           =   "frmRepoSentralSublet.frx":918A
      PICN            =   "frmRepoSentralSublet.frx":91A6
      PICH            =   "frmRepoSentralSublet.frx":95FD
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
      Left            =   6720
      TabIndex        =   12
      Top             =   4200
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
      MICON           =   "frmRepoSentralSublet.frx":9AE7
      PICN            =   "frmRepoSentralSublet.frx":9B03
      PICH            =   "frmRepoSentralSublet.frx":9FD7
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
      X1              =   480
      X2              =   7920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label LblKnow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE WAIT  (^_^) ..."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label prosen 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "     %"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   3600
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8295
      Left            =   0
      Picture         =   "frmRepoSentralSublet.frx":A522
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoSentralSublet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim rsTemp As ADODB.Recordset
Dim Report As Object

Dim doubleJob As Boolean
Dim resetJob As Boolean
Dim resetRate As Boolean
Dim vNamaTabel As String
Dim objRepo As New clsRepo
Dim rsGroup    As New ADODB.Recordset
Dim rsLaporan  As New ADODB.Recordset

Private Sub hapusForm()
    Tgl1 = isnow
    Tgl2 = isnow
    LblKnow.Visible = False
    ProgressBar1.Visible = False
    prosen.Visible = False
End Sub
 

Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
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
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Report.PaperOrientation = POr
    Report.PaperSize = Psize
    Report.PrintOut
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
         
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
'    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
'         & " drop table [tempdb]..[" & vNamaTabel & "]"
'    Call objO.UpdateData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    onTimer = False
    
    If gUstat = 0 Then
        Frame2.Visible = False
        Label2.Visible = False
        
        cbDealer.Visible = False
        Label3.Visible = False
        
        Label10.Top = Label2.Top
        Tgl1.Top = Frame2.Top
        Tgl2.Top = Frame2.Top
        
        opCetak(0).Top = cbDealer.Top
        opCetak(1).Top = cbDealer.Top
    End If
    
    Call hapusForm
    Call refreshPekerjaan
    Call refreshPromo
    Call refreshDealer
    Call refreshSupplier
End Sub

 

Private Sub TblKeluar_Click()
    onTimer = True
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim txtTemp As String
Dim Txt As String
Dim dealer As String
    
    boleh = False
    txtTemp = ""
    Txt = ""
    ProgressBar1.Visible = False
    prosen.Visible = False
    
    If NulSID(cbPekerjaan.value) <> "" Then
        Txt = Txt & " and a.[KODEITEM] = '" & NulSID(cbPekerjaan.value) & "'"
    End If
    
    If NulSID(cbPromo.value) <> "" Then
        Txt = Txt & " and e.[kdpromo] = '" & NulSID(cbPromo.value) & "'"
    End If
    
    If NulSID(cbSupplier.value) <> "" Then
        Txt = Txt & " and d.[KDSUPP] = '" & NulSID(cbSupplier.value) & "'"
    End If
    
    If Tgl1.value <> "" And Tgl2.value <> "" Then
        Txt = Txt & " and (b.[DATE] BETWEEN '" & Format(Tgl1.value, "yyyy-mm-dd") & "' AND '" & Format(Tgl2.value, "yyyy-mm-dd") & "')"
    End If
    
    If optCreator(0).value Then
        Txt = Txt & " and SUBSTRING(c.KDSUBLET,1,4) = 'NR00'"
    ElseIf optCreator(1).value Then
        Txt = Txt & " and SUBSTRING(c.KDSUBLET,1,4) <> 'NR00'"
    End If
    
    If NulSID(cbDealer.value) <> "" Then
        Txt = Txt & " and a.cabang = '" & NulSID(cbDealer.value) & "'"
    ElseIf gUstat = 0 Then
        Txt = Txt & " and a.cabang = '" & NulSID(gDb) & "'"
    End If
    
    If opCetak(0).value Then
       ' Call IsiTempRepo(txtTemp, vNamaTabel, Txt)
        dealer = ",b.FS_NM_DEALER"
    ElseIf opCetak(1).value Then
       ' Call harianMech
        dealer = ""
    End If
    
    If gUstat = 1 Then
    
        sSQL = "select a.NAMABARANG,sum(a.QTY) as qty,sum(a.HARGA*a.qty) hargajual,sum(a.HARGABELI*a.qty) hargabeli,sum(a.HARGA*a.qty)-sum(a.HARGABELI*a.qty) as total " & IIf(dealer = "", ",'ALL DEALER' as FS_NM_DEALER", dealer)
        sSQL = sSQL & " from  HIST_SUBLETnew a with (nolock)"
        sSQL = sSQL & " left join  T_INVOICEnew b with (nolock) on a.no_wo = b.no_wo and a.noinv = b.noinv and a.CABANG = b.cabang"
        sSQL = sSQL & " left join MSTSUBLET c with (nolock) on c.KDSUBLET = a.[KODEITEM]"
        sSQL = sSQL & " left join MSTSUPPLIER d with (nolock) on d.KDSUPP = c.KDSUPP"
        sSQL = sSQL & " left join MSTSUBLETPROMO e with (nolock) on e.kdpromo = a.kdpromo"
        sSQL = sSQL & " where 1 = 1 " & Txt
        sSQL = sSQL & " group by a.NAMABARANG " & dealer
        
        Set rsLaporan = objO.FetchDataHO(sSQL)
    Else
        sSQL = "select a.NAMABARANG,sum(a.QTY) as qty,sum(a.HARGA*a.qty) hargajual,sum(a.HARGABELI*a.qty) hargabeli,sum(a.HARGA*a.qty)-sum(a.HARGABELI*a.qty) as total " & IIf(dealer = "", ",'ALL DEALER' as FS_NM_DEALER", dealer)
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..HIST_SUBLETnew a with (nolock)"
        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..T_INVOICEnew b with (nolock) on a.no_wo = b.no_wo and a.noinv = b.noinv and a.CABANG = b.cabang"
        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..MSTSUBLET c with (nolock) on c.KDSUBLET = a.[KODEITEM]"
        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..MSTSUPPLIER d with (nolock) on d.KDSUPP = c.KDSUPP"
        sSQL = sSQL & " left join [" & dbReplikasiHO & "]..MSTSUBLETPROMO e with (nolock) on e.kdpromo = a.kdpromo"
        sSQL = sSQL & " where 1 = 1 " & Txt
        sSQL = sSQL & " group by a.NAMABARANG " & dealer
        
        Set rsLaporan = objO.FetchData(sSQL)
    End If
    If Not rsLaporan.EOF Then
        Dim rsRepo As ADODB.Recordset
        Screen.MousePointer = vbHourglass
            
        Set Report = New rptRepoSentralisasiSublet
        Report.database.SetDataSource rsLaporan
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
        Call DisableExport
        Call setLokasiTombol
    Else
        MsgBox "Tidak ada data yang ditampilkan.", vbInformation
    End If
End Sub
 
Private Sub createRSLaporan()
   On Error GoTo createRSLaporan_Error
    If rsLaporan.State Then rsLaporan.Close
    Set rsLaporan = Nothing
    With rsLaporan
        .Fields.Append "NO_POLISI", adVarChar, 12
        .Fields.Append "NO_WO", adVarChar, 8
        .Fields.Append "NOINV", adVarChar, 10
        .Fields.Append "CUSTOMER", adVarChar, 160
        .Fields.Append "NO_CHASSIS", adVarChar, 25
        .Fields.Append "FMDL", adVarChar, 50
        .Fields.Append "FS_NM_DEALER", adVarChar, 30
        .Fields.Append "FRONT", adVarChar, 5
        .Fields.Append "GRP1", adVarChar, 100
        .Fields.Append "RTE1", adDouble
        
        .Fields.Append "ACTRATE1", adDouble
        .Fields.Append "DISC1", adDouble
        .Fields.Append "discPartValid", adDouble
        .Fields.Append "discSublet", adDouble
        .Fields.Append "LABOR", adDouble
        .Fields.Append "part", adDouble
        .Fields.Append "SBLT", adDouble
        .Fields.Append "Total", adDouble
        .Fields.Append "KODETIPE", adDouble
        .Open
    End With
   On Error GoTo 0
   Exit Sub

createRSLaporan_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure createRSLaporan of Form frmRepoSentralSublet")
End Sub

Private Sub refreshPromo()
    sSQL = "SELECT DISTINCT A.KDPROMO,A.NAMAPROMO,A.KETERANGAN,A.HARGAJUAL,A.KDSUBLET "
    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLETPROMO A WITH (NOLOCK)"
'    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET D WITH (NOLOCK) ON A.KDSUBLET = D.KDSUBLET"
'    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMOMDL B WITH (NOLOCK) ON A.KDPROMO = B.KDPROMO"
'    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMODEALER C WITH (NOLOCK) ON A.KDPROMO = C.KDPROMO"
    sSQL = sSQL & " WHERE  A.FS_LOG_DEL = '' "

    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = sSQL
    cbPromo.DataSourceList = Adodc1
    cbPromo.DataFieldList = "KDPROMO"
    
    cbPromo.Columns(0).Width = 2000
    cbPromo.Columns(1).Width = 3500
    cbPromo.Columns(0).Caption = "Kode Promo"
    cbPromo.Columns(1).Caption = "Nama Promo"
    cbPromo.Columns(2).Visible = False
    cbPromo.Columns(3).Visible = False
    cbPromo.Columns(4).Visible = False
    cbPromo.Text = ""
End Sub


Private Sub refreshPekerjaan()
    sSQL = "SELECT DISTINCT D.KDSUBLET,D.NAMASUBLET,D.KETERANGAN,D.HARGAJUAL"
    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLET D "
    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL B WITH (NOLOCK) ON D.KDSUBLET = B.KDSUBLET"
    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER C WITH (NOLOCK) ON D.KDSUBLET = C.KDSUBLET"
    sSQL = sSQL & " WHERE D.FS_LOG_DEL = ''"

    Adodc2.ConnectionString = SQLCONBengkel
    Adodc2.RecordSource = sSQL
    cbPekerjaan.DataSourceList = Adodc2
    cbPekerjaan.DataFieldList = "KDSUBLET"
    
    cbPekerjaan.Columns(0).Width = 1700
    cbPekerjaan.Columns(1).Width = 4000
    cbPekerjaan.Columns(0).Caption = "Kode Sublet"
    cbPekerjaan.Columns(1).Caption = "Nama Sublet"
    cbPekerjaan.Columns(2).Visible = False
    cbPekerjaan.Columns(3).Visible = False
    cbPekerjaan.Columns(4).Visible = False
    cbPekerjaan.Text = ""
End Sub

Private Sub refreshDealer()
    sSQL = "select FS_KD_DEALER,FS_NM_DEALER,DBASECBG from [" & dbReplikasiHO & "]..T_DEALER where FS_LOG_DEL = '' and INTERNAL = 1 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER "
    
    Adodc3.ConnectionString = SQLCONBengkel
    Adodc3.RecordSource = sSQL
    cbDealer.DataSourceList = Adodc3
    cbDealer.DataFieldList = "DBASECBG"
    cbDealer.DataFieldToDisplay = "FS_NM_DEALER"
    
    cbDealer.Columns(0).Width = 1700
    cbDealer.Columns(1).Width = 4000
    cbDealer.Columns(0).Caption = "Kode Dealer"
    cbDealer.Columns(1).Caption = "Nama Dealer"
    cbDealer.Columns(2).Visible = False
    cbDealer.Text = ""
End Sub


Private Sub refreshSupplier()
    sSQL = "select KDSUPP,SUPPLIER,ALAMAT1 from [" & dbReplikasiHO & "]..MSTSUPPLIER with (nolock) where FS_LOG_DEL = ''"
    
    Adodc4.ConnectionString = SQLCONBengkel
    Adodc4.RecordSource = sSQL
    cbSupplier.DataSourceList = Adodc4
    cbSupplier.DataFieldList = "KDSUPP"
    cbSupplier.DataFieldToDisplay = "KDSUPP"
    
    cbSupplier.Columns(0).Width = 1700
    cbSupplier.Columns(1).Width = 4000
    cbSupplier.Columns(0).Caption = "Kode Supplier"
    cbSupplier.Columns(1).Caption = "Nama Supplier"
    cbSupplier.Columns(2).Caption = "Alamat Supplier"
    cbSupplier.Text = ""
End Sub
