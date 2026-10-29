VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MsComCtl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoSBRateRetention 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form Retention Rate"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4440
      TabIndex        =   15
      Top             =   5100
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
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
      MICON           =   "frmRepoSBRateRetention.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   5190
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5685
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
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
      Left            =   6285
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7560
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
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7560
      Width           =   1440
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   2070
      TabIndex        =   14
      Top             =   5715
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7560
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   3000
      TabIndex        =   0
      Top             =   8880
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
   Begin VB.Frame frameVinByVinRetention 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan SB Rate ke 1-9 By Model"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   480
      TabIndex        =   48
      Top             =   480
      Width           =   12375
      Begin VB.Frame frameKolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daftar Kolom"
         Height          =   4695
         Left            =   8040
         TabIndex        =   62
         Top             =   960
         Width           =   4215
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alamat STNK"
            Height          =   375
            Index           =   21
            Left            =   2160
            TabIndex        =   84
            Top             =   3840
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama STNK"
            Height          =   375
            Index           =   20
            Left            =   2160
            TabIndex        =   83
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Telp SPK"
            Height          =   375
            Index           =   19
            Left            =   2160
            TabIndex        =   82
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alamat SPK"
            Height          =   375
            Index           =   18
            Left            =   2160
            TabIndex        =   81
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama SPK"
            Height          =   375
            Index           =   17
            Left            =   2160
            TabIndex        =   80
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cara Bayar"
            Height          =   375
            Index           =   16
            Left            =   2160
            TabIndex        =   79
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Wiraniaga"
            Height          =   375
            Index           =   15
            Left            =   2160
            TabIndex        =   78
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dealer Terakhir SB"
            Height          =   375
            Index           =   14
            Left            =   2160
            TabIndex        =   77
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dealer Penjual"
            Height          =   375
            Index           =   13
            Left            =   2160
            TabIndex        =   76
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Next Service"
            Height          =   375
            Index           =   12
            Left            =   2160
            TabIndex        =   75
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "History SBI / SBE"
            Height          =   375
            Index           =   11
            Left            =   2160
            TabIndex        =   74
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Telpon"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   73
            Top             =   3840
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alamat Customer"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   72
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Model No"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   71
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Retail Sales Date"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delivery Date"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   69
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tahun"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   68
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Rangka"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   67
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Polisi"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama Model"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama Customer"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Reminder"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1815
         End
         Begin AIFCmp1.asxPowerButton cmdTutupDealer 
            Height          =   360
            Left            =   0
            TabIndex        =   85
            Top             =   4320
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Tut&up"
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
            Appearance      =   1
         End
      End
      Begin VB.CheckBox ckKolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua Kolom"
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
         Left            =   9000
         TabIndex        =   61
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdDownloadVin 
         BackColor       =   &H00FF8080&
         Caption         =   "Download Non Retention Data "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1035
         Width           =   3255
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   50
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal delivery"
            Height          =   255
            Index           =   8
            Left            =   5520
            TabIndex        =   53
            Top             =   120
            Width           =   2415
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal invoice"
            Height          =   255
            Index           =   7
            Left            =   105
            TabIndex        =   52
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal retail"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   51
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdProsesVin 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1035
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker tglAwalVin 
         Height          =   375
         Left            =   1410
         TabIndex        =   54
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102432769
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker tglAkhirVin 
         Height          =   375
         Left            =   3000
         TabIndex        =   55
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   102432769
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgVin 
         Height          =   615
         Left            =   480
         TabIndex        =   56
         Top             =   1560
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label12 
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
         Left            =   2880
         TabIndex        =   59
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   58
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   57
         Top             =   1800
         Visible         =   0   'False
         Width           =   6255
      End
   End
   Begin ucXPButton.XPButton btnTutup 
      Height          =   495
      Left            =   11760
      TabIndex        =   47
      Top             =   360
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      BTYPE           =   2
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoSBRateRetention.frx":001C
      PICN            =   "frmRepoSBRateRetention.frx":0038
      PICH            =   "frmRepoSBRateRetention.frx":3360
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExportExcel 
      Height          =   495
      Left            =   480
      TabIndex        =   46
      Top             =   240
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   873
      BTYPE           =   2
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoSBRateRetention.frx":6752
      PICN            =   "frmRepoSBRateRetention.frx":676E
      PICH            =   "frmRepoSBRateRetention.frx":9F9D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   8130
      Left            =   12720
      OleObjectBlob   =   "frmRepoSBRateRetention.frx":DA06
      TabIndex        =   41
      Top             =   390
      Visible         =   0   'False
      Width           =   12510
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan SB Rate 1K - 100K By MODEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   495
      TabIndex        =   28
      Top             =   4800
      Width           =   12375
      Begin VB.CommandButton cmdByModel 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   29
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal retail"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   32
            Top             =   120
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal invoice"
            Height          =   255
            Index           =   3
            Left            =   105
            TabIndex        =   31
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal delivery"
            Height          =   255
            Index           =   5
            Left            =   5520
            TabIndex        =   30
            Top             =   120
            Width           =   2415
         End
      End
      Begin MSComCtl2.DTPicker Tgl1Mdl 
         Height          =   375
         Left            =   1650
         TabIndex        =   34
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   100007937
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2Mdl 
         Height          =   375
         Left            =   3240
         TabIndex        =   35
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   100007937
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgressBarByModel 
         Height          =   615
         Left            =   480
         TabIndex        =   36
         Top             =   1560
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   38
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label2 
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
         TabIndex        =   37
         Top             =   1080
         Width           =   60
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan SB Rate 1K - 100K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   510
      TabIndex        =   16
      Top             =   2175
      Width           =   12375
      Begin VB.CommandButton cmdCetakGrafik 
         BackColor       =   &H00FF8080&
         Caption         =   "Cetak Grafik"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1065
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal delivery"
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   27
            Top             =   120
            Width           =   2415
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal invoice"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal retail"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   25
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdCetakK 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4740
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1050
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Tgl1K 
         Height          =   375
         Left            =   1650
         TabIndex        =   18
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62914561
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2K 
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62980097
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgressBar1K 
         Height          =   615
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
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
         TabIndex        =   23
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   22
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan SB Rate 10K - 50K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   12375
      Begin VB.CommandButton cmdCetak 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   375
         Left            =   1650
         TabIndex        =   5
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   375193601
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   375193601
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblAmbilData 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   6255
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
         TabIndex        =   9
         Top             =   1560
         Width           =   60
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   8
         Top             =   360
         Width           =   840
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
         TabIndex        =   7
         Top             =   360
         Width           =   60
      End
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1890
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1890
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   9075
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   16007
      StartColor      =   33023
      EndColor        =   12648447
      FormatString    =   "   Laporan Retention Rate"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   503
      BackColor       =   255
      HighlightColor  =   255
      ShadowColor     =   65535
   End
   Begin VB.CommandButton Command1 
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
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7410
      Width           =   1890
   End
End
Attribute VB_Name = "frmRepoSBRateRetention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Report         As Object
Private rsTemp      As ADODB.Recordset
Private rsLaporan   As New ADODB.Recordset
Private rsChart     As New ADODB.Recordset
Private rsDownload  As New ADODB.Recordset

Private Rate(20)                As Double
Private cetakRateSameOutlet(20) As Double
Private cetakRateNCM(20)        As Double
Private cetakRateNonNCM(20)     As Double
Private cetakRateTotal(20)      As Double
Private demand(20)              As Double
Private sameOutlet(20)          As Double
Private ncm(20)                 As Double
Private nonNcm(20)              As Double
Private TotalSBE(20)            As Double
Private nilaiTotal1(10)          As Double
Private nilaiTotal2(10)          As Double
Private nilaiTotal3(10)          As Double
Private nilaiTotal4(10)          As Double
Private nilaiTotal5(10)          As Double
Private nilaiTotal6(10)          As Double
Private nilaiTotal7(10)          As Double
Private i                       As Integer

Private Sub clearNilai()
    For i = 0 To 20
        Rate(i) = 0
        cetakRateSameOutlet(i) = 0
        cetakRateNCM(i) = 0
        cetakRateNonNCM(i) = 0
        cetakRateTotal(i) = 0
        demand(i) = 0
        sameOutlet(i) = 0
        ncm(i) = 0
        nonNcm(i) = 0
        TotalSBE(i) = 0
    Next
End Sub

Private Sub createTableLaporan()
    Set rsLaporan = New ADODB.Recordset
    If rsLaporan.State Then rsLaporan.Close
    'rsLaporan.Close
    
    
    With rsLaporan
        .Fields.Append "urut", adInteger
        .Fields.Append "keterangan", adVarChar, 100
        .Fields.Append "judul", adVarChar, 100
        .Fields.Append "model", adVarChar, 100
        .Fields.Append "K1", adVarChar, 100
        .Fields.Append "K10", adVarChar, 100
        .Fields.Append "K20", adVarChar, 100
        .Fields.Append "K30", adVarChar, 100
        .Fields.Append "K40", adVarChar, 100
        .Fields.Append "K50", adVarChar, 100
        .Fields.Append "K60", adVarChar, 100
        .Fields.Append "K70", adVarChar, 100
        .Fields.Append "K80", adVarChar, 100
        .Fields.Append "K90", adVarChar, 100
        .Fields.Append "K100", adVarChar, 100
        .Open
    End With
End Sub

Private Sub createTableLaporanVin()
    Set rsLaporan = New ADODB.Recordset
    If rsLaporan.State Then rsLaporan.Close
    'rsLaporan.Close
    
    
    With rsLaporan
        .Fields.Append "urut", adVarChar, 10
        .Fields.Append "keterangan", adVarChar, 100
        .Fields.Append "judul", adVarChar, 100
        .Fields.Append "model", adVarChar, 100
        .Fields.Append "K1", adVarChar, 100
        .Fields.Append "K10", adVarChar, 100
        .Fields.Append "K20", adVarChar, 100
        .Fields.Append "K30", adVarChar, 100
        .Fields.Append "K40", adVarChar, 100
        .Fields.Append "K50", adVarChar, 100
        .Fields.Append "K60", adVarChar, 100
        .Fields.Append "K70", adVarChar, 100
        .Fields.Append "K80", adVarChar, 100
        .Fields.Append "tot16", adVarChar, 100
        .Fields.Append "tot26", adVarChar, 100
        .Fields.Append "tot19", adVarChar, 100
        .Fields.Append "tot29", adVarChar, 100
        .Open
    End With
End Sub

Private Sub createTableChart()
    Set rsChart = New ADODB.Recordset
    If rsChart.State Then rsChart.Close
    'rsLaporan.Close
    
    
    With rsChart
        .Fields.Append "keterangan", adVarChar, 100
        .Fields.Append "k1", adInteger
        .Fields.Append "k10", adInteger
        .Fields.Append "k20", adInteger
        .Fields.Append "k30", adInteger
        .Fields.Append "k40", adInteger
        .Fields.Append "k50", adInteger
        .Fields.Append "k60", adInteger
        .Fields.Append "k70", adInteger
        .Fields.Append "k80", adInteger
        .Fields.Append "k90", adInteger
        .Fields.Append "k100", adInteger
        .Open
    End With
End Sub

Private Sub createTableDownload(semua As Boolean)
    If rsDownload.State Then rsDownload.Close
  
    With rsDownload
        If kolom(0).value = vbChecked Then
            .Fields.Append "lastReminder", adVarChar, 50
        End If
        If kolom(1).value = vbChecked Then
            .Fields.Append "CUSTOMER", adVarChar, 200
        End If
        If kolom(2).value = vbChecked Then
            .Fields.Append "FMDL", adVarChar, 100
        End If
        If kolom(3).value = vbChecked Then
            .Fields.Append "NOPOLISI", adVarChar, 40
        End If
        If kolom(4).value = vbChecked Then
            .Fields.Append "NORANGKA", adVarChar, 100
        End If
        If kolom(5).value = vbChecked Then
            .Fields.Append "TAHUN", adVarChar, 10
        End If
        If kolom(6).value = vbChecked Then
            .Fields.Append "DDate", adVarChar, 100
        End If
        If kolom(7).value = vbChecked Then
            .Fields.Append "retailSales", adVarChar, 50
        End If
        If kolom(8).value = vbChecked Then
            .Fields.Append "MDL", adVarChar, 100
        End If
        If kolom(9).value = vbChecked Then
            .Fields.Append "ALAMAT", adVarChar, 300
        End If
        .Fields.Append "KELURAHAN", adVarChar, 100
        .Fields.Append "KECAMATAN", adVarChar, 100
        .Fields.Append "KOTA", adVarChar, 100
        .Fields.Append "PROPINSI", adVarChar, 100
        .Fields.Append "KODEPOS", adVarChar, 100
        
        If kolom(10).value = vbChecked Then
            .Fields.Append "PTSN", adVarChar, 100
        End If
        .Fields.Append "HP1", adVarChar, 100
        .Fields.Append "HP2", adVarChar, 100
        .Fields.Append "LASTSB", adVarChar, 100
        
        If kolom(12).value = vbChecked Then
            .Fields.Append "NEXTSB", adVarChar, 100
        End If
        
        .Fields.Append "TGLNEXTSB", adVarChar, 100
        
        If kolom(13).value = vbChecked Then
            .Fields.Append "DEALERPENJUAL", adVarChar, 200
        End If
        
        If kolom(15).value = vbChecked Then
            .Fields.Append "WIRANIAGA", adVarChar, 100
        End If
        .Fields.Append "USENAME", adVarChar, 200
        .Fields.Append "USEHP", adVarChar, 100
        .Fields.Append "TIPELASTSB", adVarChar, 100
        .Fields.Append "LASTSVCSB", adVarChar, 100
        .Fields.Append "UMURLASTSB", adVarChar, 100
        
        If kolom(14).value = vbChecked Then
            .Fields.Append "DEALERLASTSB", adVarChar, 100
        End If
        If semua Then
            If kolom(16).value = vbChecked Then
                .Fields.Append "CARABAYAR", adVarChar, 100
            End If
            If kolom(17).value = vbChecked Then
                .Fields.Append "NAMASPK", adVarChar, 100
            End If
            If kolom(18).value = vbChecked Then
                .Fields.Append "ALAMATSPK", adVarChar, 100
            End If
            .Fields.Append "PROPINSISPK", adVarChar, 100
            .Fields.Append "KOTASPK", adVarChar, 100
            .Fields.Append "KECSPK", adVarChar, 100
            .Fields.Append "KODEPOSSPK", adVarChar, 100
            .Fields.Append "CDMASPK", adVarChar, 200
            
            If kolom(19).value = vbChecked Then
                .Fields.Append "HP1SPK", adVarChar, 100
            End If
            .Fields.Append "HP2SPK", adVarChar, 100
            If kolom(20).value = vbChecked Then
                .Fields.Append "NAMASTNK", adVarChar, 100
            End If
            If kolom(21).value = vbChecked Then
                .Fields.Append "ALAMATSTNK", adVarChar, 100
            End If
            .Fields.Append "PROPINSISTNK", adVarChar, 100
            .Fields.Append "KOTASTNK", adVarChar, 100
            .Fields.Append "KECSTNK", adVarChar, 100
            .Fields.Append "KODEPOSSTNK", adVarChar, 100
        End If
        .Open
    End With
End Sub

Private Sub btnTutup_Click()
    MSChart1.Visible = False
    btnTutup.Visible = False
    cmdExportExcel.Visible = False
End Sub

Private Sub ckKolom_Click()
    If ckKolom = Unchecked Then
        Call aturKolomLevel
        frameKolom.Visible = True
    ElseIf ckKolom = Checked Then
        frameKolom.Visible = False
    End If
End Sub

Private Sub aturKolomLevel()
    For i = 0 To 21
        kolom(i).Enabled = True
        kolom(i).value = vbChecked
        kolom(i).Visible = True
    Next
    
    If gTipeUser = "SA" Or gTipeUser = "MRS" Then
        For i = 16 To 21
            kolom(i).Enabled = False
            kolom(i).value = vbUnchecked
            kolom(i).Visible = False
        Next
    End If
End Sub

Private Sub cmdByModel_Click()
          
1         If Month(Tgl1K) <> Month(Tgl2K) Or year(Tgl2K) <> year(Tgl2K) Then
2             MsgBox "Range periode dalam 1 bulan", vbInformation
3             Exit Sub
4         End If
5         Screen.MousePointer = vbHourglass
          
6         ProgressBarByModel.Min = 0
7         ProgressBarByModel.Max = 10
8         ProgressBarByModel.Visible = True
9         ProgressBarByModel.value = 1
10        DoEvents
          
11        createTableLaporan
12        ProgressBarByModel.value = 2
13        DoEvents
          Dim rsModel As ADODB.Recordset
          Dim rsSBE   As ADODB.Recordset
          Dim rsDemand(10) As ADODB.Recordset
          Dim rsSBETotal As ADODB.Recordset
          Dim nilaiTemp As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10) As Double
          
          
14        If opTipe(3).value = True Then
15            GoTo prosesTanggalService
16        ElseIf opTipe(4).value = True Then
17            GoTo prosesTanggalRetail
18        ElseIf opTipe(5).value = True Then
19            GoTo prosesTanggalPDS
20        End If

prosesTanggalPDS:
21        Set Report = New rptSBERateByModel
22        Call createTableLaporan
          
          ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
23        sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
24        Set rsModel = objO.FetchData(sSQL)
          
25        If Log.RPLHistori = 0 Then
26            sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
27            sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
28            sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
29            sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
30            sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "')"
31            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
32            sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
33            sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
34        Else
35            sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
36            sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
37            sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
38            sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
39            sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "')"
40            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
41            sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
42            sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
43        End If
44        Set rsSBE = objO.FetchData(sSQL)
45        If Not rsModel.EOF Then
46            rsModel.MoveFirst
47            Do While Not rsModel.EOF
48                With rsLaporan
49                    .AddNew
50                    rsSBE.filter = ""
51                    For i = 0 To 10
52                        nilaiTempDemand(i) = 0
53                        nilaiTempTotal(i) = 0
54                    Next
55                    If Not rsSBE.EOF Then
56                        rsSBE.filter = ""
57                        nilaiTemp = 0
58                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
59                        If Not rsSBE.EOF Then
60                            Do While Not rsSBE.EOF
61                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
62                                rsSBE.MoveNext
63                            Loop
64                        End If
65                        !K1 = nilaiTemp
66                        nilaiTempTotal(0) = nilaiTemp
67                        rsSBE.filter = ""
68                        nilaiTemp = 0
                          
69                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
70                        If Not rsSBE.EOF Then
71                            Do While Not rsSBE.EOF
72                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
73                                rsSBE.MoveNext
74                            Loop
75                        End If
76                        !K10 = nilaiTemp
77                        nilaiTempTotal(1) = nilaiTemp
78                        rsSBE.filter = ""
79                        nilaiTemp = 0
                          
80                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
81                        If Not rsSBE.EOF Then
82                            Do While Not rsSBE.EOF
83                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
84                                rsSBE.MoveNext
85                            Loop
86                        End If
87                        !K20 = nilaiTemp
88                        nilaiTempTotal(2) = nilaiTemp
89                        rsSBE.filter = ""
90                        nilaiTemp = 0
                          
91                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
92                        If Not rsSBE.EOF Then
93                            Do While Not rsSBE.EOF
94                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
95                                rsSBE.MoveNext
96                            Loop
97                        End If
98                        !K30 = nilaiTemp
99                        nilaiTempTotal(3) = nilaiTemp
100                       rsSBE.filter = ""
101                       nilaiTemp = 0
                          
102                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
103                       If Not rsSBE.EOF Then
104                           Do While Not rsSBE.EOF
105                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
106                               rsSBE.MoveNext
107                           Loop
108                       End If
109                       !K40 = nilaiTemp
110                       nilaiTempTotal(4) = nilaiTemp
111                       rsSBE.filter = ""
112                       nilaiTemp = 0
                          
113                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
114                       If Not rsSBE.EOF Then
115                           Do While Not rsSBE.EOF
116                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
117                               rsSBE.MoveNext
118                           Loop
119                       End If
120                       !K50 = nilaiTemp
121                       nilaiTempTotal(5) = nilaiTemp
122                       rsSBE.filter = ""
123                       nilaiTemp = 0
                          
124                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
125                       If Not rsSBE.EOF Then
126                           Do While Not rsSBE.EOF
127                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
128                               rsSBE.MoveNext
129                           Loop
130                       End If
131                       !K60 = nilaiTemp
132                       nilaiTempTotal(6) = nilaiTemp
133                       rsSBE.filter = ""
134                       nilaiTemp = 0
                          
135                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
136                       If Not rsSBE.EOF Then
137                           Do While Not rsSBE.EOF
138                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
139                               rsSBE.MoveNext
140                           Loop
141                       End If
142                       !K70 = nilaiTemp
143                       nilaiTempTotal(7) = nilaiTemp
144                       rsSBE.filter = ""
145                       nilaiTemp = 0
                          
146                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
147                       If Not rsSBE.EOF Then
148                           Do While Not rsSBE.EOF
149                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
150                               rsSBE.MoveNext
151                           Loop
152                       End If
153                       !K80 = nilaiTemp
154                       nilaiTempTotal(8) = nilaiTemp
155                       rsSBE.filter = ""
156                       nilaiTemp = 0
                          
157                       rsSBE.filter = "fs_rc = 'R090' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
158                       If Not rsSBE.EOF Then
159                           Do While Not rsSBE.EOF
160                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
161                               rsSBE.MoveNext
162                           Loop
163                       End If
164                       !k90 = nilaiTemp
165                       nilaiTempTotal(9) = nilaiTemp
166                       rsSBE.filter = ""
167                       nilaiTemp = 0
                          
168                       rsSBE.filter = "fs_rc = 'R100' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
169                       If Not rsSBE.EOF Then
170                           Do While Not rsSBE.EOF
171                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
172                               rsSBE.MoveNext
173                           Loop
174                       End If
175                       !k100 = nilaiTemp
176                       nilaiTempTotal(10) = nilaiTemp
177                       rsSBE.filter = ""
178                       nilaiTemp = 0
179                   Else
180                       !K1 = 0
181                       !K10 = 0
182                       !K20 = 0
183                       !K30 = 0
184                       !K40 = 0
185                       !K50 = 0
186                       !K60 = 0
187                       !K70 = 0
188                       !K80 = 0
189                       !k90 = 0
190                       !k100 = 0
191                   End If
192                   !judul = "RETENTION RATE SBE SAME OUTLET"
193                   !model = Nul(rsModel.Fields(0).value)
194                   !Urut = 1
195                   .Update
196               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2017-08-18 || DEMAND
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                   
      '            nilaiTempDemand(0) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 0 and 1) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(1) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(2) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(3) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(4) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(5) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(6) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(7) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(8) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(9) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(10) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")

197               nilaiTempDemand(0) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and (dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
198               nilaiTempDemand(1) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
199               nilaiTempDemand(2) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
200               nilaiTempDemand(3) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
201               nilaiTempDemand(4) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
202               nilaiTempDemand(5) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
203               nilaiTempDemand(6) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
204               nilaiTempDemand(7) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
205               nilaiTempDemand(8) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
206               nilaiTempDemand(9) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
207               nilaiTempDemand(10) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")


      '            nilaiTempDemand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and (dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")

208               With rsLaporan
209                   .AddNew
                
210                   !K1 = nilaiTempDemand(0) '
211                   !K10 = nilaiTempDemand(1)  'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
212                   !K20 = nilaiTempDemand(2)  'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
213                   !K30 = nilaiTempDemand(3)  ' objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
214                   !K40 = nilaiTempDemand(4) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
215                   !K50 = nilaiTempDemand(5) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
216                   !K60 = nilaiTempDemand(6) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
217                   !K70 = nilaiTempDemand(7) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
218                   !K80 = nilaiTempDemand(8) ' objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
219                   !k90 = nilaiTempDemand(9) ' objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
220                   !k100 = nilaiTempDemand(10) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
                       
221                   !judul = "DEMAND RETENTION SBE SAME OUTLET"
222                   !model = Nul(rsModel.Fields(0).value)
223                   !Urut = 2
224                   .Update
225               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
226               With rsLaporan
227                   .AddNew
                      
228                   If CDbl(nilaiTempDemand(0)) > 0 Then
229                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
230                   Else
231                       !K1 = "0 %"
232                   End If
233                   If CDbl(nilaiTempDemand(1)) > 0 Then
234                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
235                   Else
236                       !K10 = "0 %"
237                   End If
238                   If CDbl(nilaiTempDemand(2)) > 0 Then
239                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
240                   Else
241                       !K20 = "0 %"
242                   End If
243                   If CDbl(nilaiTempDemand(3)) > 0 Then
244                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
245                   Else
246                       !K30 = "0 %"
247                   End If
248                   If CDbl(nilaiTempDemand(4)) > 0 Then
249                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
250                   Else
251                       !K40 = "0 %"
252                   End If
253                   If CDbl(nilaiTempDemand(5)) > 0 Then
254                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
255                   Else
256                       !K50 = "0 %"
257                   End If
258                   If CDbl(nilaiTempDemand(6)) > 0 Then
259                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
260                   Else
261                       !K60 = "0 %"
262                   End If
263                   If CDbl(nilaiTempDemand(7)) > 0 Then
264                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
265                   Else
266                       !K70 = "0 %"
267                   End If
268                   If CDbl(nilaiTempDemand(8)) > 0 Then
269                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
270                   Else
271                       !K80 = "0 %"
272                   End If
273                   If CDbl(nilaiTempDemand(9)) > 0 Then
274                       !k90 = Round((nilaiTempTotal(9) / nilaiTempDemand(9)) * 100, 0) & " %"
275                   Else
276                       !k90 = "0 %"
277                   End If
278                   If CDbl(nilaiTempDemand(10)) > 0 Then
279                       !k100 = Round((nilaiTempTotal(10) / nilaiTempDemand(10)) * 100, 0) & " %"
280                   Else
281                       !k100 = "0 %"
282                   End If
                      
283                   !judul = "PERSENTASE RETENTION SBE SAME OUTLET"
284                   !model = Nul(rsModel.Fields(0).value)
285                   !Urut = 3
286                   .Update
287               End With
                  
288               rsModel.MoveNext
289           Loop
290       End If
291       GoTo prosesCetak
          
292       Exit Sub
          
prosesTanggalRetail:
293       Set Report = New rptSBERateByModel
294       Call createTableLaporan
          
          ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
295       sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
296       Set rsModel = objO.FetchData(sSQL)
          
297       If Log.RPLHistori = 0 Then
298           sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
299           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
300           sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
301           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
302           sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' AND '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' )"
303           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
304           sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
305           sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
306       Else
307           sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
308           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
309           sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
310           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
311           sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' AND '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' )"
312           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
313           sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
314           sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
315       End If
316       Set rsSBE = objO.FetchData(sSQL)
317       If Not rsModel.EOF Then
318           rsModel.MoveFirst
319           Do While Not rsModel.EOF
320               With rsLaporan
321                   .AddNew
322                   rsSBE.filter = ""
323                   For i = 0 To 10
324                       nilaiTempDemand(i) = 0
325                       nilaiTempTotal(i) = 0
326                   Next
327                   If Not rsSBE.EOF Then
328                       rsSBE.filter = ""
329                       nilaiTemp = 0
330                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
331                       If Not rsSBE.EOF Then
332                           Do While Not rsSBE.EOF
333                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
334                               rsSBE.MoveNext
335                           Loop
336                       End If
337                       !K1 = nilaiTemp
338                       nilaiTempTotal(0) = nilaiTemp
339                       rsSBE.filter = ""
340                       nilaiTemp = 0
                          
341                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
342                       If Not rsSBE.EOF Then
343                           Do While Not rsSBE.EOF
344                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
345                               rsSBE.MoveNext
346                           Loop
347                       End If
348                       !K10 = nilaiTemp
349                       nilaiTempTotal(1) = nilaiTemp
350                       rsSBE.filter = ""
351                       nilaiTemp = 0
                          
352                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
353                       If Not rsSBE.EOF Then
354                           Do While Not rsSBE.EOF
355                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
356                               rsSBE.MoveNext
357                           Loop
358                       End If
359                       !K20 = nilaiTemp
360                       nilaiTempTotal(2) = nilaiTemp
361                       rsSBE.filter = ""
362                       nilaiTemp = 0
                          
363                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
364                       If Not rsSBE.EOF Then
365                           Do While Not rsSBE.EOF
366                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
367                               rsSBE.MoveNext
368                           Loop
369                       End If
370                       !K30 = nilaiTemp
371                       nilaiTempTotal(3) = nilaiTemp
372                       rsSBE.filter = ""
373                       nilaiTemp = 0
                          
374                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
375                       If Not rsSBE.EOF Then
376                           Do While Not rsSBE.EOF
377                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
378                               rsSBE.MoveNext
379                           Loop
380                       End If
381                       !K40 = nilaiTemp
382                       nilaiTempTotal(4) = nilaiTemp
383                       rsSBE.filter = ""
384                       nilaiTemp = 0
                          
385                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
386                       If Not rsSBE.EOF Then
387                           Do While Not rsSBE.EOF
388                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
389                               rsSBE.MoveNext
390                           Loop
391                       End If
392                       !K50 = nilaiTemp
393                       nilaiTempTotal(5) = nilaiTemp
394                       rsSBE.filter = ""
395                       nilaiTemp = 0
                          
396                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
397                       If Not rsSBE.EOF Then
398                           Do While Not rsSBE.EOF
399                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
400                               rsSBE.MoveNext
401                           Loop
402                       End If
403                       !K60 = nilaiTemp
404                       nilaiTempTotal(6) = nilaiTemp
405                       rsSBE.filter = ""
406                       nilaiTemp = 0
                          
407                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
408                       If Not rsSBE.EOF Then
409                           Do While Not rsSBE.EOF
410                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
411                               rsSBE.MoveNext
412                           Loop
413                       End If
414                       !K70 = nilaiTemp
415                       nilaiTempTotal(7) = nilaiTemp
416                       rsSBE.filter = ""
417                       nilaiTemp = 0
                          
418                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
419                       If Not rsSBE.EOF Then
420                           Do While Not rsSBE.EOF
421                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
422                               rsSBE.MoveNext
423                           Loop
424                       End If
425                       !K80 = nilaiTemp
426                       nilaiTempTotal(8) = nilaiTemp
427                       rsSBE.filter = ""
428                       nilaiTemp = 0
                          
429                       rsSBE.filter = "fs_rc = 'R090' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
430                       If Not rsSBE.EOF Then
431                           Do While Not rsSBE.EOF
432                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
433                               rsSBE.MoveNext
434                           Loop
435                       End If
436                       !k90 = nilaiTemp
437                       nilaiTempTotal(9) = nilaiTemp
438                       rsSBE.filter = ""
439                       nilaiTemp = 0
                          
440                       rsSBE.filter = "fs_rc = 'R100' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
441                       If Not rsSBE.EOF Then
442                           Do While Not rsSBE.EOF
443                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
444                               rsSBE.MoveNext
445                           Loop
446                       End If
447                       !k100 = nilaiTemp
448                       nilaiTempTotal(10) = nilaiTemp
449                       rsSBE.filter = ""
450                       nilaiTemp = 0
451                   Else
452                       !K1 = 0
453                       !K10 = 0
454                       !K20 = 0
455                       !K30 = 0
456                       !K40 = 0
457                       !K50 = 0
458                       !K60 = 0
459                       !K70 = 0
460                       !K80 = 0
461                       !k90 = 0
462                       !k100 = 0
463                   End If
464                   !judul = "RETENTION RATE SBE SAME OUTLET"
465                   !model = Nul(rsModel.Fields(0).value)
466                   !Urut = 1
467                   .Update
468               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2017-08-18 || DEMAND
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                   
      '            nilaiTempDemand(0) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 0 and 1) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(1) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(2) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(3) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(4) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(5) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(6) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(7) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(8) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(9) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
      '            nilaiTempDemand(10) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and a.TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")

      '            nilaiTempDemand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and (dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '            nilaiTempDemand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")

469               nilaiTempDemand(0) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and (dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
470               nilaiTempDemand(1) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
471               nilaiTempDemand(2) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
472               nilaiTempDemand(3) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
473               nilaiTempDemand(4) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
474               nilaiTempDemand(5) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
475               nilaiTempDemand(6) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
476               nilaiTempDemand(7) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
477               nilaiTempDemand(8) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
478               nilaiTempDemand(9) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")
479               nilaiTempDemand(10) = objO.FetchDataSPK_Double("select isnull(sum(total),0) as total from ( select isnull(AVG(a.total),0) as total,isnull(b.fs_grp_mdl,'') model from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun,b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur),b.kode_type ) a left join [" & gDb & "]..model b on b.mdl = a.kode_type where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl ) A")



480               With rsLaporan
481                   .AddNew
                
482                   !K1 = nilaiTempDemand(0) '
483                   !K10 = nilaiTempDemand(1)  'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
484                   !K20 = nilaiTempDemand(2)  'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
485                   !K30 = nilaiTempDemand(3)  ' objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
486                   !K40 = nilaiTempDemand(4) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
487                   !K50 = nilaiTempDemand(5) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
488                   !K60 = nilaiTempDemand(6) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
489                   !K70 = nilaiTempDemand(7) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
490                   !K80 = nilaiTempDemand(8) ' objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
491                   !k90 = nilaiTempDemand(9) ' objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
492                   !k100 = nilaiTempDemand(10) 'objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1Mdl, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and TglFaktur between '" & Format(Tgl1Mdl, "mm/dd/yyyy") & "' and '" & Format(Tgl2Mdl, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
                       
493                   !judul = "DEMAND RETENTION SBE SAME OUTLET"
494                   !model = Nul(rsModel.Fields(0).value)
495                   !Urut = 2
496                   .Update
497               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
498               With rsLaporan
499                   .AddNew
                      
500                   If CDbl(nilaiTempDemand(0)) > 0 Then
501                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
502                   Else
503                       !K1 = "0 %"
504                   End If
505                   If CDbl(nilaiTempDemand(1)) > 0 Then
506                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
507                   Else
508                       !K10 = "0 %"
509                   End If
510                   If CDbl(nilaiTempDemand(2)) > 0 Then
511                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
512                   Else
513                       !K20 = "0 %"
514                   End If
515                   If CDbl(nilaiTempDemand(3)) > 0 Then
516                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
517                   Else
518                       !K30 = "0 %"
519                   End If
520                   If CDbl(nilaiTempDemand(4)) > 0 Then
521                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
522                   Else
523                       !K40 = "0 %"
524                   End If
525                   If CDbl(nilaiTempDemand(5)) > 0 Then
526                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
527                   Else
528                       !K50 = "0 %"
529                   End If
530                   If CDbl(nilaiTempDemand(6)) > 0 Then
531                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
532                   Else
533                       !K60 = "0 %"
534                   End If
535                   If CDbl(nilaiTempDemand(7)) > 0 Then
536                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
537                   Else
538                       !K70 = "0 %"
539                   End If
540                   If CDbl(nilaiTempDemand(8)) > 0 Then
541                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
542                   Else
543                       !K80 = "0 %"
544                   End If
545                   If CDbl(nilaiTempDemand(9)) > 0 Then
546                       !k90 = Round((nilaiTempTotal(9) / nilaiTempDemand(9)) * 100, 0) & " %"
547                   Else
548                       !k90 = "0 %"
549                   End If
550                   If CDbl(nilaiTempDemand(10)) > 0 Then
551                       !k100 = Round((nilaiTempTotal(10) / nilaiTempDemand(10)) * 100, 0) & " %"
552                   Else
553                       !k100 = "0 %"
554                   End If
                      
555                   !judul = "PERSENTASE RETENTION SBE SAME OUTLET"
556                   !model = Nul(rsModel.Fields(0).value)
557                   !Urut = 3
558                   .Update
559               End With
                  
560               rsModel.MoveNext
561           Loop
562       End If
563       GoTo prosesCetak
          
564       Exit Sub
prosesTanggalService:
565       Set Report = New rptSBERateByModel
566       Call createTableLaporan
          
          ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
567       sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
568       Set rsModel = objO.FetchData(sSQL)
          
569       If Log.RPLHistori = 0 Then
570           sSQL = "select COUNT(noinv) total,fs_rc,"
571           sSQL = sSQL & " E.fs_grp_mdl"
572           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
573           sSQL = sSQL & " LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
574           sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
575           sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
576           sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
577           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
578           sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1Mdl.value, "mm/dd/yyyy") & "' AND '" & Format(Tgl2Mdl.value, "mm/dd/yyyy") & "'"
579           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
580           sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
581           sSQL = sSQL & " grOUP BY FS_RC, e.fs_grp_mdl"
582           sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
583       Else
584           sSQL = "select COUNT(noinv) total,fs_rc,"
585           sSQL = sSQL & " E.fs_grp_mdl"
586           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
587           sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
588           sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
589           sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
590           sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
591           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
592           sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1Mdl.value, "mm/dd/yyyy") & "' AND '" & Format(Tgl2Mdl.value, "mm/dd/yyyy") & "'"
593           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
594           sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
595           sSQL = sSQL & " grOUP BY FS_RC, e.fs_grp_mdl"
596           sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
          
597       End If
598       Set rsSBETotal = objO.FetchData(sSQL)
          
599       If Log.RPLHistori = 0 Then
600           sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
601           sSQL = sSQL & " E.fs_grp_mdl"
602           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
603           sSQL = sSQL & " LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
604           sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
605           sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
606           sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
607           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
608           sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1Mdl.value, "mm/dd/yyyy") & "' AND '" & Format(Tgl2Mdl.value, "mm/dd/yyyy") & "'"
609           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
610           sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
611           sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , e.fs_grp_mdl"
612           sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
613       Else
614           sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
615           sSQL = sSQL & " E.fs_grp_mdl"
616           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
617           sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
618           sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
619           sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
620           sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
621           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
622           sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1Mdl.value, "mm/dd/yyyy") & "' AND '" & Format(Tgl2Mdl.value, "mm/dd/yyyy") & "'"
623           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
624           sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
625           sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , e.fs_grp_mdl"
626           sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
627       End If
628       Set rsSBE = objO.FetchData(sSQL)
629       If Not rsModel.EOF Then
630           rsModel.MoveFirst
631           Do While Not rsModel.EOF
                  
632               With rsLaporan
633                   .AddNew
634                   If Not rsSBETotal.EOF Then
635                       rsSBETotal.filter = ""
636                       nilaiTemp = 0
637                       rsSBETotal.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
638                       If Not rsSBETotal.EOF Then
639                           Do While Not rsSBETotal.EOF
640                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
641                               rsSBETotal.MoveNext
642                           Loop
643                       End If
644                       !K1 = nilaiTemp
645                       rsSBETotal.filter = ""
646                       nilaiTemp = 0
                          
647                       rsSBETotal.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
648                       If Not rsSBETotal.EOF Then
649                           Do While Not rsSBETotal.EOF
650                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
651                               rsSBETotal.MoveNext
652                           Loop
653                       End If
654                       !K10 = nilaiTemp
655                       rsSBETotal.filter = ""
656                       nilaiTemp = 0
                          
657                       rsSBETotal.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
658                       If Not rsSBETotal.EOF Then
659                           Do While Not rsSBETotal.EOF
660                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
661                               rsSBETotal.MoveNext
662                           Loop
663                       End If
664                       !K20 = nilaiTemp
665                       rsSBETotal.filter = ""
666                       nilaiTemp = 0
                          
667                       rsSBETotal.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
668                       If Not rsSBETotal.EOF Then
669                           Do While Not rsSBETotal.EOF
670                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
671                               rsSBETotal.MoveNext
672                           Loop
673                       End If
674                       !K30 = nilaiTemp
675                       rsSBETotal.filter = ""
676                       nilaiTemp = 0
                          
677                       rsSBETotal.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
678                       If Not rsSBETotal.EOF Then
679                           Do While Not rsSBETotal.EOF
680                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
681                               rsSBETotal.MoveNext
682                           Loop
683                       End If
684                       !K40 = nilaiTemp
685                       rsSBETotal.filter = ""
686                       nilaiTemp = 0
                          
687                       rsSBETotal.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
688                       If Not rsSBETotal.EOF Then
689                           Do While Not rsSBETotal.EOF
690                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
691                               rsSBETotal.MoveNext
692                           Loop
693                       End If
694                       !K50 = nilaiTemp
695                       rsSBETotal.filter = ""
696                       nilaiTemp = 0
                          
697                       rsSBETotal.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
698                       If Not rsSBETotal.EOF Then
699                           Do While Not rsSBETotal.EOF
700                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
701                               rsSBETotal.MoveNext
702                           Loop
703                       End If
704                       !K60 = nilaiTemp
705                       rsSBETotal.filter = ""
706                       nilaiTemp = 0
                          
707                       rsSBETotal.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
708                       If Not rsSBETotal.EOF Then
709                           Do While Not rsSBETotal.EOF
710                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
711                               rsSBETotal.MoveNext
712                           Loop
713                       End If
714                       !K70 = nilaiTemp
715                       rsSBETotal.filter = ""
716                       nilaiTemp = 0
                          
717                       rsSBETotal.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
718                       If Not rsSBETotal.EOF Then
719                           Do While Not rsSBETotal.EOF
720                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
721                               rsSBETotal.MoveNext
722                           Loop
723                       End If
724                       !K80 = nilaiTemp
725                       rsSBETotal.filter = ""
726                       nilaiTemp = 0
                          
727                       rsSBETotal.filter = "fs_rc = 'R090' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
728                       If Not rsSBETotal.EOF Then
729                           Do While Not rsSBETotal.EOF
730                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
731                               rsSBETotal.MoveNext
732                           Loop
733                       End If
734                       !k90 = nilaiTemp
735                       rsSBETotal.filter = ""
736                       nilaiTemp = 0
                          
737                       rsSBETotal.filter = "fs_rc = 'R100' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
738                       If Not rsSBETotal.EOF Then
739                           Do While Not rsSBETotal.EOF
740                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
741                               rsSBETotal.MoveNext
742                           Loop
743                       End If
744                       !k100 = nilaiTemp
745                       rsSBETotal.filter = ""
746                       nilaiTemp = 0
747                   Else
748                       !K1 = 0
749                       !K10 = 0
750                       !K20 = 0
751                       !K30 = 0
752                       !K40 = 0
753                       !K50 = 0
754                       !K60 = 0
755                       !K70 = 0
756                       !K80 = 0
757                       !k90 = 0
758                       !k100 = 0
759                   End If
760                   !judul = "RETENTION RATE SBE TOTAL"
761                   !model = Nul(rsModel.Fields(0).value)
762                   !Urut = 1
763                   .Update
764               End With
                  
765               With rsLaporan
766                   .AddNew
767                   rsSBE.filter = ""
768                   If Not rsSBE.EOF Then
769                       rsSBE.filter = ""
770                       nilaiTemp = 0
771                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
772                       If Not rsSBE.EOF Then
773                           Do While Not rsSBE.EOF
774                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
775                               rsSBE.MoveNext
776                           Loop
777                       End If
778                       !K1 = nilaiTemp
779                       rsSBE.filter = ""
780                       nilaiTemp = 0
                          
781                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
782                       If Not rsSBE.EOF Then
783                           Do While Not rsSBE.EOF
784                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
785                               rsSBE.MoveNext
786                           Loop
787                       End If
788                       !K10 = nilaiTemp
789                       rsSBE.filter = ""
790                       nilaiTemp = 0
                          
791                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
792                       If Not rsSBE.EOF Then
793                           Do While Not rsSBE.EOF
794                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
795                               rsSBE.MoveNext
796                           Loop
797                       End If
798                       !K20 = nilaiTemp
799                       rsSBE.filter = ""
800                       nilaiTemp = 0
                          
801                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
802                       If Not rsSBE.EOF Then
803                           Do While Not rsSBE.EOF
804                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
805                               rsSBE.MoveNext
806                           Loop
807                       End If
808                       !K30 = nilaiTemp
809                       rsSBE.filter = ""
810                       nilaiTemp = 0
                          
811                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
812                       If Not rsSBE.EOF Then
813                           Do While Not rsSBE.EOF
814                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
815                               rsSBE.MoveNext
816                           Loop
817                       End If
818                       !K40 = nilaiTemp
819                       rsSBE.filter = ""
820                       nilaiTemp = 0
                          
821                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
822                       If Not rsSBE.EOF Then
823                           Do While Not rsSBE.EOF
824                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
825                               rsSBE.MoveNext
826                           Loop
827                       End If
828                       !K50 = nilaiTemp
829                       rsSBE.filter = ""
830                       nilaiTemp = 0
                          
831                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
832                       If Not rsSBE.EOF Then
833                           Do While Not rsSBE.EOF
834                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
835                               rsSBE.MoveNext
836                           Loop
837                       End If
838                       !K60 = nilaiTemp
839                       rsSBE.filter = ""
840                       nilaiTemp = 0
                          
841                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
842                       If Not rsSBE.EOF Then
843                           Do While Not rsSBE.EOF
844                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
845                               rsSBE.MoveNext
846                           Loop
847                       End If
848                       !K70 = nilaiTemp
849                       rsSBE.filter = ""
850                       nilaiTemp = 0
                          
851                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
852                       If Not rsSBE.EOF Then
853                           Do While Not rsSBE.EOF
854                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
855                               rsSBE.MoveNext
856                           Loop
857                       End If
858                       !K80 = nilaiTemp
859                       rsSBE.filter = ""
860                       nilaiTemp = 0
                          
861                       rsSBE.filter = "fs_rc = 'R090' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
862                       If Not rsSBE.EOF Then
863                           Do While Not rsSBE.EOF
864                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
865                               rsSBE.MoveNext
866                           Loop
867                       End If
868                       !k90 = nilaiTemp
869                       rsSBE.filter = ""
870                       nilaiTemp = 0
                          
871                       rsSBE.filter = "fs_rc = 'R100' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
872                       If Not rsSBE.EOF Then
873                           Do While Not rsSBE.EOF
874                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
875                               rsSBE.MoveNext
876                           Loop
877                       End If
878                       !k100 = nilaiTemp
879                       rsSBE.filter = ""
880                       nilaiTemp = 0
881                   Else
882                       !K1 = 0
883                       !K10 = 0
884                       !K20 = 0
885                       !K30 = 0
886                       !K40 = 0
887                       !K50 = 0
888                       !K60 = 0
889                       !K70 = 0
890                       !K80 = 0
891                       !k90 = 0
892                       !k100 = 0
893                   End If
894                   !judul = "RETENTION RATE SBE SAME OUTLET"
895                   !model = Nul(rsModel.Fields(0).value)
896                   !Urut = 4
897                   .Update
898               End With
                  
899               With rsLaporan
900                   .AddNew
901                   rsSBE.filter = ""
902                   If Not rsSBE.EOF Then
903                       rsSBE.filter = ""
904                       nilaiTemp = 0
905                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
906                       If Not rsSBE.EOF Then
907                           Do While Not rsSBE.EOF
908                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
909                               rsSBE.MoveNext
910                           Loop
911                       End If
912                       !K1 = nilaiTemp
913                       rsSBE.filter = ""
914                       nilaiTemp = 0
                          
915                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
916                       If Not rsSBE.EOF Then
917                           Do While Not rsSBE.EOF
918                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
919                               rsSBE.MoveNext
920                           Loop
921                       End If
922                       !K10 = nilaiTemp
923                       rsSBE.filter = ""
924                       nilaiTemp = 0
                          
925                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
926                       If Not rsSBE.EOF Then
927                           Do While Not rsSBE.EOF
928                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
929                               rsSBE.MoveNext
930                           Loop
931                       End If
932                       !K20 = nilaiTemp
933                       rsSBE.filter = ""
934                       nilaiTemp = 0
                          
935                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
936                       If Not rsSBE.EOF Then
937                           Do While Not rsSBE.EOF
938                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
939                               rsSBE.MoveNext
940                           Loop
941                       End If
942                       !K30 = nilaiTemp
943                       rsSBE.filter = ""
944                       nilaiTemp = 0
                          
945                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
946                       If Not rsSBE.EOF Then
947                           Do While Not rsSBE.EOF
948                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
949                               rsSBE.MoveNext
950                           Loop
951                       End If
952                       !K40 = nilaiTemp
953                       rsSBE.filter = ""
954                       nilaiTemp = 0
                          
955                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
956                       If Not rsSBE.EOF Then
957                           Do While Not rsSBE.EOF
958                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
959                               rsSBE.MoveNext
960                           Loop
961                       End If
962                       !K50 = nilaiTemp
963                       rsSBE.filter = ""
964                       nilaiTemp = 0
                          
965                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
966                       If Not rsSBE.EOF Then
967                           Do While Not rsSBE.EOF
968                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
969                               rsSBE.MoveNext
970                           Loop
971                       End If
972                       !K60 = nilaiTemp
973                       rsSBE.filter = ""
974                       nilaiTemp = 0
                          
975                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
976                       If Not rsSBE.EOF Then
977                           Do While Not rsSBE.EOF
978                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
979                               rsSBE.MoveNext
980                           Loop
981                       End If
982                       !K70 = nilaiTemp
983                       rsSBE.filter = ""
984                       nilaiTemp = 0
                          
985                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
986                       If Not rsSBE.EOF Then
987                           Do While Not rsSBE.EOF
988                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
989                               rsSBE.MoveNext
990                           Loop
991                       End If
992                       !K80 = nilaiTemp
993                       rsSBE.filter = ""
994                       nilaiTemp = 0
                          
995                       rsSBE.filter = "fs_rc = 'R090' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
996                       If Not rsSBE.EOF Then
997                           Do While Not rsSBE.EOF
998                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
999                               rsSBE.MoveNext
1000                          Loop
1001                      End If
1002                      !k90 = nilaiTemp
1003                      rsSBE.filter = ""
1004                      nilaiTemp = 0
                          
1005                      rsSBE.filter = "fs_rc = 'R100' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1006                      If Not rsSBE.EOF Then
1007                          Do While Not rsSBE.EOF
1008                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1009                              rsSBE.MoveNext
1010                          Loop
1011                      End If
1012                      !k100 = nilaiTemp
1013                      rsSBE.filter = ""
1014                      nilaiTemp = 0
1015                  Else
1016                      !K1 = 0
1017                      !K10 = 0
1018                      !K20 = 0
1019                      !K30 = 0
1020                      !K40 = 0
1021                      !K50 = 0
1022                      !K60 = 0
1023                      !K70 = 0
1024                      !K80 = 0
1025                      !k90 = 0
1026                      !k100 = 0
1027                  End If
1028                  !judul = "RETENTION RATE SBE NCM"
1029                  !model = Nul(rsModel.Fields(0).value)
1030                  !Urut = 6
1031                  .Update
1032              End With
                  
1033              With rsLaporan
1034                  .AddNew
1035                  rsSBE.filter = ""
1036                  If Not rsSBE.EOF Then
1037                      rsSBE.filter = ""
1038                      nilaiTemp = 0
1039                      rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1040                      If Not rsSBE.EOF Then
1041                          Do While Not rsSBE.EOF
1042                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1043                              rsSBE.MoveNext
1044                          Loop
1045                      End If
1046                      !K1 = nilaiTemp
1047                      rsSBE.filter = ""
1048                      nilaiTemp = 0
                          
1049                      rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1050                      If Not rsSBE.EOF Then
1051                          Do While Not rsSBE.EOF
1052                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1053                              rsSBE.MoveNext
1054                          Loop
1055                      End If
1056                      !K10 = nilaiTemp
1057                      rsSBE.filter = ""
1058                      nilaiTemp = 0
                          
1059                      rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1060                      If Not rsSBE.EOF Then
1061                          Do While Not rsSBE.EOF
1062                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1063                              rsSBE.MoveNext
1064                          Loop
1065                      End If
1066                      !K20 = nilaiTemp
1067                      rsSBE.filter = ""
1068                      nilaiTemp = 0
                          
1069                      rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1070                      If Not rsSBE.EOF Then
1071                          Do While Not rsSBE.EOF
1072                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1073                              rsSBE.MoveNext
1074                          Loop
1075                      End If
1076                      !K30 = nilaiTemp
1077                      rsSBE.filter = ""
1078                      nilaiTemp = 0
                          
1079                      rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1080                      If Not rsSBE.EOF Then
1081                          Do While Not rsSBE.EOF
1082                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1083                              rsSBE.MoveNext
1084                          Loop
1085                      End If
1086                      !K40 = nilaiTemp
1087                      rsSBE.filter = ""
1088                      nilaiTemp = 0
                          
1089                      rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1090                      If Not rsSBE.EOF Then
1091                          Do While Not rsSBE.EOF
1092                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1093                              rsSBE.MoveNext
1094                          Loop
1095                      End If
1096                      !K50 = nilaiTemp
1097                      rsSBE.filter = ""
1098                      nilaiTemp = 0
                          
1099                      rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1100                      If Not rsSBE.EOF Then
1101                          Do While Not rsSBE.EOF
1102                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1103                              rsSBE.MoveNext
1104                          Loop
1105                      End If
1106                      !K60 = nilaiTemp
1107                      rsSBE.filter = ""
1108                      nilaiTemp = 0
                          
1109                      rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1110                      If Not rsSBE.EOF Then
1111                          Do While Not rsSBE.EOF
1112                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1113                              rsSBE.MoveNext
1114                          Loop
1115                      End If
1116                      !K70 = nilaiTemp
1117                      rsSBE.filter = ""
1118                      nilaiTemp = 0
                          
1119                      rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1120                      If Not rsSBE.EOF Then
1121                          Do While Not rsSBE.EOF
1122                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1123                              rsSBE.MoveNext
1124                          Loop
1125                      End If
1126                      !K80 = nilaiTemp
1127                      rsSBE.filter = ""
1128                      nilaiTemp = 0
                          
1129                      rsSBE.filter = "fs_rc = 'R090' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1130                      If Not rsSBE.EOF Then
1131                          Do While Not rsSBE.EOF
1132                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1133                              rsSBE.MoveNext
1134                          Loop
1135                      End If
1136                      !k90 = nilaiTemp
1137                      rsSBE.filter = ""
1138                      nilaiTemp = 0
                          
1139                      rsSBE.filter = "fs_rc = 'R100' and keterangan = 'Non NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
1140                      If Not rsSBE.EOF Then
1141                          Do While Not rsSBE.EOF
1142                              nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
1143                              rsSBE.MoveNext
1144                          Loop
1145                      End If
1146                      !k100 = nilaiTemp
1147                      rsSBE.filter = ""
1148                      nilaiTemp = 0
1149                  Else
1150                      !K1 = 0
1151                      !K10 = 0
1152                      !K20 = 0
1153                      !K30 = 0
1154                      !K40 = 0
1155                      !K50 = 0
1156                      !K60 = 0
1157                      !K70 = 0
1158                      !K80 = 0
1159                      !k90 = 0
1160                      !k100 = 0
1161                  End If
1162                  !judul = "RETENTION RATE SBE Non NCM"
1163                  !model = Nul(rsModel.Fields(0).value)
1164                  !Urut = 8
1165                  .Update
1166              End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2017-08-18 || DEMAND
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
1167              With rsLaporan
1168                  .AddNew
                  
1169                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1170                  !K1 = objO.FetchDataSPK_Double(sSQL)
                      
1171                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1172                  !K10 = objO.FetchDataSPK_Double(sSQL)
                      
1173                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1174                  !K20 = objO.FetchDataSPK_Double(sSQL)
                      
1175                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1176                  !K30 = objO.FetchDataSPK_Double(sSQL)
                      
1177                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1178                  !K40 = objO.FetchDataSPK_Double(sSQL)
                      
1179                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1180                  !K50 = objO.FetchDataSPK_Double(sSQL)
                      
1181                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1182                  !K60 = objO.FetchDataSPK_Double(sSQL)
                      
1183                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1184                  !K70 = objO.FetchDataSPK_Double(sSQL)
                      
1185                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1186                  !K80 = objO.FetchDataSPK_Double(sSQL)
                      
1187                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-52,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1188                  !k90 = objO.FetchDataSPK_Double(sSQL)
                      
1189                  sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1Mdl), "0#") & "/01/" & year(Tgl1Mdl) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-58,'" & Format(Month(Tgl2Mdl), "0#") & "/01/" & year(Tgl2Mdl) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
1190                  !k100 = objO.FetchDataSPK_Double(sSQL)
                       
1191                  !judul = "DEMAND RETENTION RATE TOTAL"
1192                  !model = Nul(rsModel.Fields(0).value)
1193                  !Urut = 2
1194                  .Update
1195              End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  ''-----------------------------------------------------------
                  ''2017-08-21 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
1196              For i = 0 To 10
1197                  nilaiTempDemand(i) = 0
1198                  nilaiTempTotal(i) = 0
1199              Next
1200              rsLaporan.filter = ""
1201              If Not rsLaporan.EOF Then
1202                  rsLaporan.filter = "judul = 'DEMAND RETENTION RATE TOTAL' and model = '" & Nul(rsModel.Fields(0).value) & "'"
1203                  If Not rsLaporan.EOF Then
1204                      If rsLaporan!K1 > 0 Then
1205                          nilaiTempDemand(0) = rsLaporan!K1
1206                      End If
1207                      If rsLaporan!K10 > 0 Then
1208                          nilaiTempDemand(1) = rsLaporan!K10
1209                      End If
1210                      If rsLaporan!K20 > 0 Then
1211                          nilaiTempDemand(2) = rsLaporan!K20
1212                      End If
1213                      If rsLaporan!K30 > 0 Then
1214                          nilaiTempDemand(3) = rsLaporan!K30
1215                      End If
1216                      If rsLaporan!K40 > 0 Then
1217                          nilaiTempDemand(4) = rsLaporan!K40
1218                      End If
1219                      If rsLaporan!K50 > 0 Then
1220                          nilaiTempDemand(5) = rsLaporan!K50
1221                      End If
1222                      If rsLaporan!K60 > 0 Then
1223                          nilaiTempDemand(6) = rsLaporan!K60
1224                      End If
1225                      If rsLaporan!K70 > 0 Then
1226                          nilaiTempDemand(7) = rsLaporan!K70
1227                      End If
1228                      If rsLaporan!K80 > 0 Then
1229                          nilaiTempDemand(8) = rsLaporan!K80
1230                      End If
1231                      If rsLaporan!k90 > 0 Then
1232                          nilaiTempDemand(9) = rsLaporan!k90
1233                      End If
1234                      If rsLaporan!k100 > 0 Then
1235                          nilaiTempDemand(10) = rsLaporan!k100
1236                      End If
1237                  End If
1238              End If
1239              rsLaporan.filter = ""
                  
1240              If Not rsLaporan.EOF Then
1241                  rsLaporan.filter = "judul = 'RETENTION RATE SBE TOTAL' and model = '" & Nul(rsModel.Fields(0).value) & "'"
1242                  If Not rsLaporan.EOF Then
1243                      If rsLaporan!K1 > 0 Then
1244                          nilaiTempTotal(0) = rsLaporan!K1
1245                      End If
1246                      If rsLaporan!K10 > 0 Then
1247                          nilaiTempTotal(1) = rsLaporan!K10
1248                      End If
1249                      If rsLaporan!K20 > 0 Then
1250                          nilaiTempTotal(2) = rsLaporan!K20
1251                      End If
1252                      If rsLaporan!K30 > 0 Then
1253                          nilaiTempTotal(3) = rsLaporan!K30
1254                      End If
1255                      If rsLaporan!K40 > 0 Then
1256                          nilaiTempTotal(4) = rsLaporan!K40
1257                      End If
1258                      If rsLaporan!K50 > 0 Then
1259                          nilaiTempTotal(5) = rsLaporan!K50
1260                      End If
1261                      If rsLaporan!K60 > 0 Then
1262                          nilaiTempTotal(6) = rsLaporan!K60
1263                      End If
1264                      If rsLaporan!K70 > 0 Then
1265                          nilaiTempTotal(7) = rsLaporan!K70
1266                      End If
1267                      If rsLaporan!K80 > 0 Then
1268                          nilaiTempTotal(8) = rsLaporan!K80
1269                      End If
1270                      If rsLaporan!k90 > 0 Then
1271                          nilaiTempTotal(9) = rsLaporan!k90
1272                      End If
1273                      If rsLaporan!k100 > 0 Then
1274                          nilaiTempTotal(10) = rsLaporan!k100
1275                      End If
1276                  End If
1277              End If
                  
1278              With rsLaporan
1279                  .AddNew
1280                  If nilaiTempDemand(0) > 0 Then
1281                      !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
1282                  Else
1283                      !K1 = 0 & " %"
1284                  End If
1285                  If nilaiTempDemand(1) > 0 Then
1286                      !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
1287                  Else
1288                      !K10 = 0 & " %"
1289                  End If
1290                  If nilaiTempDemand(2) > 0 Then
1291                      !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
1292                  Else
1293                      !K20 = 0 & " %"
1294                  End If
1295                  If nilaiTempDemand(3) > 0 Then
1296                      !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
1297                  Else
1298                      !K30 = 0 & " %"
1299                  End If
1300                  If nilaiTempDemand(4) > 0 Then
1301                      !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
1302                  Else
1303                      !K40 = 0 & " %"
1304                  End If
1305                  If nilaiTempDemand(5) > 0 Then
1306                      !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
1307                  Else
1308                      !K50 = 0 & " %"
1309                  End If
1310                  If nilaiTempDemand(6) > 0 Then
1311                      !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
1312                  Else
1313                      !K60 = 0 & " %"
1314                  End If
1315                  If nilaiTempDemand(7) > 0 Then
1316                      !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
1317                  Else
1318                      !K70 = 0 & " %"
1319                  End If
1320                  If nilaiTempDemand(8) > 0 Then
1321                      !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
1322                  Else
1323                      !K80 = 0 & " %"
1324                  End If
1325                  If nilaiTempDemand(9) > 0 Then
1326                      !k90 = Round((nilaiTempTotal(9) / nilaiTempDemand(9)) * 100, 0) & " %"
1327                  Else
1328                      !k90 = 0 & " %"
1329                  End If
1330                  If nilaiTempDemand(10) > 0 Then
1331                      !k100 = Round((nilaiTempTotal(10) / nilaiTempDemand(10)) * 100, 0) & " %"
1332                  Else
1333                      !k100 = 0 & " %"
1334                  End If
1335                  !Urut = 3
1336                  !judul = "PERSENTASE RETENTION RATE TOTAL"
1337                  !model = Nul(rsModel.Fields(0).value)
1338                  .Update
1339              End With
                  
1340              For i = 0 To 10
1341                  nilaiTempTotal(i) = 0
1342              Next
1343              rsLaporan.filter = ""
                  
1344              If Not rsLaporan.EOF Then
1345                  rsLaporan.filter = "judul = 'RETENTION RATE SBE SAME OUTLET' and model = '" & Nul(rsModel.Fields(0).value) & "'"
1346                  If Not rsLaporan.EOF Then
1347                      If rsLaporan!K1 > 0 Then
1348                          nilaiTempTotal(0) = rsLaporan!K1
1349                      End If
1350                      If rsLaporan!K10 > 0 Then
1351                          nilaiTempTotal(1) = rsLaporan!K10
1352                      End If
1353                      If rsLaporan!K20 > 0 Then
1354                          nilaiTempTotal(2) = rsLaporan!K20
1355                      End If
1356                      If rsLaporan!K30 > 0 Then
1357                          nilaiTempTotal(3) = rsLaporan!K30
1358                      End If
1359                      If rsLaporan!K40 > 0 Then
1360                          nilaiTempTotal(4) = rsLaporan!K40
1361                      End If
1362                      If rsLaporan!K50 > 0 Then
1363                          nilaiTempTotal(5) = rsLaporan!K50
1364                      End If
1365                      If rsLaporan!K60 > 0 Then
1366                          nilaiTempTotal(6) = rsLaporan!K60
1367                      End If
1368                      If rsLaporan!K70 > 0 Then
1369                          nilaiTempTotal(7) = rsLaporan!K70
1370                      End If
1371                      If rsLaporan!K80 > 0 Then
1372                          nilaiTempTotal(8) = rsLaporan!K80
1373                      End If
1374                      If rsLaporan!k90 > 0 Then
1375                          nilaiTempTotal(9) = rsLaporan!k90
1376                      End If
1377                      If rsLaporan!k100 > 0 Then
1378                          nilaiTempTotal(10) = rsLaporan!k100
1379                      End If
1380                  End If
1381              End If
                  
1382              With rsLaporan
1383                  .AddNew
1384                  If nilaiTempDemand(0) > 0 Then
1385                      !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
1386                  Else
1387                      !K1 = 0 & " %"
1388                  End If
1389                  If nilaiTempDemand(1) > 0 Then
1390                      !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
1391                  Else
1392                      !K10 = 0 & " %"
1393                  End If
1394                  If nilaiTempDemand(2) > 0 Then
1395                      !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
1396                  Else
1397                      !K20 = 0 & " %"
1398                  End If
1399                  If nilaiTempDemand(3) > 0 Then
1400                      !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
1401                  Else
1402                      !K30 = 0 & " %"
1403                  End If
1404                  If nilaiTempDemand(4) > 0 Then
1405                      !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
1406                  Else
1407                      !K40 = 0 & " %"
1408                  End If
1409                  If nilaiTempDemand(5) > 0 Then
1410                      !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
1411                  Else
1412                      !K50 = 0 & " %"
1413                  End If
1414                  If nilaiTempDemand(6) > 0 Then
1415                      !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
1416                  Else
1417                      !K60 = 0 & " %"
1418                  End If
1419                  If nilaiTempDemand(7) > 0 Then
1420                      !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
1421                  Else
1422                      !K70 = 0 & " %"
1423                  End If
1424                  If nilaiTempDemand(8) > 0 Then
1425                      !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
1426                  Else
1427                      !K80 = 0 & " %"
1428                  End If
1429                  If nilaiTempDemand(9) > 0 Then
1430                      !k90 = Round((nilaiTempTotal(9) / nilaiTempDemand(9)) * 100, 0) & " %"
1431                  Else
1432                      !k90 = 0 & " %"
1433                  End If
1434                  If nilaiTempDemand(10) > 0 Then
1435                      !k100 = Round((nilaiTempTotal(10) / nilaiTempDemand(10)) * 100, 0) & " %"
1436                  Else
1437                      !k100 = 0 & " %"
1438                  End If
1439                  !Urut = 5
1440                  !judul = "PERSENTASE RETENTION RATE SBE SAME OUTLET"
1441                  !model = Nul(rsModel.Fields(0).value)
1442                  .Update
1443              End With
                   
1444              For i = 0 To 10
1445                  nilaiTempTotal(i) = 0
1446              Next
1447              rsLaporan.filter = ""
                  
1448              If Not rsLaporan.EOF Then
1449                  rsLaporan.filter = "judul = 'RETENTION RATE SBE NCM' and model = '" & Nul(rsModel.Fields(0).value) & "'"
1450                  If Not rsLaporan.EOF Then
1451                      If rsLaporan!K1 > 0 Then
1452                          nilaiTempTotal(0) = rsLaporan!K1
1453                      End If
1454                      If rsLaporan!K10 > 0 Then
1455                          nilaiTempTotal(1) = rsLaporan!K10
1456                      End If
1457                      If rsLaporan!K20 > 0 Then
1458                          nilaiTempTotal(2) = rsLaporan!K20
1459                      End If
1460                      If rsLaporan!K30 > 0 Then
1461                          nilaiTempTotal(3) = rsLaporan!K30
1462                      End If
1463                      If rsLaporan!K40 > 0 Then
1464                          nilaiTempTotal(4) = rsLaporan!K40
1465                      End If
1466                      If rsLaporan!K50 > 0 Then
1467                          nilaiTempTotal(5) = rsLaporan!K50
1468                      End If
1469                      If rsLaporan!K60 > 0 Then
1470                          nilaiTempTotal(6) = rsLaporan!K60
1471                      End If
1472                      If rsLaporan!K70 > 0 Then
1473                          nilaiTempTotal(7) = rsLaporan!K70
1474                      End If
1475                      If rsLaporan!K80 > 0 Then
1476                          nilaiTempTotal(8) = rsLaporan!K80
1477                      End If
1478                      If rsLaporan!k90 > 0 Then
1479                          nilaiTempTotal(9) = rsLaporan!k90
1480                      End If
1481                      If rsLaporan!k100 > 0 Then
1482                          nilaiTempTotal(10) = rsLaporan!k100
1483                      End If
1484                  End If
1485              End If
                  
1486              With rsLaporan
1487                  .AddNew
1488                  If nilaiTempDemand(0) > 0 Then
1489                      !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
1490                  Else
1491                      !K1 = 0 & " %"
1492                  End If
1493                  If nilaiTempDemand(1) > 0 Then
1494                      !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
1495                  Else
1496                      !K10 = 0 & " %"
1497                  End If
1498                  If nilaiTempDemand(2) > 0 Then
1499                      !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
1500                  Else
1501                      !K20 = 0 & " %"
1502                  End If
1503                  If nilaiTempDemand(3) > 0 Then
1504                      !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
1505                  Else
1506                      !K30 = 0 & " %"
1507                  End If
1508                  If nilaiTempDemand(4) > 0 Then
1509                      !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
1510                  Else
1511                      !K40 = 0 & " %"
1512                  End If
1513                  If nilaiTempDemand(5) > 0 Then
1514                      !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
1515                  Else
1516                      !K50 = 0 & " %"
1517                  End If
1518                  If nilaiTempDemand(6) > 0 Then
1519                      !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
1520                  Else
1521                      !K60 = 0 & " %"
1522                  End If
1523                  If nilaiTempDemand(7) > 0 Then
1524                      !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
1525                  Else
1526                      !K70 = 0 & " %"
1527                  End If
1528                  If nilaiTempDemand(8) > 0 Then
1529                      !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
1530                  Else
1531                      !K80 = 0 & " %"
1532                  End If
1533                  If nilaiTempDemand(9) > 0 Then
1534                      !k90 = Round((nilaiTempTotal(9) / nilaiTempDemand(9)) * 100, 0) & " %"
1535                  Else
1536                      !k90 = 0 & " %"
1537                  End If
1538                  If nilaiTempDemand(10) > 0 Then
1539                      !k100 = Round((nilaiTempTotal(10) / nilaiTempDemand(10)) * 100, 0) & " %"
1540                  Else
1541                      !k100 = 0 & " %"
1542                  End If
1543                  !Urut = 7
1544                  !judul = "PERSENTASE RETENTION RATE SBE NCM"
1545                  !model = Nul(rsModel.Fields(0).value)
1546                  .Update
1547              End With
                  
                  
1548              For i = 0 To 10
1549                  nilaiTempTotal(i) = 0
1550              Next
1551              rsLaporan.filter = ""
                  
1552              If Not rsLaporan.EOF Then
1553                  rsLaporan.filter = "judul = 'RETENTION RATE SBE Non NCM' and model = '" & Nul(rsModel.Fields(0).value) & "'"
1554                  If Not rsLaporan.EOF Then
1555                      If rsLaporan!K1 > 0 Then
1556                          nilaiTempTotal(0) = rsLaporan!K1
1557                      End If
1558                      If rsLaporan!K10 > 0 Then
1559                          nilaiTempTotal(1) = rsLaporan!K10
1560                      End If
1561                      If rsLaporan!K20 > 0 Then
1562                          nilaiTempTotal(2) = rsLaporan!K20
1563                      End If
1564                      If rsLaporan!K30 > 0 Then
1565                          nilaiTempTotal(3) = rsLaporan!K30
1566                      End If
1567                      If rsLaporan!K40 > 0 Then
1568                          nilaiTempTotal(4) = rsLaporan!K40
1569                      End If
1570                      If rsLaporan!K50 > 0 Then
1571                          nilaiTempTotal(5) = rsLaporan!K50
1572                      End If
1573                      If rsLaporan!K60 > 0 Then
1574                          nilaiTempTotal(6) = rsLaporan!K60
1575                      End If
1576                      If rsLaporan!K70 > 0 Then
1577                          nilaiTempTotal(7) = rsLaporan!K70
1578                      End If
1579                      If rsLaporan!K80 > 0 Then
1580                          nilaiTempTotal(8) = rsLaporan!K80
1581                      End If
1582                      If rsLaporan!k90 > 0 Then
1583                          nilaiTempTotal(9) = rsLaporan!k90
1584                      End If
1585                      If rsLaporan!k100 > 0 Then
1586                          nilaiTempTotal(10) = rsLaporan!k100
1587                      End If
1588                  End If
1589              End If
                  
1590              With rsLaporan
1591                  .AddNew
1592                  If nilaiTempDemand(0) > 0 Then
1593                      !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
1594                  Else
1595                      !K1 = 0 & " %"
1596                  End If
1597                  If nilaiTempDemand(1) > 0 Then
1598                      !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
1599                  Else
1600                      !K10 = 0 & " %"
1601                  End If
1602                  If nilaiTempDemand(2) > 0 Then
1603                      !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
1604                  Else
1605                      !K20 = 0 & " %"
1606                  End If
1607                  If nilaiTempDemand(3) > 0 Then
1608                      !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
1609                  Else
1610                      !K30 = 0 & " %"
1611                  End If
1612                  If nilaiTempDemand(4) > 0 Then
1613                      !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
1614                  Else
1615                      !K40 = 0 & " %"
1616                  End If
1617                  If nilaiTempDemand(5) > 0 Then
1618                      !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
1619                  Else
1620                      !K50 = 0 & " %"
1621                  End If
1622                  If nilaiTempDemand(6) > 0 Then
1623                      !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
1624                  Else
1625                      !K60 = 0 & " %"
1626                  End If
1627                  If nilaiTempDemand(7) > 0 Then
1628                      !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
1629                  Else
1630                      !K70 = 0 & " %"
1631                  End If
1632                  If nilaiTempDemand(8) > 0 Then
1633                      !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
1634                  Else
1635                      !K80 = 0 & " %"
1636                  End If
1637                  If nilaiTempDemand(9) > 0 Then
1638                      !k90 = Round((nilaiTempTotal(9) / nilaiTempDemand(9)) * 100, 0) & " %"
1639                  Else
1640                      !k90 = 0 & " %"
1641                  End If
1642                  If nilaiTempDemand(10) > 0 Then
1643                      !k100 = Round((nilaiTempTotal(10) / nilaiTempDemand(10)) * 100, 0) & " %"
1644                  Else
1645                      !k100 = 0 & " %"
1646                  End If
1647                  !Urut = 9
1648                  !judul = "PERSENTASE RETENTION RATE SBE Non NCM"
1649                  !model = Nul(rsModel.Fields(0).value)
1650                  .Update
1651              End With
                  ''
                  ''----------------------------------------
1652              rsLaporan.filter = ""
1653              rsModel.MoveNext
1654          Loop
1655      End If
prosesCetak:
1656      ProgressBarByModel.value = 5
1657      DoEvents
1658      Report.database.SetDataSource rsLaporan
          
1659      With CRViewer1
1660          .ReportSource = Report
1661          .ViewReport
1662          .Zoom 100
1663          .Visible = True
1664          .Left = 300
1665          .Height = Me.Height - 400
1666          .Width = Me.Width - 750
1667      End With
         
1668      cmdTutup.Visible = True
1669      cmdTutup.Left = CRViewer1.Left + 20
1670      cmdTutup.Top = CRViewer1.Top + 400
1671      Screen.MousePointer = vbDefault
1672      ProgressBarByModel.Visible = False
          
1673      Call DisableExport
1674      Call setLokasiTombol
End Sub

Private Sub cmdCetak_Click()
1         If Month(Tgl1) <> Month(Tgl2) Or year(Tgl2) <> year(Tgl2) Then
2             MsgBox "Range periode dalam 1 bulan", vbInformation
3             Exit Sub
4         End If
5        On Error GoTo cmdCetak_Click_Error
6         Screen.MousePointer = vbHourglass
7         cmdCetak.Enabled = False
8         ProgressBar1.Min = 0
9         ProgressBar1.Max = 10
10        ProgressBar1.Visible = True
11        ProgressBar1.value = 1
12        DoEvents
13        Set Report = New rptRepoSBRate
          
          Dim oRs As ADODB.Recordset
          Dim rsTempAll As ADODB.Recordset
          Dim rsTemp As ADODB.Recordset
          
          Dim konter As Integer
          Dim isTambah As Boolean
          Dim TotalTemp As Double
          
          Dim total1 As Double
          Dim total2 As Double
          Dim total3 As Double
          Dim total4 As Double
          Dim total5 As Double
          
          Dim Demand1 As Double
          Dim Demand2 As Double
          Dim Demand3 As Double
          Dim Demand4 As Double
          Dim Demand5 As Double
          
          Dim SM1 As Double
          Dim SM2 As Double
          Dim SM3 As Double
          Dim SM4 As Double
          Dim SM5 As Double
          
          Dim NCM1 As Double
          Dim NCM2 As Double
          Dim NCM3 As Double
          Dim NCM4 As Double
          Dim NCM5 As Double
          
          Dim NonNCM1 As Double
          Dim NonNCM2 As Double
          Dim NonNCM3 As Double
          Dim NonNCM4 As Double
          Dim NonNCM5 As Double
          
          Dim Cetak As Double
          'Call createRS
          
          Dim i As Integer
              
14            Demand1 = 0
15            Demand2 = 0
16            Demand3 = 0
17            Demand4 = 0
18            Demand5 = 0
              
19            SM1 = 0
20            SM2 = 0
21            SM3 = 0
22            SM4 = 0
23            SM5 = 0
              
24            NCM1 = 0
25            NCM2 = 0
26            NCM3 = 0
27            NCM4 = 0
28            NCM5 = 0
              
29            NonNCM1 = 0
30            NonNCM2 = 0
31            NonNCM3 = 0
32            NonNCM4 = 0
33            NonNCM5 = 0
              
             ' lblPengambilanDataRet.Visible = True
              ''proses pengambilan data same outlate
              
34            If Log.RPLHistori = 0 Then
35                sSQL = "select COUNT(noinv) total,fs_rc"
36                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
37                sSQL = sSQL & " INNER JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
38                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
39                sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
40                sSQL = sSQL & " and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050') "
41                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
42                sSQL = sSQL & " grOUP BY FS_RC"
43                sSQL = sSQL & " order by FS_RC"
44            Else
45                sSQL = "select COUNT(noinv) total,fs_rc"
46                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
47                sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
48                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
49                sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
50                sSQL = sSQL & " and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050') "
51                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
52                sSQL = sSQL & " grOUP BY FS_RC"
53                sSQL = sSQL & " order by FS_RC"
54            End If
55            Set rsTemp = objO.FetchData(sSQL)
56            Cetak = 0
57            ProgressBar1.value = 2
58            DoEvents
59            If Not rsTemp.EOF Then
                  ''SBE 10
60                rsTemp.filter = "fs_rc = 'R010'"
61                Cetak = 0
62                If Not rsTemp.EOF Then
63                    Cetak = CDbl(rsTemp("TOTAL"))
64                End If
65                rsTemp.filter = ""
66                Report.txtSO1.SetText Cetak
67                total1 = total1 + Cetak
68                SM1 = SM1 + Cetak
                  
                  ''SBE 20
69                rsTemp.filter = "fs_rc = 'R020'"
70                Cetak = 0
71                If Not rsTemp.EOF Then
72                    Cetak = CDbl(rsTemp("TOTAL"))
73                End If
74                rsTemp.filter = ""
75                Report.txtSO2.SetText Cetak
76                total2 = total2 + Cetak
77                SM2 = SM2 + Cetak
                  
                  ''SBE 30
78                rsTemp.filter = "fs_rc = 'R030'"
79                Cetak = 0
80                If Not rsTemp.EOF Then
81                    Cetak = CDbl(rsTemp("TOTAL"))
82                End If
83                rsTemp.filter = ""
84                Report.txtSO3.SetText Cetak
85                total3 = total3 + Cetak
86                SM3 = SM3 + Cetak
                  
                  ''SBE 40
87                rsTemp.filter = "fs_rc = 'R040'"
88                Cetak = 0
89                If Not rsTemp.EOF Then
90                    Cetak = CDbl(rsTemp("TOTAL"))
91                End If
92                rsTemp.filter = ""
93                Report.txtSO4.SetText Cetak
94                total4 = total4 + Cetak
95                SM4 = SM4 + Cetak
                  
                  ''SBE 50
96                rsTemp.filter = "fs_rc = 'R050'"
97                Cetak = 0
98                If Not rsTemp.EOF Then
99                    Cetak = CDbl(rsTemp("TOTAL"))
100               End If
101               rsTemp.filter = ""
102               Report.txtSO5.SetText Cetak
103               total5 = total5 + Cetak
104               SM5 = SM5 + Cetak
105           Else
106               Report.txtSO1.SetText Cetak
107               Report.txtSO2.SetText Cetak
108               Report.txtSO3.SetText Cetak
109               Report.txtSO4.SetText Cetak
110               Report.txtSO5.SetText Cetak
111           End If
              
112           If Log.RPLHistori = 0 Then
113               sSQL = "select fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
114               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
115               sSQL = sSQL & " left join SALESUNIT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA"
116               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
117               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC < 'R055' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
118               sSQL = sSQL & " group by fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end"
119               sSQL = sSQL & " order by fs_Rc"
120           Else
121               sSQL = "select fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
122               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
123               sSQL = sSQL & " left join [" & dbReplikasiHO & "]..SUMMARY B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA  and B.tipetr = 'ND' And B.[Status] = '1' "
124               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
125               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC < 'R055' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
126               sSQL = sSQL & " group by fs_rc, case when b.norangka is null then 'Non NCM' else 'NCM' end"
127               sSQL = sSQL & " order by fs_Rc"
128           End If
129           Set rsTemp = objO.FetchData(sSQL)
130           ProgressBar1.value = 3
131           DoEvents
132           Cetak = 0
133           If Not rsTemp.EOF Then
                  ''SBE 10
134               rsTemp.filter = "fs_rc = 'R010' AND STATUS = 'NCM'"
135               Cetak = 0
136               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
137                   Cetak = CDbl(rsTemp("unit")) - SM1
138               End If
139               Report.txtNCM1.SetText Cetak
140               total1 = total1 + Cetak
141               NCM1 = NCM1 + Cetak
142               rsTemp.filter = ""
                  
143               rsTemp.filter = "fs_rc = 'R010' AND STATUS = 'Non NCM'"
144               Cetak = 0
145               If Not rsTemp.EOF Then
146                   Cetak = CDbl(rsTemp("unit"))
147               End If
148               Report.txtNonNCM1.SetText Cetak
149               total1 = total1 + Cetak
150               NonNCM1 = NonNCM1 + Cetak
151               rsTemp.filter = ""
                  
                  ''SBE 20
152               rsTemp.filter = "fs_rc = 'R020' AND STATUS = 'NCM'"
153               Cetak = 0
154               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
155                   Cetak = CDbl(rsTemp("unit")) - SM2
156               End If
157               Report.txtNCM2.SetText Cetak
158               total2 = total2 + Cetak
159               NCM2 = NCM2 + Cetak
160               rsTemp.filter = ""
                  
161               rsTemp.filter = "fs_rc = 'R020' AND STATUS = 'Non NCM'"
162               Cetak = 0
163               If Not rsTemp.EOF Then
164                   Cetak = CDbl(rsTemp("unit"))
165               End If
166               Report.txtNonNCM2.SetText Cetak
167               total2 = total2 + Cetak
168               NonNCM2 = NonNCM2 + Cetak
169               rsTemp.filter = ""
                  
                  ''SBE 30
170               rsTemp.filter = "fs_rc = 'R030' AND STATUS = 'NCM'"
171               Cetak = 0
172               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
173                   Cetak = CDbl(rsTemp("unit")) - SM3
174               End If
175               Report.txtNCM3.SetText Cetak
176               total3 = total3 + Cetak
177               NCM3 = NCM3 + Cetak
178               rsTemp.filter = "" '
                  
179               rsTemp.filter = "fs_rc = 'R030' AND STATUS = 'Non NCM'"
180               Cetak = 0
181               If Not rsTemp.EOF Then
182                   Cetak = CDbl(rsTemp("unit"))
183               End If
184               Report.txtNonNCM3.SetText Cetak
185               total3 = total3 + Cetak
186               NonNCM3 = NonNCM3 + Cetak
187               rsTemp.filter = ""
                  
                  ''SBE 40
188               rsTemp.filter = "fs_rc = 'R040' AND STATUS = 'NCM'"
189               Cetak = 0
190               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
191                   Cetak = CDbl(rsTemp("unit")) - SM4
192               End If
193               Report.txtNCM4.SetText Cetak
194               total4 = total4 + Cetak
195               NCM4 = NCM4 + Cetak
196               rsTemp.filter = ""
                  
197               rsTemp.filter = "fs_rc = 'R040' AND STATUS = 'Non NCM'"
198               Cetak = 0
199               If Not rsTemp.EOF Then
200                   Cetak = CDbl(rsTemp("unit"))
201               End If
202               Report.txtNonNCM4.SetText Cetak
203               total4 = total4 + Cetak
204               NonNCM4 = NonNCM4 + Cetak
205               rsTemp.filter = ""
                  
                  ''SBE 50
206               rsTemp.filter = "fs_rc = 'R050' AND STATUS = 'NCM'"
207               Cetak = 0
208               If Not rsTemp.EOF Then
                      ''2016-02-10 || exclude unit same outlet
209                   Cetak = CDbl(rsTemp("unit")) - SM5
210               End If
211               Report.txtNCM5.SetText Cetak
212               total5 = total5 + Cetak
213               NCM5 = NCM5 + Cetak
214               rsTemp.filter = ""
                  
215               rsTemp.filter = "fs_rc = 'R050' AND STATUS = 'Non NCM'"
216               Cetak = 0
217               If Not rsTemp.EOF Then
218                   Cetak = CDbl(rsTemp("unit"))
219               End If
220               Report.txtNonNCM5.SetText Cetak
221               total5 = total5 + Cetak
222               NonNCM5 = NonNCM5 + Cetak
223               rsTemp.filter = ""
224           Else
225               Report.txtNCM1.SetText Cetak
226               Report.txtNonNCM1.SetText Cetak
227               Report.txtNCM2.SetText Cetak
228               Report.txtNonNCM2.SetText Cetak
229               Report.txtNCM3.SetText Cetak
230               Report.txtNonNCM3.SetText Cetak
231               Report.txtNCM4.SetText Cetak
232               Report.txtNonNCM4.SetText Cetak
233               Report.txtNCM5.SetText Cetak
234               Report.txtNonNCM5.SetText Cetak
235           End If
236           ProgressBar1.value = 4
237           DoEvents
238           Report.txtS41.SetText total1
239           Report.txtS42.SetText total2
240           Report.txtS43.SetText total3
241           Report.txtS44.SetText total4
242           Report.txtS45.SetText total5
              
243           Demand1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
244           Demand2 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
245           Demand3 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
246           Demand4 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
247           Demand5 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Format(Tgl1, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Format(Tgl2, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
              
248           Report.txtS61.SetText Demand1
249           Report.txtS62.SetText Demand2
250           Report.txtS63.SetText Demand3
251           Report.txtS64.SetText Demand4
252           Report.txtS65.SetText Demand5
253           ProgressBar1.value = 5
254           DoEvents
              ''ratention rate sameoutlate
255           If Demand1 = 0 Then
256             Report.txtS81.SetText 0 & " %"
257           Else
258             Report.txtS81.SetText Round(((SM1 / Demand1) * 100), 0) & " %"
259           End If
260           If Demand2 = 0 Then
261             Report.txtS82.SetText 0 & " %"
262           Else
263             Report.txtS82.SetText Round(((SM2 / Demand2) * 100), 0) & " %"
264           End If
265           If Demand3 = 0 Then
266             Report.txtS83.SetText 0 & " %"
267           Else
268             Report.txtS83.SetText Round(((SM3 / Demand3) * 100), 0) & " %"
269           End If
270           If Demand4 = 0 Then
271             Report.txtS84.SetText 0 & " %"
272           Else
273             Report.txtS84.SetText Round(((SM4 / Demand4) * 100), 0) & " %"
274           End If
275           If Demand5 = 0 Then
276             Report.txtS85.SetText 0 & " %"
277           Else
278             Report.txtS85.SetText Round(((SM5 / Demand5) * 100), 0) & " %"
279           End If
              '' MENGGUNAKAN IIF TIDAK BISA DIGUNAKAN
      '
      '        Report.txtS81.SetText IIf(Demand1 = 0, 0, Round(((SM1 / Demand1) * 100), 0)) & " %"
      '        Report.txtS82.SetText IIf(Demand2 = 0, 0, Round(((SM2 / Demand2) * 100), 0)) & " %"
      '        Report.txtS83.SetText IIf(Demand3 = 0, 0, Round(((SM3 / Demand3) * 100), 0)) & " %"
      '        Report.txtS84.SetText IIf(Demand4 = 0, 0, Round(((SM4 / Demand4) * 100), 0)) & " %"
      '        Report.txtS85.SetText IIf(Demand5 = 0, 0, Round(((SM5 / Demand5) * 100), 0)) & " %"
              ''--------------------------
              
              ''total ratention rate
280           If Demand1 = 0 Then
281               Report.txtS91.SetText 0 & " %"
282           Else
283               Report.txtS91.SetText Round(((total1 / Demand1) * 100), 0) & " %"
284           End If
285           If Demand2 = 0 Then
286               Report.txtS92.SetText 0 & " %"
287           Else
288               Report.txtS92.SetText Round(((total2 / Demand2) * 100), 0) & " %"
289           End If
290           If Demand3 = 0 Then
291               Report.txtS93.SetText 0 & " %"
292           Else
293               Report.txtS93.SetText Round(((total3 / Demand3) * 100), 0) & " %"
294           End If
295           If Demand4 = 0 Then
296               Report.txtS94.SetText 0 & " %"
297           Else
298               Report.txtS94.SetText Round(((total4 / Demand4) * 100), 0) & " %"
299           End If
300           If Demand5 = 0 Then
301               Report.txtS95.SetText 0 & " %"
302           Else
303               Report.txtS95.SetText Round(((total5 / Demand5) * 100), 0) & " %"
304           End If
              ''
      '        Report.txtS91.SetText IIf(Demand1 = 0, 0, Round(((total1 / Demand1) * 100), 0)) & " %"
      '        Report.txtS92.SetText IIf(Demand2 = 0, 0, Round(((total2 / Demand2) * 100), 0)) & " %"
      '        Report.txtS93.SetText IIf(Demand3 = 0, 0, Round(((total3 / Demand3) * 100), 0)) & " %"
      '        Report.txtS94.SetText IIf(Demand4 = 0, 0, Round(((total4 / Demand4) * 100), 0)) & " %"
      '        Report.txtS95.SetText IIf(Demand5 = 0, 0, Round(((total5 / Demand5) * 100), 0)) & " %"
              ''--------------------
              
              ''Non NCM / Total
305           If total1 = 0 Then
306               Report.txtS121.SetText 0 & " %"
307           Else
308               Report.txtS121.SetText Round(((NonNCM1 / total1) * 100), 0) & " %"
309           End If
310           If total2 = 0 Then
311               Report.txtS122.SetText 0 & " %"
312           Else
313               Report.txtS122.SetText Round(((NonNCM2 / total2) * 100), 0) & " %"
314           End If
315           If total3 = 0 Then
316               Report.txtS123.SetText 0 & " %"
317           Else
318               Report.txtS123.SetText Round(((NonNCM3 / total3) * 100), 0) & " %"
319           End If
320           If total4 = 0 Then
321               Report.txtS124.SetText 0 & " %"
322           Else
323               Report.txtS124.SetText Round(((NonNCM4 / total4) * 100), 0) & " %"
324           End If
325           If total5 = 0 Then
326               Report.txtS125.SetText 0 & " %"
327           Else
328               Report.txtS125.SetText Round(((NonNCM5 / total5) * 100), 0) & " %"
329           End If
      '        Report.txtS121.SetText IIf(total1 = 0, 0, Round(((NonNCM1 / total1) * 100), 0)) & " %"
      '        Report.txtS122.SetText IIf(total2 = 0, 0, Round(((NonNCM2 / total2) * 100), 0)) & " %"
      '        Report.txtS123.SetText IIf(total3 = 0, 0, Round(((NonNCM3 / total3) * 100), 0)) & " %"
      '        Report.txtS124.SetText IIf(total4 = 0, 0, Round(((NonNCM4 / total4) * 100), 0)) & " %"
      '        Report.txtS125.SetText IIf(total5 = 0, 0, Round(((NonNCM5 / total5) * 100), 0)) & " %"
              ''---------------
              
330           Report.txtSMTotal.SetText (SM1 + SM2 + SM3 + SM4 + SM5)
331           Report.Text10.SetText (NCM1 + NCM2 + NCM3 + NCM4 + NCM5)
332           Report.Text13.SetText (NonNCM1 + NonNCM2 + NonNCM3 + NonNCM4 + NonNCM5)
333           Report.Text14.SetText (total1 + total2 + total3 + total4 + total5)
334           Report.Text15.SetText Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5) / 5, 0)
              ''Report.Text15.SetText Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5), 0)
              
335           If Round((Demand1 + Demand2 + Demand3 + Demand4 + Demand5), 0) > 0 Then
336               Report.Text17.SetText Round(((SM1 + SM2 + SM3 + SM4 + SM5) / (Demand1 + Demand2 + Demand3 + Demand4 + Demand5)) * 100, 0) & " %"
337               Report.Text21.SetText Round(((total1 + total2 + total3 + total4 + total5) / (Demand1 + Demand2 + Demand3 + Demand4 + Demand5)) * 100, 0) & " %"
338           Else
339               Report.Text17.SetText "0 %"
340               Report.Text21.SetText "0 %"
341           End If
              
342           ProgressBar1.value = 6
343           DoEvents
344           TotalTemp = 0
345           If Demand1 > 0 Then
346               TotalTemp = TotalTemp + Round(((SM1 / Demand1) * 100), 0)
347           End If
348           If Demand2 > 0 Then
349               TotalTemp = TotalTemp + Round(((SM2 / Demand2) * 100), 0)
350           End If
351           If Demand3 > 0 Then
352               TotalTemp = TotalTemp + Round(((SM3 / Demand3) * 100), 0)
353           End If
354           If Demand4 > 0 Then
355               TotalTemp = TotalTemp + Round(((SM4 / Demand4) * 100), 0)
356           End If
357           If Demand5 > 0 Then
358               TotalTemp = TotalTemp + Round(((SM5 / Demand5) * 100), 0)
359           End If
              ''Report.Text17.SetText Round(TotalTemp / 5, 0) & " %"
              'Report.Text17.SetText Round((((IIf(Demand1 = 0, 0, Round(((SM1 / Demand1) * 100), 0))) + (IIf(Demand2 = 0, 0, Round(((SM2 / Demand2) * 100), 0))) + (IIf(Demand3 = 0, 0, Round(((SM3 / Demand3) * 100), 0))) + (IIf(Demand4 = 0, 0, Round(((SM4 / Demand4) * 100), 0))) + (IIf(Demand5 = 0, 0, Round(((SM5 / Demand5) * 100), 0)))) / 5), 0) & " %"
              
360           TotalTemp = 0
361           If Demand1 > 0 Then
362               TotalTemp = TotalTemp + Round(((total1 / Demand1) * 100), 0)
363           End If
364           If Demand2 > 0 Then
365               TotalTemp = TotalTemp + Round(((total2 / Demand2) * 100), 0)
366           End If
367           If Demand3 > 0 Then
368               TotalTemp = TotalTemp + Round(((total3 / Demand3) * 100), 0)
369           End If
370           If Demand4 > 0 Then
371               TotalTemp = TotalTemp + Round(((total4 / Demand4) * 100), 0)
372           End If
373           If Demand5 > 0 Then
374               TotalTemp = TotalTemp + Round(((total5 / Demand5) * 100), 0)
375           End If
              ''Report.Text21.SetText Round(TotalTemp / 5, 0) & " %"
              'Report.Text21.SetText Round((((IIf(Demand1 = 0, 0, Round(((total1 / Demand1) * 100), 0))) + (IIf(Demand2 = 0, 0, Round(((total2 / Demand2) * 100), 0))) + (IIf(Demand3 = 0, 0, Round(((total3 / Demand3) * 100), 0))) + (IIf(Demand4 = 0, 0, Round(((total4 / Demand4) * 100), 0))) + (IIf(Demand5 = 0, 0, Round(((total5 / Demand5) * 100), 0)))) / 5), 0) & " %"
              
376           TotalTemp = 0
377           If total1 > 0 Then
378               TotalTemp = TotalTemp + Round(((NonNCM1 / total1) * 100), 0)
379           End If
380           If total2 > 0 Then
381               TotalTemp = TotalTemp + Round(((NonNCM2 / total2) * 100), 0)
382           End If
383           If total3 > 0 Then
384               TotalTemp = TotalTemp + Round(((NonNCM3 / total3) * 100), 0)
385           End If
386           If total4 > 0 Then
387               TotalTemp = TotalTemp + Round(((NonNCM4 / total4) * 100), 0)
388           End If
389           If total5 > 0 Then
390               TotalTemp = TotalTemp + Round(((NonNCM5 / total5) * 100), 0)
391           End If
392           Report.Text23.SetText Round(TotalTemp / 5, 0) & " %"
              
              'Report.Text23.SetText Round((((IIf(total1 = 0, 0, Round(((NonNCM1 / total1) * 100), 0))) + (IIf(total2 = 0, 0, Round(((NonNCM2 / total2) * 100), 0))) + (IIf(total3 = 0, 0, Round(((NonNCM3 / total3) * 100), 0))) + (IIf(total4 = 0, 0, Round(((NonNCM4 / total4) * 100), 0))) + (IIf(total5 = 0, 0, Round(((NonNCM5 / total5) * 100), 0)))) / 5), 0) & " %"
393           ProgressBar1.value = 7
394           DoEvents
395           sSQL = "select COUNT(noinv) total "
396           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
397           sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
398           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
399           sSQL = sSQL & " AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0 AND INVDATE between '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "'"
              
400           total1 = 0
401           Set rsTemp = objO.FetchData(sSQL)
402           If Not rsTemp.EOF Then
403               rsTemp.MoveFirst
404               total1 = CDbl(rsTemp!total)
405           End If
406           Report.Text43.SetText total1
              
407           ProgressBar1.value = 8
408           DoEvents
              
409           If Log.RPLHistori = 0 Then
410               sSQL = "select case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
411               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
412               sSQL = sSQL & " left join SALESUNIT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA"
413               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
414               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
415               sSQL = sSQL & " group by case when b.norangka is null then 'Non NCM' else 'NCM' end"
416           Else
417               sSQL = "select case when b.norangka is null then 'Non NCM' else 'NCM' end as status, count(no_chassis) as unit"
418               sSQL = sSQL & " from t_invoice A WITH (NOLOCK)"
419               sSQL = sSQL & " left join [" & dbReplikasiHO & "]..SUMMARY B WITH (NOLOCK) ON A.NO_CHASSIS = B.NORANGKA and B.tipetr = 'ND' And B.[Status] = '1' "
420               sSQL = sSQL & " where invdate between '" & Format(Tgl1, "MM/DD/YYYY") & "' and '" & Format(Tgl2, "MM/DD/YYYY") & "'  and noinv not like '%-b'"
421               sSQL = sSQL & " and a.fs_kd_dealer = '" & Log.KodeDealer & "' and fs_log_del = '' AND FS_RC <> '' AND FS_NO_RETUR = ''  AND FS_RC NOT LIKE '%5' AND TPSVC <> 'W' AND FS_RC <> 'R001' AND PDS = 0"
422               sSQL = sSQL & " group by case when b.norangka is null then 'Non NCM' else 'NCM' end"
423           End If
424           Set rsTemp = objO.FetchData(sSQL)
425           NCM1 = 0
426           NonNCM1 = 0
427           If Not rsTemp.EOF Then
428               rsTemp.MoveFirst
429               rsTemp.filter = "STATUS = 'NCM'"
430               If Not rsTemp.EOF Then
431                   NCM1 = CDbl(rsTemp!unit)
432               End If
433               rsTemp.filter = ""
434               rsTemp.filter = "STATUS = 'Non NCM'"
435               If Not rsTemp.EOF Then
436                   NonNCM1 = CDbl(rsTemp!unit)
437               End If
438           End If
              
439           total2 = 0
440           total2 = (total1 + NCM1 + NonNCM1)
              
441           Report.Text44.SetText NCM1
442           Report.Text45.SetText NonNCM1
443           Report.Text46.SetText total2 '''(total1 + NCM1 + NonNCM1)
              
444           If total2 = 0 Then
445               Report.Text50.SetText "0 %"
446           Else
447               Report.Text50.SetText Round((NonNCM1 / total2) * 100, 0) & " %"
448           End If
              ''demand service share : rata2 per bulan penjualan 8 thn kebelakang * 2.2
              
449           ProgressBar1.value = 9
450           DoEvents
              
451           Demand1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(YEAR,-8,'" & Format(Tgl1, "MM/DD/YYYY") & "') and '" & Format(Tgl2, "MM/DD/YYYY") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
452           Demand1 = Round(Demand1 * 2.2, 0)
              
453           Report.Text47.SetText Demand1
              
454           If Demand1 > 0 Then
455               Report.Text48.SetText Round((total1 / Demand1) * 100, 0) & " %"
456               Report.Text49.SetText Round((total2 / Demand1) * 100, 0) & " %"
457           Else
458               Report.Text48.SetText "0 %"
459               Report.Text49.SetText "0 %"
460           End If
              
461           ProgressBar1.value = 10
462           DoEvents
463           ProgressBar1.Visible = False
464           cmdCetak.Enabled = True
              
              ''
465           Report.Text11.SetText "Laporan Retention Rate " & gPerush
466           Report.Text9.SetText GetNamaBulan(Month(Tgl1), 1) & " " & year(Tgl1)
467           Report.Text27.SetText "Laporan Service Share " & gPerush
468           Report.Text26.SetText GetNamaBulan(Month(Tgl1), 1) & " " & year(Tgl1)
              
469       With CRViewer1
470           .ReportSource = Report
471           .ViewReport
472           .Zoom 100
473           .Visible = True
474           .Left = 300
475           .Height = Me.Height - 400
476           .Width = Me.Width - .Left
477       End With
         
478               cmdTutup.Visible = True
479               cmdTutup.Left = CRViewer1.Left + 20
480               cmdTutup.Top = CRViewer1.Top + 400
                  
          'ProgressBar2.Visible = False
          'Frame2.Visible = False
481       Screen.MousePointer = vbDefault
482       Call DisableExport
483       Call setLokasiTombol
484      Exit Sub

cmdCetak_Click_Error:
485       Screen.MousePointer = vbDefault
486       cmdCetak.Enabled = True
487       ProgressBar1.Visible = True
488       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmRepoSBRateRetention")
End Sub
 


Private Sub cmdCetakGrafik_Click()
   On Error GoTo cmdCetakK_Click_Error

    If Month(Tgl1K) <> Month(Tgl2K) Or year(Tgl2K) <> year(Tgl2K) Then
        MsgBox "Range periode dalam 1 bulan", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    ProgressBar1K.Min = 0
    ProgressBar1K.Max = 10
    ProgressBar1K.Visible = True
    ProgressBar1K.value = 1
    DoEvents
    
    createTableLaporan
    ProgressBar1K.value = 2
    DoEvents
    
    If opTipe(0).value = True Then
        GoTo prosesTanggalService
    ElseIf opTipe(1).value = True Then
        GoTo prosesTanggalRetail
    Else
        GoTo prosesTanggalPDS
    End If
    
    
prosesTanggalService:
    ''mencari unit
    Call clearNilai
    If Log.RPLHistori = 0 Then
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan "
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
        sSQL = sSQL & " LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B'"
        sSQL = sSQL & " AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1K, "MM/DD/YYYY") & "' AND '" & Format(Tgl2K, "MM/DD/YYYY") & "'"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end"
        sSQL = sSQL & " order by FS_RC"
    Else
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan "
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA  and s.tipetr = 'ND' And s.[Status] = '1' "
        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B'"
        sSQL = sSQL & " AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1K, "MM/DD/YYYY") & "' AND '" & Format(Tgl2K, "MM/DD/YYYY") & "'"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end"
        sSQL = sSQL & " order by FS_RC"
    End If
    Set rsTemp = objO.FetchData(sSQL)
    Cetak = 0
    ProgressBar1K.value = 2
    DoEvents
    If Not rsTemp.EOF Then
        ''--- SBE 1K ------------------
        rsTemp.filter = "fs_rc = 'R001' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(0) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(0) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R001' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(0) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(0) = ncm(0) + nonNcm(0) + sameOutlet(0)
        ''--- AKHIR SBE 1K ------------
        
        ''--- SBE 10K ------------------
        rsTemp.filter = "fs_rc = 'R010' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(1) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(1) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R010' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(1) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(1) = ncm(1) + nonNcm(1) + sameOutlet(1)
        ''--- AKHIR SBE 10K ------------
        
        ''--- SBE 20K ------------------
        rsTemp.filter = "fs_rc = 'R020' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(2) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(2) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R020' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(2) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(2) = ncm(2) + nonNcm(2) + sameOutlet(2)
        ''--- AKHIR SBE 20K ------------
        
        ''--- SBE 30K ------------------
        rsTemp.filter = "fs_rc = 'R030' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(3) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(3) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R030' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(3) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(3) = ncm(3) + nonNcm(3) + sameOutlet(3)
        ''--- AKHIR SBE 30K ------------
        
        ''--- SBE 40K ------------------
        rsTemp.filter = "fs_rc = 'R040' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(4) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(4) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R040' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(4) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(4) = ncm(4) + nonNcm(4) + sameOutlet(4)
        ''--- AKHIR SBE 40K ------------
        
        ''--- SBE 50K ------------------
        rsTemp.filter = "fs_rc = 'R050' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(5) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(5) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R050' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(5) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(5) = ncm(5) + nonNcm(5) + sameOutlet(5)
        ''--- AKHIR SBE 50K ------------
        
        ''--- SBE 60K ------------------
        rsTemp.filter = "fs_rc = 'R060' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(6) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(6) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R060' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(6) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(6) = ncm(6) + nonNcm(6) + sameOutlet(6)
        ''--- AKHIR SBE 60K ------------
        
        ''--- SBE 70K ------------------
        rsTemp.filter = "fs_rc = 'R070' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(7) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(7) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R070' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(7) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(7) = ncm(7) + nonNcm(7) + sameOutlet(7)
        ''--- AKHIR SBE 70K ------------
        
        ''--- SBE 80K ------------------
        rsTemp.filter = "fs_rc = 'R080' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(8) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(8) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R080' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(8) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(8) = ncm(8) + nonNcm(8) + sameOutlet(8)
        ''--- AKHIR SBE 80K ------------
        
        ''--- SBE 90K ------------------
        rsTemp.filter = "fs_rc = 'R090' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(9) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R090' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(9) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R090' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(9) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(9) = ncm(9) + nonNcm(9) + sameOutlet(9)
        ''--- AKHIR SBE 90K ------------
        
        ''--- SBE 100K ------------------
        rsTemp.filter = "fs_rc = 'R100' and keterangan = 'NCM'"
        If Not rsTemp.EOF Then
            ncm(10) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R100' and keterangan = 'Same Outlet'"
        If Not rsTemp.EOF Then
            sameOutlet(10) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        rsTemp.filter = "fs_rc = 'R100' and keterangan = 'Non NCM'"
        If Not rsTemp.EOF Then
            nonNcm(10) = CDbl(rsTemp!total)
        End If
        rsTemp.filter = ""
        TotalSBE(10) = ncm(10) + nonNcm(10) + sameOutlet(10)
        ''--- AKHIR SBE 100K ------------
   
    End If
    
    ProgressBar1K.value = 3
    DoEvents
    
    ''--- MENCARI NILAI DEMAND UNTUK SETIAP SBE
    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    
    
    ''--- AKHIR NILAI DEMAND -------------------
    GoTo prosesCetak
    
prosesTanggalRetail:
    ''2017-08-11
    ''mencari data service dari tanggal retail
    ''yang diambil data service kelipatan sbe masuk range tanggal retail
    
    Call clearNilai
    If Log.RPLHistori = 0 Then
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
        sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(Tgl1K, "mm/dd/yyyy") & "' AND '" & Format(Tgl2K, "mm/dd/yyyy") & "' )"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
        sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
        'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
        sSQL = sSQL & " order by FS_RC"
    Else
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
        sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(Tgl1K, "mm/dd/yyyy") & "' AND '" & Format(Tgl2K, "mm/dd/yyyy") & "' )"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
        sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
        'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
        sSQL = sSQL & " order by FS_RC"
    End If
    Set rsTemp = objO.FetchData(sSQL)
    ProgressBar1K.value = 2
    DoEvents
    
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Nul(rsTemp!FS_RC) = "R001" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(0) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(0) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(0) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R010" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(1) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(1) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(1) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R020" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(2) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(2) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(2) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R030" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(3) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(3) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(3) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R040" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(4) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(4) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(4) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R050" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(5) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(5) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(5) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R060" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(6) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(6) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(6) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R070" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(7) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(7) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(7) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R080" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(8) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(8) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(8) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R090" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(9) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(9) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(9) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R100" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(10) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(10) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(10) = CDbl(rsTemp!total)
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    ProgressBar1K.value = 3
    DoEvents
    
    ''--- MENCARI NILAI DEMAND BERDASARKAN TANGGAL RETAIL
'    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 0 and 1) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    
    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and (dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")

    
'    For i = 1 To 10
'        demand(i) = demand(0)
'    Next
    
    ''--- AKHIR NILAI DEMAND -------------------
    GoTo prosesCetak
    
prosesTanggalPDS:
    ''2017-08-11
    ''mencari data service dari tanggal retail
    ''yang diambil data service kelipatan sbe masuk range tanggal retail
    
    Call clearNilai
    If Log.RPLHistori = 0 Then
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
        sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "')"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
        sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
        'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
        sSQL = sSQL & " order by FS_RC"
    Else
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
        sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
        sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "')"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
        sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
        'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
        sSQL = sSQL & " order by FS_RC"
    End If
    Set rsTemp = objO.FetchData(sSQL)
    ProgressBar1K.value = 2
    DoEvents
    
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Nul(rsTemp!FS_RC) = "R001" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(0) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(0) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(0) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R010" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(1) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(1) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(1) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R020" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(2) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(2) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(2) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R030" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(3) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(3) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(3) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R040" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(4) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(4) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(4) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R050" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(5) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(5) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(5) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R060" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(6) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(6) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(6) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R070" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(7) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(7) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(7) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R080" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(8) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(8) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(8) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R090" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(9) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(9) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(9) = CDbl(rsTemp!total)
                End If
            ElseIf Nul(rsTemp!FS_RC) = "R100" Then
                If Nul(rsTemp!keterangan) = "Same Outlet" Then
                    sameOutlet(10) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "NCM" Then
                    ncm(10) = CDbl(rsTemp!total)
                ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
                    nonNcm(10) = CDbl(rsTemp!total)
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    ProgressBar1K.value = 3
    DoEvents
    
    ''--- MENCARI NILAI DEMAND BERDASARKAN TANGGAL RETAIL
    'demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 0 and 1) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
'    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    
    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and (dateadd(day,-1,(DATEADD(month,-1,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")

    
    ''--- AKHIR NILAI DEMAND -------------------
    GoTo prosesCetak
    

prosesCetak:
        
    ProgressBar1K.value = 4
    DoEvents
    
    ''--- MENCARI NILAI RETENTION TIAP SBE -----
    
    If demand(0) > 0 Then
        cetakRateSameOutlet(0) = Round((sameOutlet(0) / demand(0)) * 100, 0)
        cetakRateNCM(0) = Round((ncm(0) / demand(0)) * 100, 0)
        cetakRateNonNCM(0) = Round((nonNcm(0) / demand(0)) * 100, 0)
        cetakRateTotal(0) = Round((TotalSBE(0) / demand(0)) * 100, 0)
    End If
    
    If demand(1) > 0 Then
        cetakRateSameOutlet(1) = Round((sameOutlet(1) / demand(1)) * 100, 0)
        cetakRateNCM(1) = Round((ncm(1) / demand(1)) * 100, 0)
        cetakRateNonNCM(1) = Round((nonNcm(1) / demand(1)) * 100, 0)
        cetakRateTotal(1) = Round((TotalSBE(1) / demand(1)) * 100, 0)
    End If
    
    If demand(2) > 0 Then
        cetakRateSameOutlet(2) = Round((sameOutlet(2) / demand(2)) * 100, 0)
        cetakRateNCM(2) = Round((ncm(2) / demand(2)) * 100, 0)
        cetakRateNonNCM(2) = Round((nonNcm(2) / demand(2)) * 100, 0)
        cetakRateTotal(2) = Round((TotalSBE(2) / demand(2)) * 100, 0)
    End If
    
    If demand(3) > 0 Then
        cetakRateSameOutlet(3) = Round((sameOutlet(3) / demand(3)) * 100, 0)
        cetakRateNCM(3) = Round((ncm(3) / demand(3)) * 100, 0)
        cetakRateNonNCM(3) = Round((nonNcm(3) / demand(3)) * 100, 0)
        cetakRateTotal(3) = Round((TotalSBE(3) / demand(3)) * 100, 0)
    End If
    
    If demand(4) > 0 Then
        cetakRateSameOutlet(4) = Round((sameOutlet(4) / demand(4)) * 100, 0)
        cetakRateNCM(4) = Round((ncm(4) / demand(4)) * 100, 0)
        cetakRateNonNCM(4) = Round((nonNcm(4) / demand(4)) * 100, 0)
        cetakRateTotal(4) = Round((TotalSBE(4) / demand(4)) * 100, 0)
    End If
    
    If demand(5) > 0 Then
        cetakRateSameOutlet(5) = Round((sameOutlet(5) / demand(5)) * 100, 0)
        cetakRateNCM(5) = Round((ncm(5) / demand(5)) * 100, 0)
        cetakRateNonNCM(5) = Round((nonNcm(5) / demand(5)) * 100, 0)
        cetakRateTotal(5) = Round((TotalSBE(5) / demand(5)) * 100, 0)
    End If
    
    If demand(6) > 0 Then
        cetakRateSameOutlet(6) = Round((sameOutlet(6) / demand(6)) * 100, 0)
        cetakRateNCM(6) = Round((ncm(6) / demand(6)) * 100, 0)
        cetakRateNonNCM(6) = Round((nonNcm(6) / demand(6)) * 100, 0)
        cetakRateTotal(6) = Round((TotalSBE(6) / demand(6)) * 100, 0)
    End If
    
    If demand(7) > 0 Then
        cetakRateSameOutlet(7) = Round((sameOutlet(7) / demand(7)) * 100, 0)
        cetakRateNCM(7) = Round((ncm(7) / demand(7)) * 100, 0)
        cetakRateNonNCM(7) = Round((nonNcm(7) / demand(7)) * 100, 0)
        cetakRateTotal(7) = Round((TotalSBE(7) / demand(7)) * 100, 0)
    End If
    
    If demand(8) > 0 Then
        cetakRateSameOutlet(8) = Round((sameOutlet(8) / demand(8)) * 100, 0)
        cetakRateNCM(8) = Round((ncm(8) / demand(8)) * 100, 0)
        cetakRateNonNCM(8) = Round((nonNcm(8) / demand(8)) * 100, 0)
        cetakRateTotal(8) = Round((TotalSBE(8) / demand(8)) * 100, 0)
    End If
    
    If demand(9) > 0 Then
        cetakRateSameOutlet(9) = Round((sameOutlet(9) / demand(9)) * 100, 0)
        cetakRateNCM(9) = Round((ncm(9) / demand(9)) * 100, 0)
        cetakRateNonNCM(9) = Round((nonNcm(9) / demand(9)) * 100, 0)
        cetakRateTotal(9) = Round((TotalSBE(9) / demand(9)) * 100, 0)
    End If
    
    If demand(10) > 0 Then
        cetakRateSameOutlet(10) = Round((sameOutlet(10) / demand(10)) * 100, 0)
        cetakRateNCM(10) = Round((ncm(10) / demand(10)) * 100, 0)
        cetakRateNonNCM(10) = Round((nonNcm(10) / demand(10)) * 100, 0)
        cetakRateTotal(10) = Round((TotalSBE(10) / demand(10)) * 100, 0)
    End If
    
    ''--- AKHIR PENCARIAN NILAI RETENTION  -----
    'Set Report = New rptRepoSBRate1K
    Call createTableChart
    
    With rsChart
         
        .AddNew
        !keterangan = "Same Outlet"
        !K1 = cetakRateSameOutlet(0)
        !K10 = cetakRateSameOutlet(1)
        !K20 = cetakRateSameOutlet(2)
        !K30 = cetakRateSameOutlet(3)
        !K40 = cetakRateSameOutlet(4)
        !K50 = cetakRateSameOutlet(5)
        !K60 = cetakRateSameOutlet(6)
        !K70 = cetakRateSameOutlet(7)
        !K80 = cetakRateSameOutlet(8)
        !k90 = cetakRateSameOutlet(9)
        !k100 = cetakRateSameOutlet(10)
        .Update
        
        .AddNew
        !keterangan = "Claim from other NRM (NCM)"
        !K1 = cetakRateNCM(0)
        !K10 = cetakRateNCM(1)
        !K20 = cetakRateNCM(2)
        !K30 = cetakRateNCM(3)
        !K40 = cetakRateNCM(4)
        !K50 = cetakRateNCM(5)
        !K60 = cetakRateNCM(6)
        !K70 = cetakRateNCM(7)
        !K80 = cetakRateNCM(8)
        !k90 = cetakRateNCM(9)
        !k100 = cetakRateNCM(10)
        .Update
        
        .AddNew
        !keterangan = "Claim from other DLR (Non NCM)"
        !K1 = cetakRateNonNCM(0)
        !K10 = cetakRateNonNCM(1)
        !K20 = cetakRateNonNCM(2)
        !K30 = cetakRateNonNCM(3)
        !K40 = cetakRateNonNCM(4)
        !K50 = cetakRateNonNCM(5)
        !K60 = cetakRateNonNCM(6)
        !K70 = cetakRateNonNCM(7)
        !K80 = cetakRateNonNCM(8)
        !k90 = cetakRateNonNCM(9)
        !k100 = cetakRateNonNCM(10)
        .Update
    End With
    
    'Set rsChart = objO.FetchData("SELECT TOP 10 NO_CHASSIS,10 AS NILAI FROM JOBDTA")
    
   ' MSChart1.chartType = VtChChartType2dBar
    With MSChart1
        Set .DataSource = rsChart
        '.ChartData
         .ShowLegend = True
         
         .Stacking = True
    End With
    
    ProgressBar1K.value = 5
     
    Screen.MousePointer = vbDefault
    ProgressBar1K.Visible = False
    
    MSChart1.Visible = True
    btnTutup.Visible = True
    cmdExportExcel.Visible = True
   On Error GoTo 0
   Exit Sub

cmdCetakK_Click_Error:
    Screen.MousePointer = vbDefault
    ProgressBar1.Visible = False
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCetakK_Click of Form frmRepoSBRateRetention"
End Sub

Private Sub cmdCetakK_Click()
1        On Error GoTo cmdCetakK_Click_Error

2         If opTipe(0).value = False Then
3             If Month(Tgl1K) <> Month(Tgl2K) Or year(Tgl2K) <> year(Tgl2K) Then
4                 MsgBox "Range periode dalam 1 bulan", vbInformation
5                 Exit Sub
6             End If
7         End If
8         Screen.MousePointer = vbHourglass
          
9         ProgressBar1K.Min = 0
10        ProgressBar1K.Max = 10
11        ProgressBar1K.Visible = True
12        ProgressBar1K.value = 1
13        DoEvents
          
14        createTableLaporan
15        ProgressBar1K.value = 2
16        DoEvents
          
17        If opTipe(0).value = True Then
18            GoTo prosesTanggalService
19        ElseIf opTipe(1).value = True Then
20            GoTo prosesTanggalRetail
21        Else
22            GoTo prosesTanggalPDS
23        End If
          
          
prosesTanggalService:
          ''mencari unit
24        Call clearNilai
25        If Log.RPLHistori = 0 Then
26            sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan "
27            sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
28            sSQL = sSQL & " LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
29            sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
30            sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
31            sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B'"
32            sSQL = sSQL & " AND FS_NO_RETUR = ''  and FS_RC <> ''"
33            sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1K, "MM/DD/YYYY") & "' AND '" & Format(Tgl2K, "MM/DD/YYYY") & "'"
34            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
35            sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
36            sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end"
37            sSQL = sSQL & " order by FS_RC"
38        Else
39            sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan "
40            sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
41            sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
42            sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
43            sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
44            sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B'"
45            sSQL = sSQL & " AND FS_NO_RETUR = ''  and FS_RC <> ''"
46            sSQL = sSQL & " AND INVDATE between '" & Format(Tgl1K, "MM/DD/YYYY") & "' AND '" & Format(Tgl2K, "MM/DD/YYYY") & "'"
47            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100')"
48            sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
49            sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end"
50            sSQL = sSQL & " order by FS_RC"
51        End If
52        Set rsTemp = objO.FetchData(sSQL)
53        Cetak = 0
54        ProgressBar1K.value = 2
55        DoEvents
56        If Not rsTemp.EOF Then
              ''--- SBE 1K ------------------
57            rsTemp.filter = "fs_rc = 'R001' and keterangan = 'NCM'"
58            If Not rsTemp.EOF Then
59                ncm(0) = CDbl(rsTemp!total)
60            End If
61            rsTemp.filter = ""
62            rsTemp.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet'"
63            If Not rsTemp.EOF Then
64                sameOutlet(0) = CDbl(rsTemp!total)
65            End If
66            rsTemp.filter = ""
67            rsTemp.filter = "fs_rc = 'R001' and keterangan = 'Non NCM'"
68            If Not rsTemp.EOF Then
69                nonNcm(0) = CDbl(rsTemp!total)
70            End If
71            rsTemp.filter = ""
72            TotalSBE(0) = ncm(0) + nonNcm(0) + sameOutlet(0)
              ''--- AKHIR SBE 1K ------------
              
              ''--- SBE 10K ------------------
73            rsTemp.filter = "fs_rc = 'R010' and keterangan = 'NCM'"
74            If Not rsTemp.EOF Then
75                ncm(1) = CDbl(rsTemp!total)
76            End If
77            rsTemp.filter = ""
78            rsTemp.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet'"
79            If Not rsTemp.EOF Then
80                sameOutlet(1) = CDbl(rsTemp!total)
81            End If
82            rsTemp.filter = ""
83            rsTemp.filter = "fs_rc = 'R010' and keterangan = 'Non NCM'"
84            If Not rsTemp.EOF Then
85                nonNcm(1) = CDbl(rsTemp!total)
86            End If
87            rsTemp.filter = ""
88            TotalSBE(1) = ncm(1) + nonNcm(1) + sameOutlet(1)
              ''--- AKHIR SBE 10K ------------
              
              ''--- SBE 20K ------------------
89            rsTemp.filter = "fs_rc = 'R020' and keterangan = 'NCM'"
90            If Not rsTemp.EOF Then
91                ncm(2) = CDbl(rsTemp!total)
92            End If
93            rsTemp.filter = ""
94            rsTemp.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet'"
95            If Not rsTemp.EOF Then
96                sameOutlet(2) = CDbl(rsTemp!total)
97            End If
98            rsTemp.filter = ""
99            rsTemp.filter = "fs_rc = 'R020' and keterangan = 'Non NCM'"
100           If Not rsTemp.EOF Then
101               nonNcm(2) = CDbl(rsTemp!total)
102           End If
103           rsTemp.filter = ""
104           TotalSBE(2) = ncm(2) + nonNcm(2) + sameOutlet(2)
              ''--- AKHIR SBE 20K ------------
              
              ''--- SBE 30K ------------------
105           rsTemp.filter = "fs_rc = 'R030' and keterangan = 'NCM'"
106           If Not rsTemp.EOF Then
107               ncm(3) = CDbl(rsTemp!total)
108           End If
109           rsTemp.filter = ""
110           rsTemp.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet'"
111           If Not rsTemp.EOF Then
112               sameOutlet(3) = CDbl(rsTemp!total)
113           End If
114           rsTemp.filter = ""
115           rsTemp.filter = "fs_rc = 'R030' and keterangan = 'Non NCM'"
116           If Not rsTemp.EOF Then
117               nonNcm(3) = CDbl(rsTemp!total)
118           End If
119           rsTemp.filter = ""
120           TotalSBE(3) = ncm(3) + nonNcm(3) + sameOutlet(3)
              ''--- AKHIR SBE 30K ------------
              
              ''--- SBE 40K ------------------
121           rsTemp.filter = "fs_rc = 'R040' and keterangan = 'NCM'"
122           If Not rsTemp.EOF Then
123               ncm(4) = CDbl(rsTemp!total)
124           End If
125           rsTemp.filter = ""
126           rsTemp.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet'"
127           If Not rsTemp.EOF Then
128               sameOutlet(4) = CDbl(rsTemp!total)
129           End If
130           rsTemp.filter = ""
131           rsTemp.filter = "fs_rc = 'R040' and keterangan = 'Non NCM'"
132           If Not rsTemp.EOF Then
133               nonNcm(4) = CDbl(rsTemp!total)
134           End If
135           rsTemp.filter = ""
136           TotalSBE(4) = ncm(4) + nonNcm(4) + sameOutlet(4)
              ''--- AKHIR SBE 40K ------------
              
              ''--- SBE 50K ------------------
137           rsTemp.filter = "fs_rc = 'R050' and keterangan = 'NCM'"
138           If Not rsTemp.EOF Then
139               ncm(5) = CDbl(rsTemp!total)
140           End If
141           rsTemp.filter = ""
142           rsTemp.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet'"
143           If Not rsTemp.EOF Then
144               sameOutlet(5) = CDbl(rsTemp!total)
145           End If
146           rsTemp.filter = ""
147           rsTemp.filter = "fs_rc = 'R050' and keterangan = 'Non NCM'"
148           If Not rsTemp.EOF Then
149               nonNcm(5) = CDbl(rsTemp!total)
150           End If
151           rsTemp.filter = ""
152           TotalSBE(5) = ncm(5) + nonNcm(5) + sameOutlet(5)
              ''--- AKHIR SBE 50K ------------
              
              ''--- SBE 60K ------------------
153           rsTemp.filter = "fs_rc = 'R060' and keterangan = 'NCM'"
154           If Not rsTemp.EOF Then
155               ncm(6) = CDbl(rsTemp!total)
156           End If
157           rsTemp.filter = ""
158           rsTemp.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet'"
159           If Not rsTemp.EOF Then
160               sameOutlet(6) = CDbl(rsTemp!total)
161           End If
162           rsTemp.filter = ""
163           rsTemp.filter = "fs_rc = 'R060' and keterangan = 'Non NCM'"
164           If Not rsTemp.EOF Then
165               nonNcm(6) = CDbl(rsTemp!total)
166           End If
167           rsTemp.filter = ""
168           TotalSBE(6) = ncm(6) + nonNcm(6) + sameOutlet(6)
              ''--- AKHIR SBE 60K ------------
              
              ''--- SBE 70K ------------------
169           rsTemp.filter = "fs_rc = 'R070' and keterangan = 'NCM'"
170           If Not rsTemp.EOF Then
171               ncm(7) = CDbl(rsTemp!total)
172           End If
173           rsTemp.filter = ""
174           rsTemp.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet'"
175           If Not rsTemp.EOF Then
176               sameOutlet(7) = CDbl(rsTemp!total)
177           End If
178           rsTemp.filter = ""
179           rsTemp.filter = "fs_rc = 'R070' and keterangan = 'Non NCM'"
180           If Not rsTemp.EOF Then
181               nonNcm(7) = CDbl(rsTemp!total)
182           End If
183           rsTemp.filter = ""
184           TotalSBE(7) = ncm(7) + nonNcm(7) + sameOutlet(7)
              ''--- AKHIR SBE 70K ------------
              
              ''--- SBE 80K ------------------
185           rsTemp.filter = "fs_rc = 'R080' and keterangan = 'NCM'"
186           If Not rsTemp.EOF Then
187               ncm(8) = CDbl(rsTemp!total)
188           End If
189           rsTemp.filter = ""
190           rsTemp.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet'"
191           If Not rsTemp.EOF Then
192               sameOutlet(8) = CDbl(rsTemp!total)
193           End If
194           rsTemp.filter = ""
195           rsTemp.filter = "fs_rc = 'R080' and keterangan = 'Non NCM'"
196           If Not rsTemp.EOF Then
197               nonNcm(8) = CDbl(rsTemp!total)
198           End If
199           rsTemp.filter = ""
200           TotalSBE(8) = ncm(8) + nonNcm(8) + sameOutlet(8)
              ''--- AKHIR SBE 80K ------------
              
              ''--- SBE 90K ------------------
201           rsTemp.filter = "fs_rc = 'R090' and keterangan = 'NCM'"
202           If Not rsTemp.EOF Then
203               ncm(9) = CDbl(rsTemp!total)
204           End If
205           rsTemp.filter = ""
206           rsTemp.filter = "fs_rc = 'R090' and keterangan = 'Same Outlet'"
207           If Not rsTemp.EOF Then
208               sameOutlet(9) = CDbl(rsTemp!total)
209           End If
210           rsTemp.filter = ""
211           rsTemp.filter = "fs_rc = 'R090' and keterangan = 'Non NCM'"
212           If Not rsTemp.EOF Then
213               nonNcm(9) = CDbl(rsTemp!total)
214           End If
215           rsTemp.filter = ""
216           TotalSBE(9) = ncm(9) + nonNcm(9) + sameOutlet(9)
              ''--- AKHIR SBE 90K ------------
              
              ''--- SBE 100K ------------------
217           rsTemp.filter = "fs_rc = 'R100' and keterangan = 'NCM'"
218           If Not rsTemp.EOF Then
219               ncm(10) = CDbl(rsTemp!total)
220           End If
221           rsTemp.filter = ""
222           rsTemp.filter = "fs_rc = 'R100' and keterangan = 'Same Outlet'"
223           If Not rsTemp.EOF Then
224               sameOutlet(10) = CDbl(rsTemp!total)
225           End If
226           rsTemp.filter = ""
227           rsTemp.filter = "fs_rc = 'R100' and keterangan = 'Non NCM'"
228           If Not rsTemp.EOF Then
229               nonNcm(10) = CDbl(rsTemp!total)
230           End If
231           rsTemp.filter = ""
232           TotalSBE(10) = ncm(10) + nonNcm(10) + sameOutlet(10)
              ''--- AKHIR SBE 100K ------------
         
233       End If
          
234       ProgressBar1K.value = 3
235       DoEvents
          
          ''--- MENCARI NILAI DEMAND UNTUK SETIAP SBE
          'demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Format(Tgl1K, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(Format(Tgl2K, "MM/DD/YYYY")), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
          ''20180323 '' REV '' NILAI DEMAND DI AMBIL DARI TANGGAL YANG DIPILIH
236       demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and (dateadd(day,-1,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
237       demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
238       demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
239       demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
240       demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
241       demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
242       demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
243       demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
244       demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
245       demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-52,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
246       demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-58,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
          
      '    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and (dateadd(day,-1,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-52,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-58,'" & Format(Month(isnow), "0#") & "/01/" & year(isnow) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
          
          
          ''--- AKHIR NILAI DEMAND -------------------
247       GoTo prosesCetak
          
prosesTanggalRetail:
          ''2017-08-11
          ''mencari data service dari tanggal retail
          ''yang diambil data service kelipatan sbe masuk range tanggal retail
          
248       Call clearNilai
249       If Log.RPLHistori = 0 Then
250           sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
251           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
252           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
253           sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
254           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
255           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
256           sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(Tgl1K, "mm/dd/yyyy") & "' AND '" & Format(Tgl2K, "mm/dd/yyyy") & "' )"
257           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
258           sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
259           sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
              'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
260           sSQL = sSQL & " order by FS_RC"
261       Else
262           sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
263           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
264           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
265           sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
266           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA  and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
267           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
268           sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(Tgl1K, "mm/dd/yyyy") & "' AND '" & Format(Tgl2K, "mm/dd/yyyy") & "' )"
269           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
270           sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
271           sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
              'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
272           sSQL = sSQL & " order by FS_RC"
273       End If
274       Set rsTemp = objO.FetchData(sSQL)
275       ProgressBar1K.value = 2
276       DoEvents
          
277       If Not rsTemp.EOF Then
278           rsTemp.MoveFirst
279           Do While Not rsTemp.EOF
280               If Nul(rsTemp!FS_RC) = "R001" Then
281                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
282                       sameOutlet(0) = CDbl(rsTemp!total)
283                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
284                       ncm(0) = CDbl(rsTemp!total)
285                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
286                       nonNcm(0) = CDbl(rsTemp!total)
287                   End If
288               ElseIf Nul(rsTemp!FS_RC) = "R010" Then
289                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
290                       sameOutlet(1) = CDbl(rsTemp!total)
291                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
292                       ncm(1) = CDbl(rsTemp!total)
293                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
294                       nonNcm(1) = CDbl(rsTemp!total)
295                   End If
296               ElseIf Nul(rsTemp!FS_RC) = "R020" Then
297                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
298                       sameOutlet(2) = CDbl(rsTemp!total)
299                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
300                       ncm(2) = CDbl(rsTemp!total)
301                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
302                       nonNcm(2) = CDbl(rsTemp!total)
303                   End If
304               ElseIf Nul(rsTemp!FS_RC) = "R030" Then
305                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
306                       sameOutlet(3) = CDbl(rsTemp!total)
307                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
308                       ncm(3) = CDbl(rsTemp!total)
309                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
310                       nonNcm(3) = CDbl(rsTemp!total)
311                   End If
312               ElseIf Nul(rsTemp!FS_RC) = "R040" Then
313                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
314                       sameOutlet(4) = CDbl(rsTemp!total)
315                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
316                       ncm(4) = CDbl(rsTemp!total)
317                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
318                       nonNcm(4) = CDbl(rsTemp!total)
319                   End If
320               ElseIf Nul(rsTemp!FS_RC) = "R050" Then
321                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
322                       sameOutlet(5) = CDbl(rsTemp!total)
323                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
324                       ncm(5) = CDbl(rsTemp!total)
325                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
326                       nonNcm(5) = CDbl(rsTemp!total)
327                   End If
328               ElseIf Nul(rsTemp!FS_RC) = "R060" Then
329                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
330                       sameOutlet(6) = CDbl(rsTemp!total)
331                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
332                       ncm(6) = CDbl(rsTemp!total)
333                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
334                       nonNcm(6) = CDbl(rsTemp!total)
335                   End If
336               ElseIf Nul(rsTemp!FS_RC) = "R070" Then
337                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
338                       sameOutlet(7) = CDbl(rsTemp!total)
339                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
340                       ncm(7) = CDbl(rsTemp!total)
341                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
342                       nonNcm(7) = CDbl(rsTemp!total)
343                   End If
344               ElseIf Nul(rsTemp!FS_RC) = "R080" Then
345                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
346                       sameOutlet(8) = CDbl(rsTemp!total)
347                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
348                       ncm(8) = CDbl(rsTemp!total)
349                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
350                       nonNcm(8) = CDbl(rsTemp!total)
351                   End If
352               ElseIf Nul(rsTemp!FS_RC) = "R090" Then
353                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
354                       sameOutlet(9) = CDbl(rsTemp!total)
355                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
356                       ncm(9) = CDbl(rsTemp!total)
357                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
358                       nonNcm(9) = CDbl(rsTemp!total)
359                   End If
360               ElseIf Nul(rsTemp!FS_RC) = "R100" Then
361                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
362                       sameOutlet(10) = CDbl(rsTemp!total)
363                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
364                       ncm(10) = CDbl(rsTemp!total)
365                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
366                       nonNcm(10) = CDbl(rsTemp!total)
367                   End If
368               End If
369               rsTemp.MoveNext
370           Loop
371       End If
          
372       ProgressBar1K.value = 3
373       DoEvents
          
          ''--- MENCARI NILAI DEMAND BERDASARKAN TANGGAL RETAIL
          
      '    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 0 and 1) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
          
          ''demand diambil dari tanggal retail + 1 -1 sesuai KM
          
374       demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and (dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
375       demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
376       demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
377       demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
378       demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
379       demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
380       demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-33,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
381       demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
382       demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
383       demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-52,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
384       demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-58,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
          
          
      '    For i = 1 To 10
      '        demand(i) = demand(0)
      '    Next
          
          ''--- AKHIR NILAI DEMAND -------------------
385       GoTo prosesCetak
          
prosesTanggalPDS:
          ''2017-08-11
          ''mencari data service dari tanggal retail
          ''yang diambil data service kelipatan sbe masuk range tanggal retail
          
386       Call clearNilai
387       If Log.RPLHistori = 0 Then
388           sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
389           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
390           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
391           sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
392           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
393           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
394           sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "')"
395           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
396           sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
397           sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
              'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
398           sSQL = sSQL & " order by FS_RC"
399       Else
400           sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
401           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end as kelipatanBulan,"
402           sSQL = sSQL & " case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end as AKHIRBulan,"
403           sSQL = sSQL & " DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) selisih"
404           sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
405           sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
406           sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "')"
407           sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') and t.TPSVC <> 'w' and t.PDS = 0"
408           sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end ,"
409           sSQL = sSQL & " SUBSTRING(FS_RC,2,2)"
              'sSQL = sSQL & " HAVING DATEDIFF(MONTH,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) BETWEEN case when SUBSTRING(FS_RC,2,2) = '00' then 1 else SUBSTRING(FS_RC,2,2) *6 end and case when SUBSTRING(FS_RC,2,2) = '00' then 1 else (SUBSTRING(FS_RC,2,2) *6)+5 end"
410           sSQL = sSQL & " order by FS_RC"
411       End If
412       Set rsTemp = objO.FetchData(sSQL)
413       ProgressBar1K.value = 2
414       DoEvents
          
415       If Not rsTemp.EOF Then
416           rsTemp.MoveFirst
417           Do While Not rsTemp.EOF
418               If Nul(rsTemp!FS_RC) = "R001" Then
419                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
420                       sameOutlet(0) = CDbl(rsTemp!total)
421                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
422                       ncm(0) = CDbl(rsTemp!total)
423                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
424                       nonNcm(0) = CDbl(rsTemp!total)
425                   End If
426               ElseIf Nul(rsTemp!FS_RC) = "R010" Then
427                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
428                       sameOutlet(1) = CDbl(rsTemp!total)
429                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
430                       ncm(1) = CDbl(rsTemp!total)
431                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
432                       nonNcm(1) = CDbl(rsTemp!total)
433                   End If
434               ElseIf Nul(rsTemp!FS_RC) = "R020" Then
435                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
436                       sameOutlet(2) = CDbl(rsTemp!total)
437                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
438                       ncm(2) = CDbl(rsTemp!total)
439                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
440                       nonNcm(2) = CDbl(rsTemp!total)
441                   End If
442               ElseIf Nul(rsTemp!FS_RC) = "R030" Then
443                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
444                       sameOutlet(3) = CDbl(rsTemp!total)
445                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
446                       ncm(3) = CDbl(rsTemp!total)
447                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
448                       nonNcm(3) = CDbl(rsTemp!total)
449                   End If
450               ElseIf Nul(rsTemp!FS_RC) = "R040" Then
451                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
452                       sameOutlet(4) = CDbl(rsTemp!total)
453                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
454                       ncm(4) = CDbl(rsTemp!total)
455                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
456                       nonNcm(4) = CDbl(rsTemp!total)
457                   End If
458               ElseIf Nul(rsTemp!FS_RC) = "R050" Then
459                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
460                       sameOutlet(5) = CDbl(rsTemp!total)
461                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
462                       ncm(5) = CDbl(rsTemp!total)
463                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
464                       nonNcm(5) = CDbl(rsTemp!total)
465                   End If
466               ElseIf Nul(rsTemp!FS_RC) = "R060" Then
467                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
468                       sameOutlet(6) = CDbl(rsTemp!total)
469                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
470                       ncm(6) = CDbl(rsTemp!total)
471                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
472                       nonNcm(6) = CDbl(rsTemp!total)
473                   End If
474               ElseIf Nul(rsTemp!FS_RC) = "R070" Then
475                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
476                       sameOutlet(7) = CDbl(rsTemp!total)
477                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
478                       ncm(7) = CDbl(rsTemp!total)
479                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
480                       nonNcm(7) = CDbl(rsTemp!total)
481                   End If
482               ElseIf Nul(rsTemp!FS_RC) = "R080" Then
483                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
484                       sameOutlet(8) = CDbl(rsTemp!total)
485                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
486                       ncm(8) = CDbl(rsTemp!total)
487                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
488                       nonNcm(8) = CDbl(rsTemp!total)
489                   End If
490               ElseIf Nul(rsTemp!FS_RC) = "R090" Then
491                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
492                       sameOutlet(9) = CDbl(rsTemp!total)
493                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
494                       ncm(9) = CDbl(rsTemp!total)
495                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
496                       nonNcm(9) = CDbl(rsTemp!total)
497                   End If
498               ElseIf Nul(rsTemp!FS_RC) = "R100" Then
499                   If Nul(rsTemp!keterangan) = "Same Outlet" Then
500                       sameOutlet(10) = CDbl(rsTemp!total)
501                   ElseIf Nul(rsTemp!keterangan) = "NCM" Then
502                       ncm(10) = CDbl(rsTemp!total)
503                   ElseIf Nul(rsTemp!keterangan) = "Non NCM" Then
504                       nonNcm(10) = CDbl(rsTemp!total)
505                   End If
506               End If
507               rsTemp.MoveNext
508           Loop
509       End If
          
510       ProgressBar1K.value = 3
511       DoEvents
          
          ''--- MENCARI NILAI DEMAND BERDASARKAN TANGGAL RETAIL
          'demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 0 and 1) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 2 and 6) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 7 and 12) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 13 and 18) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 19 and 24) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 25 and 30) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 31 and 36) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 37 and 42) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 43 and 48) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 49 and 54) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
      '    demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where (DATEDIFF(month,'" & Format(Tgl1K, "mm/dd/yyyy") & "',getdate()) between 55 and 60) and TglFaktur between '" & Format(Tgl1K, "mm/dd/yyyy") & "' and '" & Format(Tgl2K, "mm/dd/yyyy") & "' group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
          
512       demand(0) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and (dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
513       demand(1) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
514       demand(2) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
515       demand(3) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
516       demand(4) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
517       demand(5) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
518       demand(6) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
519       demand(7) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
520       demand(8) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
521       demand(9) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-52,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
522       demand(10) = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select (sum(case typetrn when 'A' then 1 else 0 end) - sum(case typetrn when 'K' then 1 else 0 end)) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & Format(Month(Tgl1K), "0#") & "/01/" & year(Tgl1K) & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-58,'" & Format(Month(Tgl2K), "0#") & "/01/" & year(Tgl2K) & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")

          
          ''--- AKHIR NILAI DEMAND -------------------
523       GoTo prosesCetak
          

prosesCetak:
              
524       ProgressBar1K.value = 4
525       DoEvents
          
          ''--- MENCARI NILAI RETENTION TIAP SBE -----
          
526       If demand(0) > 0 Then
527           cetakRateSameOutlet(0) = Round((sameOutlet(0) / demand(0)) * 100, 0)
528           cetakRateNCM(0) = Round((ncm(0) / demand(0)) * 100, 0)
529           cetakRateNonNCM(0) = Round((nonNcm(0) / demand(0)) * 100, 0)
530           cetakRateTotal(0) = Round((TotalSBE(0) / demand(0)) * 100, 0)
531       End If
          
532       If demand(1) > 0 Then
533           cetakRateSameOutlet(1) = Round((sameOutlet(1) / demand(1)) * 100, 0)
534           cetakRateNCM(1) = Round((ncm(1) / demand(1)) * 100, 0)
535           cetakRateNonNCM(1) = Round((nonNcm(1) / demand(1)) * 100, 0)
536           cetakRateTotal(1) = Round((TotalSBE(1) / demand(1)) * 100, 0)
537       End If
          
538       If demand(2) > 0 Then
539           cetakRateSameOutlet(2) = Round((sameOutlet(2) / demand(2)) * 100, 0)
540           cetakRateNCM(2) = Round((ncm(2) / demand(2)) * 100, 0)
541           cetakRateNonNCM(2) = Round((nonNcm(2) / demand(2)) * 100, 0)
542           cetakRateTotal(2) = Round((TotalSBE(2) / demand(2)) * 100, 0)
543       End If
          
544       If demand(3) > 0 Then
545           cetakRateSameOutlet(3) = Round((sameOutlet(3) / demand(3)) * 100, 0)
546           cetakRateNCM(3) = Round((ncm(3) / demand(3)) * 100, 0)
547           cetakRateNonNCM(3) = Round((nonNcm(3) / demand(3)) * 100, 0)
548           cetakRateTotal(3) = Round((TotalSBE(3) / demand(3)) * 100, 0)
549       End If
          
550       If demand(4) > 0 Then
551           cetakRateSameOutlet(4) = Round((sameOutlet(4) / demand(4)) * 100, 0)
552           cetakRateNCM(4) = Round((ncm(4) / demand(4)) * 100, 0)
553           cetakRateNonNCM(4) = Round((nonNcm(4) / demand(4)) * 100, 0)
554           cetakRateTotal(4) = Round((TotalSBE(4) / demand(4)) * 100, 0)
555       End If
          
556       If demand(5) > 0 Then
557           cetakRateSameOutlet(5) = Round((sameOutlet(5) / demand(5)) * 100, 0)
558           cetakRateNCM(5) = Round((ncm(5) / demand(5)) * 100, 0)
559           cetakRateNonNCM(5) = Round((nonNcm(5) / demand(5)) * 100, 0)
560           cetakRateTotal(5) = Round((TotalSBE(5) / demand(5)) * 100, 0)
561       End If
          
562       If demand(6) > 0 Then
563           cetakRateSameOutlet(6) = Round((sameOutlet(6) / demand(6)) * 100, 0)
564           cetakRateNCM(6) = Round((ncm(6) / demand(6)) * 100, 0)
565           cetakRateNonNCM(6) = Round((nonNcm(6) / demand(6)) * 100, 0)
566           cetakRateTotal(6) = Round((TotalSBE(6) / demand(6)) * 100, 0)
567       End If
          
568       If demand(7) > 0 Then
569           cetakRateSameOutlet(7) = Round((sameOutlet(7) / demand(7)) * 100, 0)
570           cetakRateNCM(7) = Round((ncm(7) / demand(7)) * 100, 0)
571           cetakRateNonNCM(7) = Round((nonNcm(7) / demand(7)) * 100, 0)
572           cetakRateTotal(7) = Round((TotalSBE(7) / demand(7)) * 100, 0)
573       End If
          
574       If demand(8) > 0 Then
575           cetakRateSameOutlet(8) = Round((sameOutlet(8) / demand(8)) * 100, 0)
576           cetakRateNCM(8) = Round((ncm(8) / demand(8)) * 100, 0)
577           cetakRateNonNCM(8) = Round((nonNcm(8) / demand(8)) * 100, 0)
578           cetakRateTotal(8) = Round((TotalSBE(8) / demand(8)) * 100, 0)
579       End If
          
580       If demand(9) > 0 Then
581           cetakRateSameOutlet(9) = Round((sameOutlet(9) / demand(9)) * 100, 0)
582           cetakRateNCM(9) = Round((ncm(9) / demand(9)) * 100, 0)
583           cetakRateNonNCM(9) = Round((nonNcm(9) / demand(9)) * 100, 0)
584           cetakRateTotal(9) = Round((TotalSBE(9) / demand(9)) * 100, 0)
585       End If
          
586       If demand(10) > 0 Then
587           cetakRateSameOutlet(10) = Round((sameOutlet(10) / demand(10)) * 100, 0)
588           cetakRateNCM(10) = Round((ncm(10) / demand(10)) * 100, 0)
589           cetakRateNonNCM(10) = Round((nonNcm(10) / demand(10)) * 100, 0)
590           cetakRateTotal(10) = Round((TotalSBE(10) / demand(10)) * 100, 0)
591       End If
          
          ''--- AKHIR PENCARIAN NILAI RETENTION  -----
592       Set Report = New rptRepoSBRate1K
          
593       ProgressBar1K.value = 5
594       DoEvents
          
595       Report.txtRateSO1.SetText cetakRateSameOutlet(0) & " %"
596       Report.txtRateSO10.SetText cetakRateSameOutlet(1) & " %"
597       Report.txtRateSO20.SetText cetakRateSameOutlet(2) & " %"
598       Report.txtRateSO30.SetText cetakRateSameOutlet(3) & " %"
599       Report.txtRateSO40.SetText cetakRateSameOutlet(4) & " %"
600       Report.txtRateSO50.SetText cetakRateSameOutlet(5) & " %"
601       Report.txtRateSO60.SetText cetakRateSameOutlet(6) & " %"
602       Report.txtRateSO70.SetText cetakRateSameOutlet(7) & " %"
603       Report.txtRateSO80.SetText cetakRateSameOutlet(8) & " %"
604       Report.txtRateSO90.SetText cetakRateSameOutlet(9) & " %"
605       Report.txtRateSO100.SetText cetakRateSameOutlet(10) & " %"
          
          
606       Report.txtRateNRM1.SetText cetakRateNCM(0) & " %"
607       Report.txtRateNRM10.SetText cetakRateNCM(1) & " %"
608       Report.txtRateNRM20.SetText cetakRateNCM(2) & " %"
609       Report.txtRateNRM30.SetText cetakRateNCM(3) & " %"
610       Report.txtRateNRM40.SetText cetakRateNCM(4) & " %"
611       Report.txtRateNRM50.SetText cetakRateNCM(5) & " %"
612       Report.txtRateNRM60.SetText cetakRateNCM(6) & " %"
613       Report.txtRateNRM70.SetText cetakRateNCM(7) & " %"
614       Report.txtRateNRM80.SetText cetakRateNCM(8) & " %"
615       Report.txtRateNRM90.SetText cetakRateNCM(9) & " %"
616       Report.txtRateNRM100.SetText cetakRateNCM(10) & " %"
           
617       Report.txtRateNonNRM1.SetText cetakRateNonNCM(0) & " %"
618       Report.txtRateNonNRM10.SetText cetakRateNonNCM(1) & " %"
619       Report.txtRateNonNRM20.SetText cetakRateNonNCM(2) & " %"
620       Report.txtRateNonNRM30.SetText cetakRateNonNCM(3) & " %"
621       Report.txtRateNonNRM40.SetText cetakRateNonNCM(4) & " %"
622       Report.txtRateNonNRM50.SetText cetakRateNonNCM(5) & " %"
623       Report.txtRateNonNRM60.SetText cetakRateNonNCM(6) & " %"
624       Report.txtRateNonNRM70.SetText cetakRateNonNCM(7) & " %"
625       Report.txtRateNonNRM80.SetText cetakRateNonNCM(8) & " %"
626       Report.txtRateNonNRM90.SetText cetakRateNonNCM(9) & " %"
627       Report.txtRateNonNRM100.SetText cetakRateNonNCM(10) & " %"
           
628       ProgressBar1K.value = 6
629       DoEvents
          
630       Report.txtRateTotal1.SetText cetakRateTotal(0) & " %"
631       Report.txtRateTotal10.SetText cetakRateTotal(1) & " %"
632       Report.txtRateTotal20.SetText cetakRateTotal(2) & " %"
633       Report.txtRateTotal30.SetText cetakRateTotal(3) & " %"
634       Report.txtRateTotal40.SetText cetakRateTotal(4) & " %"
635       Report.txtRateTotal50.SetText cetakRateTotal(5) & " %"
636       Report.txtRateTotal60.SetText cetakRateTotal(6) & " %"
637       Report.txtRateTotal70.SetText cetakRateTotal(7) & " %"
638       Report.txtRateTotal80.SetText cetakRateTotal(8) & " %"
639       Report.txtRateTotal90.SetText cetakRateTotal(9) & " %"
640       Report.txtRateTotal100.SetText cetakRateTotal(10) & " %"
          
          ''--- NILAI SBE ---------------------------------
          
641       Report.txtValSO1.SetText sameOutlet(0)
642       Report.txtValSO10.SetText sameOutlet(1)
643       Report.txtValSO20.SetText sameOutlet(2)
644       Report.txtValSO30.SetText sameOutlet(3)
645       Report.txtValSO40.SetText sameOutlet(4)
646       Report.txtValSO50.SetText sameOutlet(5)
647       Report.txtValSO60.SetText sameOutlet(6)
648       Report.txtValSO70.SetText sameOutlet(7)
649       Report.txtValSO80.SetText sameOutlet(8)
650       Report.txtValSO90.SetText sameOutlet(9)
651       Report.txtValSO100.SetText sameOutlet(10)
          
652       ProgressBar1K.value = 7
653       DoEvents
654       Report.txtValNCM1.SetText ncm(0)
655       Report.txtValNCM10.SetText ncm(1)
656       Report.txtValNCM20.SetText ncm(2)
657       Report.txtValNCM30.SetText ncm(3)
658       Report.txtValNCM40.SetText ncm(4)
659       Report.txtValNCM50.SetText ncm(5)
660       Report.txtValNCM60.SetText ncm(6)
661       Report.txtValNCM70.SetText ncm(7)
662       Report.txtValNCM80.SetText ncm(8)
663       Report.txtValNCM90.SetText ncm(9)
664       Report.txtValNCM100.SetText ncm(10)
           
665       Report.txtValNonNCM1.SetText nonNcm(0)
666       Report.txtValNonNCM10.SetText nonNcm(1)
667       Report.txtValNonNCM20.SetText nonNcm(2)
668       Report.txtValNonNCM30.SetText nonNcm(3)
669       Report.txtValNonNCM40.SetText nonNcm(4)
670       Report.txtValNonNCM50.SetText nonNcm(5)
671       Report.txtValNonNCM60.SetText nonNcm(6)
672       Report.txtValNonNCM70.SetText nonNcm(7)
673       Report.txtValNonNCM80.SetText nonNcm(8)
674       Report.txtValNonNCM90.SetText nonNcm(9)
675       Report.txtValNonNCM100.SetText nonNcm(10)
          
676       ProgressBar1K.value = 8
677       DoEvents
          
678       Report.txtValTotal1.SetText TotalSBE(0)
679       Report.txtValTotal10.SetText TotalSBE(1)
680       Report.txtValTotal20.SetText TotalSBE(2)
681       Report.txtValTotal30.SetText TotalSBE(3)
682       Report.txtValTotal40.SetText TotalSBE(4)
683       Report.txtValTotal50.SetText TotalSBE(5)
684       Report.txtValTotal60.SetText TotalSBE(6)
685       Report.txtValTotal70.SetText TotalSBE(7)
686       Report.txtValTotal80.SetText TotalSBE(8)
687       Report.txtValTotal90.SetText TotalSBE(9)
688       Report.txtValTotal100.SetText TotalSBE(10)
          
          ''--- AKHIR SBE ---------------------------------
689       ProgressBar1K.value = 9
690       DoEvents
          ''--- CETAK DEMAND ------------------------------
691       Report.txtDemand1.SetText demand(0)
692       Report.txtDemand10.SetText demand(1)
693       Report.txtDemand20.SetText demand(2)
694       Report.txtDemand30.SetText demand(3)
695       Report.txtDemand40.SetText demand(4)
696       Report.txtDemand50.SetText demand(5)
697       Report.txtDemand60.SetText demand(6)
698       Report.txtDemand70.SetText demand(7)
699       Report.txtDemand80.SetText demand(8)
700       Report.txtDemand90.SetText demand(9)
701       Report.txtDemand100.SetText demand(10)
          ''--- AKHIR DEMAND ------------------------------
          
702       ProgressBar1K.value = 10
703       DoEvents
704       ProgressBar1K.Visible = False
705       cmdCetak.Enabled = True
          
          ''
706       Report.Text11.SetText "Laporan Retention Rate " & gPerush
707       Report.Text9.SetText Format(Day(Tgl1K.value), "0#") & "-" & GetNamaBulan(Month(Tgl1K), 1) & "-" & year(Tgl1K) & " S/D " & Format(Day(Tgl2K.value), "0#") & "-" & GetNamaBulan(Month(Tgl2K), 1) & "-" & year(Tgl2K)
708       Report.Text29.SetText "Laporan Total SBE " & gPerush
           
709       With CRViewer1
710           .ReportSource = Report
711           .ViewReport
712           .Zoom 100
713           .Visible = True
714           .Left = 300
715           .Height = Me.Height - 400
              '.Width = Me.Width - .Left
716       End With
         
717       cmdTutup.Visible = True
718       cmdTutup.Left = CRViewer1.Left + 20
719       cmdTutup.Top = CRViewer1.Top + 400
720       Screen.MousePointer = vbDefault
721       ProgressBar1K.Visible = False
          
722       Call DisableExport
723       Call setLokasiTombol
          
724      On Error GoTo 0
725      Exit Sub

cmdCetakK_Click_Error:
726       Screen.MousePointer = vbDefault
727       ProgressBar1.Visible = False
728       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdCetakK_Click of Form frmRepoSBRateRetention"
End Sub

Private Sub cmdDownloadVin_Click()
1     On Error Resume Next
          Dim rsNorangka      As ADODB.Recordset
          Dim rsData          As ADODB.Recordset
          Dim isDownloadAll   As Boolean
          
2         If Month(tglAwalVin) <> Month(tglAkhirVin) Or year(tglAwalVin) <> year(tglAkhirVin) Then
3             MsgBox "Range periode dalam 1 bulan", vbInformation
4             Exit Sub
5         End If
6         isDownloadAll = False
7         If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Then
8             Call createTableDownload(True)
9             isDownloadAll = True
10        ElseIf gTipeUser = "SA" Or gTipeUser = "MRS" Then
11            Call createTableDownload(False)
12            isDownloadAll = False
13        ElseIf gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
14            Call createTableDownload(True)
15            isDownloadAll = True
16        Else
17            MsgBox "Tipe user anda tidak diperbolehkan mengakses fitur ini", vbInformation
18            Exit Sub
19        End If
          
20        sSQL = "select b.NoRangka,b.kode_sales,case when b.Syarat_Bayar = 0 then 'TUNAI' else 'KREDIT' end bayarcust,b.Nama_Faktur,b.Alamat1_Faktur,b.Kota,b.Kecamatan,b.Kode_Pos,b.HP1,b.HP2,f.propinsi"
21        sSQL = sSQL & " from [SPK" & Right(gDb, 2) & "]..Sihdr a"
22        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1"
23        sSQL = sSQL & " left join SingleIDOri..Bridging e on a.KodeLgn = e.KodeLgn and e.KodeDealer = '" & Right(gDb, 2) & "'"
24        sSQL = sSQL & " left join SingleIDOri..NewCustomer f on e.IDCust = f.IDCust"
25        sSQL = sSQL & " left join SingleIDOri..UnitCustomer g on b.NoRangka = g.NoRangka"
26        sSQL = sSQL & " left join SingleIDOri..NewCustomer H on h.IDCust = g.IDCustSTNK"
27        sSQL = sSQL & " where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "'))))"
28        sSQL = sSQL & " and b.NoRangka not in ("
29        sSQL = sSQL & " select no_chassis from T_INVOICE where  NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
30        sSQL = sSQL & " and FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
31        sSQL = sSQL & " and TPSVC <> 'w' and PDS = 0"
32        sSQL = sSQL & " )"
          ''sSQL = sSQL & " group by b.NoRangka"
          
33        Set rsNorangka = objO.FetchData(sSQL)
34        If rsNorangka.RecordCount = 0 Then
35            MsgBox "Tidak ada data untuk di download", vbInformation
36            Exit Sub
37        Else
38            rsNorangka.MoveFirst
39            ProgVin.Max = rsNorangka.RecordCount + 1
40            ProgVin.value = 1
41            ProgVin.Visible = True
42            DoEvents
43            Do While Not rsNorangka.EOF
44                sSQL = "select"
45                sSQL = sSQL & " ISNULL((select top 1 tglfollow from m_listremindercall where norangka = a.NoRangka order by tglfollow desc),'1900-01-01') lastreminder,"
46                sSQL = sSQL & " b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3 as namacustomer,"
47                sSQL = sSQL & " a.Fmdl as namamodel,a.NoPolisi,a.NoRangka,a.Thmobil,a.Ddate,"
48                sSQL = sSQL & " ISNULL((select top 1 TglFaktur from [SPK" & Right(gDb, 2) & "]..Sihdr where NoRangka = a.NoRangka order by TglFaktur desc),'1900-01-01') as tglretail,"
49                sSQL = sSQL & " a.Mdl,b.Alamat,b.Kelurahan,b.Kecamatan,b.Kota,b.Propinsi,b.KodePos,b.IdTelp+b.Telp AS TELP,b.HP1,b.HP2,a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV  as usename,"
50                sSQL = sSQL & " a.HP1SV as usehp,ISNULL((select Top 1 fs_rc from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''),'') tipelastsb,"
51                sSQL = sSQL & " ISNULL((select Top 1 invdate from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''),'') tgllastsb,"
52                sSQL = sSQL & " ISNULL((select Top 1 FS_NM_DEALER from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''),'') dealerlassb,"
53                sSQL = sSQL & " ISNULL((select Top 1 FRONT from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''),'') lastsa,"
54                sSQL = sSQL & " ISNULL((select top 1 (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) from m_estimasi where NoRangka = a.NoRangka and nowo = '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) asc ),'1900-01-01') tglnextservice,"
55                sSQL = sSQL & " ISNULL((select top 1 sbenext from m_estimasi where NoRangka = a.NoRangka and nowo = '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) asc ),0) as sbnext,"
56                sSQL = sSQL & " a.NmDlrPenjual,"
57                sSQL = sSQL & " ISNULL((select Top 1 FS_NM_DEALER from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = ''),'') dealerterakhirsv,"
58                sSQL = sSQL & " a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK as namastnk,"
59                sSQL = sSQL & " a.AlamatSTNK , a.PropinsiSTNK, a.KotaSTNK, a.KecamatanSTNK, a.KodePosSTNK"
60                sSQL = sSQL & " from SingleIDOri..UnitCustomer a"
61                sSQL = sSQL & " left join SingleIDOri..NewCustomer b on b.IDCust = case when a.IDCustSales <> '' then a.IDCustSales when a.idcustservice <> '' then a.idcustservice   end"
62                sSQL = sSQL & " where NoRangka = '" & NulSID(rsNorangka!noRangka) & "'"
63                Set rsData = objO.FetchData(sSQL)
64                If Not rsData.EOF Then
65                    rsData.MoveFirst
66                    With rsDownload
67                        .AddNew
68                        If kolom(0).value = vbChecked Then
69                            !lastReminder = Format(rsData!lastReminder, "dd-mm-yyyy")
70                        End If
71                        If kolom(1).value = vbChecked Then
72                            !customer = Left(NulSID(rsData!namacustomer), 200)
73                        End If
74                        If kolom(2).value = vbChecked Then
75                            !Fmdl = Left(NulSID(rsData!namaModel), 100)
76                        End If
77                        If kolom(3).value = vbChecked Then
78                            !NoPolisi = Left(NulSID(rsData!NoPolisi), 40)
79                        End If
80                        If kolom(4).value = vbChecked Then
81                            !noRangka = Left(NulSID(rsData!noRangka), 100)
82                        End If
83                        If kolom(5).value = vbChecked Then
84                            !Tahun = Left(NulSID(rsData!thMobil), 10)
85                        End If
86                        If kolom(6).value = vbChecked Then
87                            !dDate = Format(rsData!dDate, "dd-mm-yyyy")
88                        End If
89                        If kolom(7).value = vbChecked Then
90                            !retailSales = Left(Format(NulSID(rsData!tglretail), "DD-MM-YYYY"), 50)
91                        End If
92                        If kolom(8).value = vbChecked Then
93                            !MDL = Left(NulSID(rsData!MDL), 100)
94                        End If
95                        If kolom(9).value = vbChecked Then
96                            !Alamat = Left(NulSID(rsData!Alamat), 300)
97                        End If
98                        !kelurahan = Left(NulSID(rsData!kelurahan), 100)
99                        !kecamatan = Left(NulSID(rsData!kecamatan), 100)
100                       !Kota = Left(NulSID(rsData!Kota), 100)
101                       !Propinsi = Left(NulSID(rsData!Propinsi), 100)
102                       !KodePos = Left(NulSID(rsData!KodePos), 100)
                          
103                       If kolom(10).value = vbChecked Then
104                           !PTSN = Left(NulSID(rsData!telp), 100)
105                       End If
                          
106                       !hp1 = Left(NulSID(rsData!hp1), 100)
107                       !hp2 = Left(NulSID(rsData!hp2), 100)
108                       !lastSB = Left(NulSID(rsData!tipelastsb), 100)
109                       If kolom(12).value = vbChecked Then
110                           !NEXTSB = Left(NulSID(rsData!SBNEXT), 100)
111                       End If
112                       !TGLNEXTSB = Left(NulSID(rsData!tglnextservice), 100)
113                       If kolom(13).value = vbChecked Then
114                           !DealerPenjual = Left(NulSID(rsData!NmDlrPenjual), 200)
115                       End If
116                       If kolom(15).value = vbChecked Then
117                           !WIRANIAGA = Left(NulSID(rsNorangka!kode_sales), 100)
118                       End If
119                       !UseName = Left(NulSID(rsData!UseName), 200)
120                       !USEHP = Left(NulSID(rsData!hp1), 100)
121                       !tipelastsb = Left(NulSID(rsData!tipelastsb), 100)
122                       !LASTSVCSB = Left(NulSID(rsData!tgllastsb), 100)
123                       !UMURLASTSB = "" 'Left(NulSID(rsData!namaModel), 100)
124                       If kolom(14).value = vbChecked Then
125                           !dealerLastSB = Left(NulSID(rsData!dealerlassb), 100)
126                       End If
127                       If isDownloadAll Then
128                           If kolom(16).value = vbChecked Then
129                               !CARABAYAR = Left(NulSID(rsNorangka!bayarcust), 100)
130                           End If
131                           If kolom(17).value = vbChecked Then
132                               !NAMASPK = Left(NulSID(rsNorangka!Nama_Faktur), 100)
133                           End If
134                           If kolom(18).value = vbChecked Then
135                               !ALAMATSPK = Left(NulSID(rsNorangka!Alamat1_Faktur), 100)
136                           End If
137                           !PROPINSISPK = Left(rsNorangka(rsData!Propinsi), 100)
138                           !KOTASPK = Left(NulSID(rsNorangka!Kota), 100)
139                           !KECSPK = Left(NulSID(rsNorangka!kecamatan), 100)
140                           !KODEPOSSPK = Left(NulSID(rsNorangka!Kode_Pos), 100)
141                           !CDMASPK = Left(NulSID(rsNorangka!hp2), 100)
142                           If kolom(19).value = vbChecked Then
143                               !HP1SPK = Left(NulSID(rsNorangka!hp1), 100)
144                           End If
145                           !HP2SPK = Left(NulSID(rsNorangka!hp2), 100)
146                           If kolom(20).value = vbChecked Then
147                               !namastnk = Left(NulSID(rsData!namastnk), 100)
148                           End If
149                           If kolom(21).value = vbChecked Then
150                               !AlamatSTNK = Left(NulSID(rsData!AlamatSTNK), 100)
151                           End If
152                           !PropinsiSTNK = Left(NulSID(rsData!PropinsiSTNK), 100)
153                           !Kotastnk = Left(NulSID(rsData!Kotastnk), 100)
154                           !KECSTNK = Left(NulSID(rsData!KecamatanSTNK), 100)
155                           !KodePosSTNK = Left(NulSID(rsData!KodePosSTNK), 100)
156                       End If
157                       .Update
158                   End With
159               End If
                  
160               rsNorangka.MoveNext
161               ProgVin.value = ProgVin.value + 1
162               DoEvents
163           Loop
164       End If
165       ProgVin.Visible = False
          Dim Er As New ErrInfo
166       If rsDownload.RecordCount = 0 Then
167           Er.HeadCap = "Export To Excel !"
168           Er.InfoCap = "List Data Reminder Service Kosong !"
169           Beep
170           Er.show (vbModal)
171           Exit Sub
172       End If
          Dim MyPath As String
          Dim onVal As String
173       MyPath = ""
          Dim f As New FormExplorer
174       f.show (vbModal)
175       MyPath = f.PathSaving
176       If MyPath = "" Then Exit Sub
177       If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
178       If f.FileName = "" Then Exit Sub
179       onVal = MyPath & f.FileName & ".XLS"
          
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
180       sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
181       sSQL = sSQL & " values ('" & gSA & "','" & Me.Name & "',getdate(),'" & Format(tglAwalVin.value, "MM/DD/YYYY") & "','" & Format(tglAkhirVin.value, "MM/DD/YYYY") & "','','" & rsDownload.RecordCount & "','" & onVal & "','" & ComptName & "') "
182       Call objO.InsertData(sSQL)
          
183       Call ProsesExport(onVal)
184       Er.HeadCap = "Export to excel !"
185       Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                       "File : " & onVal
186       Er.show (vbModal)
End Sub

Private Sub cmdExportExcel_Click()
    
Dim myExcel As Excel.Application
Set myExcel = New Excel.Application
myExcel.Visible = True
myExcel.Workbooks.Add

MSChart1.EditCopy

' Below is the code for pasting a metafile
myExcel.ActiveSheet.PasteSpecial Format:="Picture (Enhanced Metafile)"
       
' Below is the code for pasting a bitmap
myExcel.ActiveSheet.PasteSpecial Format:="Bitmap"
       
Set myExcel = Nothing
End Sub

Private Sub cmdProsesVin_Click()
    If Month(tglAwalVin) <> Month(tglAkhirVin) Or year(tglAwalVin) <> year(tglAkhirVin) Then
        MsgBox "Range periode dalam 1 bulan", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If opTipe(7).value = True Then
        Call ambilLaporanVinRetentionTglInvoice
    ElseIf opTipe(6).value = True Then
        Call ambilLaporanVinRetentionTglRetail
    ElseIf opTipe(8).value = True Then
        Call ambilLaporanVinRetentionTglDelivery
    End If
    
    
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
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()
    frameKolom.Visible = False
End Sub

Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Form_Activate()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    
End Sub

Private Sub Form_Load()
    Tgl1 = isnow
    Tgl2 = isnow
    
    Tgl1K = isnow
    Tgl2K = isnow
    
    Tgl1Mdl = isnow
    Tgl2Mdl = isnow
    
    
    tglAwalVin = isnow
    tglAkhirVin = isnow
    
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
     
        
    If Log.vinRetention = 1 Then
        frameVinByVinRetention.Visible = True
        ckKolom.value = vbChecked
        Call aturKolomLevel
    Else
        frameVinByVinRetention.Visible = False
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
   ' Set MSChart1.DataSource = rsChart
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
    
    cmdPrintOut.Top = CRViewer1.Top + 345
    cmdPrintOut.Left = cboPrinter.Left + cboPrinter.Width + 30
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
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdPrintOut_Click()
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

Private Sub ambilLaporanVinRetentionTglInvoice()
    Dim rsModel             As ADODB.Recordset
    Dim rsSBE               As ADODB.Recordset
    Dim rsDemand(10)        As ADODB.Recordset
    Dim rsSBETotal          As ADODB.Recordset
    Dim nilaiTemp           As Double
    Dim nilaiTempDemand(10) As Double
    Dim nilaiTempTotal(10)  As Double
    
    Set Report = New rptSBERateByModelVinRetention
    Call createTableLaporanVin
    
    ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
    sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
    Set rsModel = objO.FetchData(sSQL)
    
    If Log.RPLHistori = 0 Then
        sSQL = "select COUNT(noinv) total,fs_rc,"
        sSQL = sSQL & " E.fs_grp_mdl"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
        sSQL = sSQL & " LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC, e.fs_grp_mdl"
        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
    Else
        sSQL = "select COUNT(noinv) total,fs_rc,"
        sSQL = sSQL & " E.fs_grp_mdl"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC, e.fs_grp_mdl"
        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
    
    End If
    Set rsSBETotal = objO.FetchData(sSQL)
    
    If Log.RPLHistori = 0 Then
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
        sSQL = sSQL & " E.fs_grp_mdl"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
        sSQL = sSQL & " LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , e.fs_grp_mdl"
        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
    Else
        sSQL = "select COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
        sSQL = sSQL & " E.fs_grp_mdl"
        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
        sSQL = sSQL & " grOUP BY FS_RC , case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , e.fs_grp_mdl"
        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
    End If
    Set rsSBE = objO.FetchData(sSQL)
    If Not rsModel.EOF Then
        rsModel.MoveFirst
        ProgVin.Max = rsModel.RecordCount + 1
        ProgVin.value = 1
        ProgVin.Visible = True
        DoEvents
        
        For i = 0 To 10
            nilaiTotal1(i) = 0
            nilaiTotal2(i) = 0
            nilaiTotal3(i) = 0
            nilaiTotal4(i) = 0
            nilaiTotal5(i) = 0
            nilaiTotal6(i) = 0
            nilaiTotal7(i) = 0
        Next
        
        Do While Not rsModel.EOF
            
            ''LAPORAN SAME OUTLET
            With rsLaporan
                .AddNew
                rsSBE.filter = ""
                If Not rsSBE.EOF Then
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K1 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K10 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K20 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K30 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K40 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K50 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K60 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K70 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K80 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                    !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                    
                    !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                    !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                    
                Else
                    !K1 = 0
                    !K10 = 0
                    !K20 = 0
                    !K30 = 0
                    !K40 = 0
                    !K50 = 0
                    !K60 = 0
                    !K70 = 0
                    !K80 = 0
                    !tot16 = 0
                    !tot19 = 0
                    
                    !tot26 = 0
                    !tot29 = 0
                    
                End If
                
                nilaiTotal1(0) = nilaiTotal1(0) + !K1
                nilaiTotal1(1) = nilaiTotal1(1) + !K10
                nilaiTotal1(2) = nilaiTotal1(2) + !K20
                nilaiTotal1(3) = nilaiTotal1(3) + !K30
                nilaiTotal1(4) = nilaiTotal1(4) + !K40
                nilaiTotal1(5) = nilaiTotal1(5) + !K50
                nilaiTotal1(6) = nilaiTotal1(6) + !K60
                nilaiTotal1(7) = nilaiTotal1(7) + !K70
                nilaiTotal1(8) = nilaiTotal1(8) + !K80
                
                !judul = "SB Same Outlet"
                !model = Nul(rsModel.Fields(0).value)
                !Urut = 1
                .Update
            End With
            
            ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            ''2020-03-18 || DEMAND SAME OUTLET
            ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            With rsLaporan
                .AddNew
            
                sSQL = "select round(isnull(sum(a.total),0)/2,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K1 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/3,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K10 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/5,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K20 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/7,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K30 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/9,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K40 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/11,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K50 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/13,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K60 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/15,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K70 = objO.FetchDataSPK_Double(sSQL)
                
                sSQL = "select round(isnull(sum(a.total),0)/17,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
                !K80 = objO.FetchDataSPK_Double(sSQL)
                
                !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                
                !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                
                nilaiTotal2(0) = nilaiTotal2(0) + !K1
                nilaiTotal2(1) = nilaiTotal2(1) + !K10
                nilaiTotal2(2) = nilaiTotal2(2) + !K20
                nilaiTotal2(3) = nilaiTotal2(3) + !K30
                nilaiTotal2(4) = nilaiTotal2(4) + !K40
                nilaiTotal2(5) = nilaiTotal2(5) + !K50
                nilaiTotal2(6) = nilaiTotal2(6) + !K60
                nilaiTotal2(7) = nilaiTotal2(7) + !K70
                nilaiTotal2(8) = nilaiTotal2(8) + !K80
                 
                !judul = "Demand SB Same Outlet"
                !model = Nul(rsModel.Fields(0).value)
                !Urut = 2
                .Update
            End With
            ''------------------------------------------------------------
            '' AKHIR DEMAND
            ''------------------------------------------------------------
            
            
            ''-----------------------------------------------------------
            ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
            ''-----------------------------------------------------------
            For i = 0 To 10
                nilaiTempDemand(i) = 0
                nilaiTempTotal(i) = 0
            Next
            rsLaporan.filter = ""
            If Not rsLaporan.EOF Then
                rsLaporan.filter = "judul = 'Demand SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
                If Not rsLaporan.EOF Then
                    If rsLaporan!K1 > 0 Then
                        nilaiTempDemand(0) = rsLaporan!K1
                    End If
                    If rsLaporan!K10 > 0 Then
                        nilaiTempDemand(1) = rsLaporan!K10
                    End If
                    If rsLaporan!K20 > 0 Then
                        nilaiTempDemand(2) = rsLaporan!K20
                    End If
                    If rsLaporan!K30 > 0 Then
                        nilaiTempDemand(3) = rsLaporan!K30
                    End If
                    If rsLaporan!K40 > 0 Then
                        nilaiTempDemand(4) = rsLaporan!K40
                    End If
                    If rsLaporan!K50 > 0 Then
                        nilaiTempDemand(5) = rsLaporan!K50
                    End If
                    If rsLaporan!K60 > 0 Then
                        nilaiTempDemand(6) = rsLaporan!K60
                    End If
                    If rsLaporan!K70 > 0 Then
                        nilaiTempDemand(7) = rsLaporan!K70
                    End If
                    If rsLaporan!K80 > 0 Then
                        nilaiTempDemand(8) = rsLaporan!K80
                    End If
                End If
            End If
            rsLaporan.filter = ""
            
            If Not rsLaporan.EOF Then
                rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
                If Not rsLaporan.EOF Then
                    If rsLaporan!K1 > 0 Then
                        nilaiTempTotal(0) = rsLaporan!K1
                    End If
                    If rsLaporan!K10 > 0 Then
                        nilaiTempTotal(1) = rsLaporan!K10
                    End If
                    If rsLaporan!K20 > 0 Then
                        nilaiTempTotal(2) = rsLaporan!K20
                    End If
                    If rsLaporan!K30 > 0 Then
                        nilaiTempTotal(3) = rsLaporan!K30
                    End If
                    If rsLaporan!K40 > 0 Then
                        nilaiTempTotal(4) = rsLaporan!K40
                    End If
                    If rsLaporan!K50 > 0 Then
                        nilaiTempTotal(5) = rsLaporan!K50
                    End If
                    If rsLaporan!K60 > 0 Then
                        nilaiTempTotal(6) = rsLaporan!K60
                    End If
                    If rsLaporan!K70 > 0 Then
                        nilaiTempTotal(7) = rsLaporan!K70
                    End If
                    If rsLaporan!K80 > 0 Then
                        nilaiTempTotal(8) = rsLaporan!K80
                    End If
                End If
            End If
            
            With rsLaporan
                .AddNew
                If nilaiTempDemand(0) > 0 Then
                    !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
                Else
                    !K1 = 0 & " %"
                End If
                If nilaiTempDemand(1) > 0 Then
                    !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
                Else
                    !K10 = 0 & " %"
                End If
                If nilaiTempDemand(2) > 0 Then
                    !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
                Else
                    !K20 = 0 & " %"
                End If
                If nilaiTempDemand(3) > 0 Then
                    !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
                Else
                    !K30 = 0 & " %"
                End If
                If nilaiTempDemand(4) > 0 Then
                    !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
                Else
                    !K40 = 0 & " %"
                End If
                If nilaiTempDemand(5) > 0 Then
                    !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
                Else
                    !K50 = 0 & " %"
                End If
                If nilaiTempDemand(6) > 0 Then
                    !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
                Else
                    !K60 = 0 & " %"
                End If
                If nilaiTempDemand(7) > 0 Then
                    !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
                Else
                    !K70 = 0 & " %"
                End If
                If nilaiTempDemand(8) > 0 Then
                    !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
                Else
                    !K80 = 0 & " %"
                End If
                
                If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
                    !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
                Else
                    !tot16 = 0 & " %"
                End If
                
                If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
                    !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
                Else
                    !tot26 = 0 & " %"
                End If
                 
                If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
                    !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
                Else
                    !tot19 = 0 & " %"
                End If
                 
                If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
                    !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
                Else
                    !tot29 = 0 & " %"
                End If
                 
                !Urut = 3
                !judul = "Retention Rate Same Outlate"
                !model = Nul(rsModel.Fields(0).value)
                .Update
            End With
            
            With rsLaporan
                .AddNew
                rsSBE.filter = ""
                If Not rsSBE.EOF Then
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K1 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K10 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K20 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K30 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K40 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K50 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K60 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K70 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                    
                    rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBE.EOF Then
                        Do While Not rsSBE.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
                            rsSBE.MoveNext
                        Loop
                    End If
                    !K80 = nilaiTemp
                    rsSBE.filter = ""
                    nilaiTemp = 0
                     
                    !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                    !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                    
                    !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                    !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                    
                Else
                    !K1 = 0
                    !K10 = 0
                    !K20 = 0
                    !K30 = 0
                    !K40 = 0
                    !K50 = 0
                    !K60 = 0
                    !K70 = 0
                    !K80 = 0
                    !tot16 = 0
                    !tot19 = 0
                    
                    !tot26 = 0
                    !tot29 = 0
                End If
                
                nilaiTotal4(0) = nilaiTotal4(0) + !K1
                nilaiTotal4(1) = nilaiTotal4(1) + !K10
                nilaiTotal4(2) = nilaiTotal4(2) + !K20
                nilaiTotal4(3) = nilaiTotal4(3) + !K30
                nilaiTotal4(4) = nilaiTotal4(4) + !K40
                nilaiTotal4(5) = nilaiTotal4(5) + !K50
                nilaiTotal4(6) = nilaiTotal4(6) + !K60
                nilaiTotal4(7) = nilaiTotal4(7) + !K70
                nilaiTotal4(8) = nilaiTotal4(8) + !K80
                
                !judul = "SB Same Group (NCM)"
                !model = Nul(rsModel.Fields(0).value)
                !Urut = 4
                .Update
            End With
            
            
            For i = 0 To 10
                nilaiTempTotal(i) = 0
            Next
            rsLaporan.filter = ""
            
            If Not rsLaporan.EOF Then
                rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
                If Not rsLaporan.EOF Then
                    If rsLaporan!K1 > 0 Then
                        nilaiTempTotal(0) = rsLaporan!K1
                    End If
                    If rsLaporan!K10 > 0 Then
                        nilaiTempTotal(1) = rsLaporan!K10
                    End If
                    If rsLaporan!K20 > 0 Then
                        nilaiTempTotal(2) = rsLaporan!K20
                    End If
                    If rsLaporan!K30 > 0 Then
                        nilaiTempTotal(3) = rsLaporan!K30
                    End If
                    If rsLaporan!K40 > 0 Then
                        nilaiTempTotal(4) = rsLaporan!K40
                    End If
                    If rsLaporan!K50 > 0 Then
                        nilaiTempTotal(5) = rsLaporan!K50
                    End If
                    If rsLaporan!K60 > 0 Then
                        nilaiTempTotal(6) = rsLaporan!K60
                    End If
                    If rsLaporan!K70 > 0 Then
                        nilaiTempTotal(7) = rsLaporan!K70
                    End If
                    If rsLaporan!K80 > 0 Then
                        nilaiTempTotal(8) = rsLaporan!K80
                    End If
                End If
            End If
            
            With rsLaporan
                .AddNew
                If nilaiTempDemand(0) > 0 Then
                    !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
                Else
                    !K1 = 0 & " %"
                End If
                If nilaiTempDemand(1) > 0 Then
                    !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
                Else
                    !K10 = 0 & " %"
                End If
                If nilaiTempDemand(2) > 0 Then
                    !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
                Else
                    !K20 = 0 & " %"
                End If
                If nilaiTempDemand(3) > 0 Then
                    !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
                Else
                    !K30 = 0 & " %"
                End If
                If nilaiTempDemand(4) > 0 Then
                    !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
                Else
                    !K40 = 0 & " %"
                End If
                If nilaiTempDemand(5) > 0 Then
                    !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
                Else
                    !K50 = 0 & " %"
                End If
                If nilaiTempDemand(6) > 0 Then
                    !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
                Else
                    !K60 = 0 & " %"
                End If
                If nilaiTempDemand(7) > 0 Then
                    !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
                Else
                    !K70 = 0 & " %"
                End If
                If nilaiTempDemand(8) > 0 Then
                    !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
                Else
                    !K80 = 0 & " %"
                End If
                
                
                If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
                    !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
                Else
                    !tot16 = 0 & " %"
                End If
                
                If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
                    !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
                Else
                    !tot26 = 0 & " %"
                End If
                 
                If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
                    !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
                Else
                    !tot19 = 0 & " %"
                End If
                 
                If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
                    !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
                Else
                    !tot29 = 0 & " %"
                End If
                
                !Urut = 5
                !judul = "Retention Rate Same Group (NCM)"
                !model = Nul(rsModel.Fields(0).value)
                .Update
            End With
             
            With rsLaporan
                .AddNew
                If Not rsSBETotal.EOF Then
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    rsSBETotal.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K1 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K10 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K20 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K30 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K40 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K50 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K60 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K70 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    rsSBETotal.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
                    If Not rsSBETotal.EOF Then
                        Do While Not rsSBETotal.EOF
                            nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
                            rsSBETotal.MoveNext
                        Loop
                    End If
                    !K80 = nilaiTemp
                    rsSBETotal.filter = ""
                    nilaiTemp = 0
                    
                    !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                    !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                    
                    !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
                    !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                Else
                    !K1 = 0
                    !K10 = 0
                    !K20 = 0
                    !K30 = 0
                    !K40 = 0
                    !K50 = 0
                    !K60 = 0
                    !K70 = 0
                    !K80 = 0
                    !tot16 = 0
                    !tot19 = 0
                    !tot29 = 0
                    !tot26 = 0
                End If
                
                nilaiTotal6(0) = nilaiTotal6(0) + !K1
                nilaiTotal6(1) = nilaiTotal6(1) + !K10
                nilaiTotal6(2) = nilaiTotal6(2) + !K20
                nilaiTotal6(3) = nilaiTotal6(3) + !K30
                nilaiTotal6(4) = nilaiTotal6(4) + !K40
                nilaiTotal6(5) = nilaiTotal6(5) + !K50
                nilaiTotal6(6) = nilaiTotal6(6) + !K60
                nilaiTotal6(7) = nilaiTotal6(7) + !K70
                nilaiTotal6(8) = nilaiTotal6(8) + !K80
                
                !judul = "SB Total (NCM & Non NCM)"
                !model = Nul(rsModel.Fields(0).value)
                !Urut = 6
                .Update
            End With
            rsLaporan.filter = ""
            
            
            For i = 0 To 10
                nilaiTempTotal(i) = 0
            Next
            rsLaporan.filter = ""
            
            If Not rsLaporan.EOF Then
                rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
                If Not rsLaporan.EOF Then
                    If rsLaporan!K1 > 0 Then
                        nilaiTempTotal(0) = rsLaporan!K1
                    End If
                    If rsLaporan!K10 > 0 Then
                        nilaiTempTotal(1) = rsLaporan!K10
                    End If
                    If rsLaporan!K20 > 0 Then
                        nilaiTempTotal(2) = rsLaporan!K20
                    End If
                    If rsLaporan!K30 > 0 Then
                        nilaiTempTotal(3) = rsLaporan!K30
                    End If
                    If rsLaporan!K40 > 0 Then
                        nilaiTempTotal(4) = rsLaporan!K40
                    End If
                    If rsLaporan!K50 > 0 Then
                        nilaiTempTotal(5) = rsLaporan!K50
                    End If
                    If rsLaporan!K60 > 0 Then
                        nilaiTempTotal(6) = rsLaporan!K60
                    End If
                    If rsLaporan!K70 > 0 Then
                        nilaiTempTotal(7) = rsLaporan!K70
                    End If
                    If rsLaporan!K80 > 0 Then
                        nilaiTempTotal(8) = rsLaporan!K80
                    End If
                End If
            End If
            
            With rsLaporan
                .AddNew
                If nilaiTempDemand(0) > 0 Then
                    !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
                Else
                    !K1 = 0 & " %"
                End If
                If nilaiTempDemand(1) > 0 Then
                    !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
                Else
                    !K10 = 0 & " %"
                End If
                If nilaiTempDemand(2) > 0 Then
                    !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
                Else
                    !K20 = 0 & " %"
                End If
                If nilaiTempDemand(3) > 0 Then
                    !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
                Else
                    !K30 = 0 & " %"
                End If
                If nilaiTempDemand(4) > 0 Then
                    !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
                Else
                    !K40 = 0 & " %"
                End If
                If nilaiTempDemand(5) > 0 Then
                    !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
                Else
                    !K50 = 0 & " %"
                End If
                If nilaiTempDemand(6) > 0 Then
                    !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
                Else
                    !K60 = 0 & " %"
                End If
                If nilaiTempDemand(7) > 0 Then
                    !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
                Else
                    !K70 = 0 & " %"
                End If
                If nilaiTempDemand(8) > 0 Then
                    !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
                Else
                    !K80 = 0 & " %"
                End If
                
                
                If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
                    !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
                Else
                    !tot16 = 0 & " %"
                End If
                
                If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
                    !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
                Else
                    !tot26 = 0 & " %"
                End If
                 
                If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
                    !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
                Else
                    !tot19 = 0 & " %"
                End If
                 
                If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
                    !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
                Else
                    !tot29 = 0 & " %"
                End If
                
                !Urut = 7
                !judul = "Retention Rate Total (NCM & Non NCM)"
                !model = Nul(rsModel.Fields(0).value)
                .Update
            End With
            rsLaporan.filter = ""
            rsModel.MoveNext
            ProgVin.value = ProgVin.value + 1
            DoEvents
        Loop
    Else
        MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
        Exit Sub
    End If
    
    Call addTotalRow
    
    Report.database.SetDataSource rsLaporan
    ProgVin.Visible = False
    Report.txtHeaderLaporan.SetText "Laporan Retention " & gNmCbg
    Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
        .Left = 300
        .Height = Me.Height - 1200
        .Width = Me.Width - 750
    End With
   
    cmdTutup.Visible = True
    cmdTutup.Left = CRViewer1.Left + 20
    cmdTutup.Top = CRViewer1.Top + 400
    Screen.MousePointer = vbDefault
    
    Call DisableExport
    Call setLokasiTombol
End Sub

Private Sub ambilLaporanVinRetentionTglRetail()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
          ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchData(sSQL)
          
5         If Log.RPLHistori = 0 Then
6             sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
7             sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
8             sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
9             sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
10            sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin, "mm/dd/yyyy") & "' )"
11            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
12            sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
13            sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
14        Else
15            sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
16            sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
17            sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
18            sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
19            sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [" & Log.LogSPK & "]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin, "mm/dd/yyyy") & "' )"
20            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
21            sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
22            sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
23        End If
          
24        Set rsSBE = objO.FetchData(sSQL)
25        If Not rsModel.EOF Then
26            rsModel.MoveFirst
27            ProgVin.Max = rsModel.RecordCount + 1
28            ProgVin.value = 1
29            ProgVin.Visible = True
30            DoEvents
              
31            For i = 0 To 10
32                nilaiTotal1(i) = 0
33                nilaiTotal2(i) = 0
34                nilaiTotal3(i) = 0
35                nilaiTotal4(i) = 0
36                nilaiTotal5(i) = 0
37                nilaiTotal6(i) = 0
38                nilaiTotal7(i) = 0
39            Next
              
40            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
41                With rsLaporan
42                    .AddNew
43                    rsSBE.filter = ""
44                    If Not rsSBE.EOF Then
45                        rsSBE.filter = ""
46                        nilaiTemp = 0
47                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
48                        If Not rsSBE.EOF Then
49                            Do While Not rsSBE.EOF
50                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
51                                rsSBE.MoveNext
52                            Loop
53                        End If
54                        !K1 = nilaiTemp
55                        rsSBE.filter = ""
56                        nilaiTemp = 0
                          
57                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
58                        If Not rsSBE.EOF Then
59                            Do While Not rsSBE.EOF
60                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
61                                rsSBE.MoveNext
62                            Loop
63                        End If
64                        !K10 = nilaiTemp
65                        rsSBE.filter = ""
66                        nilaiTemp = 0
                          
67                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
68                        If Not rsSBE.EOF Then
69                            Do While Not rsSBE.EOF
70                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
71                                rsSBE.MoveNext
72                            Loop
73                        End If
74                        !K20 = nilaiTemp
75                        rsSBE.filter = ""
76                        nilaiTemp = 0
                          
77                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
78                        If Not rsSBE.EOF Then
79                            Do While Not rsSBE.EOF
80                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
81                                rsSBE.MoveNext
82                            Loop
83                        End If
84                        !K30 = nilaiTemp
85                        rsSBE.filter = ""
86                        nilaiTemp = 0
                          
87                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
88                        If Not rsSBE.EOF Then
89                            Do While Not rsSBE.EOF
90                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
91                                rsSBE.MoveNext
92                            Loop
93                        End If
94                        !K40 = nilaiTemp
95                        rsSBE.filter = ""
96                        nilaiTemp = 0
                          
97                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
98                        If Not rsSBE.EOF Then
99                            Do While Not rsSBE.EOF
100                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
101                               rsSBE.MoveNext
102                           Loop
103                       End If
104                       !K50 = nilaiTemp
105                       rsSBE.filter = ""
106                       nilaiTemp = 0
                          
107                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
108                       If Not rsSBE.EOF Then
109                           Do While Not rsSBE.EOF
110                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
111                               rsSBE.MoveNext
112                           Loop
113                       End If
114                       !K60 = nilaiTemp
115                       rsSBE.filter = ""
116                       nilaiTemp = 0
                          
117                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
118                       If Not rsSBE.EOF Then
119                           Do While Not rsSBE.EOF
120                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
121                               rsSBE.MoveNext
122                           Loop
123                       End If
124                       !K70 = nilaiTemp
125                       rsSBE.filter = ""
126                       nilaiTemp = 0
                          
127                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
128                       If Not rsSBE.EOF Then
129                           Do While Not rsSBE.EOF
130                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
131                               rsSBE.MoveNext
132                           Loop
133                       End If
134                       !K80 = nilaiTemp
135                       rsSBE.filter = ""
136                       nilaiTemp = 0
                          
137                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
138                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
139                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
140                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
141                   Else
142                       !K1 = 0
143                       !K10 = 0
144                       !K20 = 0
145                       !K30 = 0
146                       !K40 = 0
147                       !K50 = 0
148                       !K60 = 0
149                       !K70 = 0
150                       !K80 = 0
151                       !tot16 = 0
152                       !tot19 = 0
                          
153                       !tot26 = 0
154                       !tot29 = 0
                          
155                   End If
                      
156                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
157                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
158                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
159                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
160                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
161                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
162                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
163                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
164                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
165                   !judul = "SB Same Outlet"
166                   !model = Nul(rsModel.Fields(0).value)
167                   !Urut = 1
168                   .Update
169               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
170               With rsLaporan
171                   .AddNew
                  
172                   sSQL = "select round(isnull(sum(a.total),0)/2,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
173                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
174                   sSQL = "select round(isnull(sum(a.total),0)/3,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
175                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
176                   sSQL = "select round(isnull(sum(a.total),0)/4,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
177                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
178                   sSQL = "select round(isnull(sum(a.total),0)/5,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
179                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
180                   sSQL = "select round(isnull(sum(a.total),0)/6,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
181                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
182                   sSQL = "select round(isnull(sum(a.total),0)/7,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
183                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
184                   sSQL = "select round(isnull(sum(a.total),0)/8,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
185                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
186                   sSQL = "select round(isnull(sum(a.total),0)/9,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
187                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
188                   sSQL = "select round(isnull(sum(a.total),0)/10,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
189                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
190                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
191                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
192                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
193                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
194                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
195                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
196                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
197                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
198                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
199                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
200                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
201                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
202                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                      
203                   !judul = "Demand SB Same Outlet"
204                   !model = Nul(rsModel.Fields(0).value)
205                   !Urut = 2
206                   .Update
207               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
208               For i = 0 To 10
209                   nilaiTempDemand(i) = 0
210                   nilaiTempTotal(i) = 0
211               Next
212               rsLaporan.filter = ""
213               If Not rsLaporan.EOF Then
214                   rsLaporan.filter = "judul = 'Demand SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
215                   If Not rsLaporan.EOF Then
216                       If rsLaporan!K1 > 0 Then
217                           nilaiTempDemand(0) = rsLaporan!K1
218                       End If
219                       If rsLaporan!K10 > 0 Then
220                           nilaiTempDemand(1) = rsLaporan!K10
221                       End If
222                       If rsLaporan!K20 > 0 Then
223                           nilaiTempDemand(2) = rsLaporan!K20
224                       End If
225                       If rsLaporan!K30 > 0 Then
226                           nilaiTempDemand(3) = rsLaporan!K30
227                       End If
228                       If rsLaporan!K40 > 0 Then
229                           nilaiTempDemand(4) = rsLaporan!K40
230                       End If
231                       If rsLaporan!K50 > 0 Then
232                           nilaiTempDemand(5) = rsLaporan!K50
233                       End If
234                       If rsLaporan!K60 > 0 Then
235                           nilaiTempDemand(6) = rsLaporan!K60
236                       End If
237                       If rsLaporan!K70 > 0 Then
238                           nilaiTempDemand(7) = rsLaporan!K70
239                       End If
240                       If rsLaporan!K80 > 0 Then
241                           nilaiTempDemand(8) = rsLaporan!K80
242                       End If
243                   End If
244               End If
245               rsLaporan.filter = ""
                  
246               If Not rsLaporan.EOF Then
247                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
248                   If Not rsLaporan.EOF Then
249                       If rsLaporan!K1 > 0 Then
250                           nilaiTempTotal(0) = rsLaporan!K1
251                       End If
252                       If rsLaporan!K10 > 0 Then
253                           nilaiTempTotal(1) = rsLaporan!K10
254                       End If
255                       If rsLaporan!K20 > 0 Then
256                           nilaiTempTotal(2) = rsLaporan!K20
257                       End If
258                       If rsLaporan!K30 > 0 Then
259                           nilaiTempTotal(3) = rsLaporan!K30
260                       End If
261                       If rsLaporan!K40 > 0 Then
262                           nilaiTempTotal(4) = rsLaporan!K40
263                       End If
264                       If rsLaporan!K50 > 0 Then
265                           nilaiTempTotal(5) = rsLaporan!K50
266                       End If
267                       If rsLaporan!K60 > 0 Then
268                           nilaiTempTotal(6) = rsLaporan!K60
269                       End If
270                       If rsLaporan!K70 > 0 Then
271                           nilaiTempTotal(7) = rsLaporan!K70
272                       End If
273                       If rsLaporan!K80 > 0 Then
274                           nilaiTempTotal(8) = rsLaporan!K80
275                       End If
276                   End If
277               End If
                  
278               With rsLaporan
279                   .AddNew
280                   If nilaiTempDemand(0) > 0 Then
281                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
282                   Else
283                       !K1 = 0 & " %"
284                   End If
285                   If nilaiTempDemand(1) > 0 Then
286                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
287                   Else
288                       !K10 = 0 & " %"
289                   End If
290                   If nilaiTempDemand(2) > 0 Then
291                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
292                   Else
293                       !K20 = 0 & " %"
294                   End If
295                   If nilaiTempDemand(3) > 0 Then
296                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
297                   Else
298                       !K30 = 0 & " %"
299                   End If
300                   If nilaiTempDemand(4) > 0 Then
301                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
302                   Else
303                       !K40 = 0 & " %"
304                   End If
305                   If nilaiTempDemand(5) > 0 Then
306                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
307                   Else
308                       !K50 = 0 & " %"
309                   End If
310                   If nilaiTempDemand(6) > 0 Then
311                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
312                   Else
313                       !K60 = 0 & " %"
314                   End If
315                   If nilaiTempDemand(7) > 0 Then
316                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
317                   Else
318                       !K70 = 0 & " %"
319                   End If
320                   If nilaiTempDemand(8) > 0 Then
321                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
322                   Else
323                       !K80 = 0 & " %"
324                   End If
                      
325                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
326                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
327                   Else
328                       !tot16 = 0 & " %"
329                   End If
                      
330                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
331                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
332                   Else
333                       !tot26 = 0 & " %"
334                   End If
                       
335                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
336                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
337                   Else
338                       !tot19 = 0 & " %"
339                   End If
                       
340                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
341                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
342                   Else
343                       !tot29 = 0 & " %"
344                   End If
                       
345                   !Urut = 3
346                   !judul = "Retention Rate Same Outlate"
347                   !model = Nul(rsModel.Fields(0).value)
348                   .Update
349               End With
                  
350               With rsLaporan
351                   .AddNew
352                   rsSBE.filter = ""
353                   If Not rsSBE.EOF Then
354                       rsSBE.filter = ""
355                       nilaiTemp = 0
356                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
357                       If Not rsSBE.EOF Then
358                           Do While Not rsSBE.EOF
359                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
360                               rsSBE.MoveNext
361                           Loop
362                       End If
363                       !K1 = nilaiTemp
364                       rsSBE.filter = ""
365                       nilaiTemp = 0
                          
366                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
367                       If Not rsSBE.EOF Then
368                           Do While Not rsSBE.EOF
369                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
370                               rsSBE.MoveNext
371                           Loop
372                       End If
373                       !K10 = nilaiTemp
374                       rsSBE.filter = ""
375                       nilaiTemp = 0
                          
376                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
377                       If Not rsSBE.EOF Then
378                           Do While Not rsSBE.EOF
379                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
380                               rsSBE.MoveNext
381                           Loop
382                       End If
383                       !K20 = nilaiTemp
384                       rsSBE.filter = ""
385                       nilaiTemp = 0
                          
386                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
387                       If Not rsSBE.EOF Then
388                           Do While Not rsSBE.EOF
389                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
390                               rsSBE.MoveNext
391                           Loop
392                       End If
393                       !K30 = nilaiTemp
394                       rsSBE.filter = ""
395                       nilaiTemp = 0
                          
396                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
397                       If Not rsSBE.EOF Then
398                           Do While Not rsSBE.EOF
399                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
400                               rsSBE.MoveNext
401                           Loop
402                       End If
403                       !K40 = nilaiTemp
404                       rsSBE.filter = ""
405                       nilaiTemp = 0
                          
406                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
407                       If Not rsSBE.EOF Then
408                           Do While Not rsSBE.EOF
409                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
410                               rsSBE.MoveNext
411                           Loop
412                       End If
413                       !K50 = nilaiTemp
414                       rsSBE.filter = ""
415                       nilaiTemp = 0
                          
416                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
417                       If Not rsSBE.EOF Then
418                           Do While Not rsSBE.EOF
419                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
420                               rsSBE.MoveNext
421                           Loop
422                       End If
423                       !K60 = nilaiTemp
424                       rsSBE.filter = ""
425                       nilaiTemp = 0
                          
426                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
427                       If Not rsSBE.EOF Then
428                           Do While Not rsSBE.EOF
429                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
430                               rsSBE.MoveNext
431                           Loop
432                       End If
433                       !K70 = nilaiTemp
434                       rsSBE.filter = ""
435                       nilaiTemp = 0
                          
436                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
437                       If Not rsSBE.EOF Then
438                           Do While Not rsSBE.EOF
439                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
440                               rsSBE.MoveNext
441                           Loop
442                       End If
443                       !K80 = nilaiTemp
444                       rsSBE.filter = ""
445                       nilaiTemp = 0
                           
446                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
447                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
448                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
449                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
450                   Else
451                       !K1 = 0
452                       !K10 = 0
453                       !K20 = 0
454                       !K30 = 0
455                       !K40 = 0
456                       !K50 = 0
457                       !K60 = 0
458                       !K70 = 0
459                       !K80 = 0
460                       !tot16 = 0
461                       !tot19 = 0
                          
462                       !tot26 = 0
463                       !tot29 = 0
464                   End If
                      
465                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
466                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
467                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
468                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
469                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
470                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
471                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
472                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
473                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
474                   !judul = "SB Same Group (NCM)"
475                   !model = Nul(rsModel.Fields(0).value)
476                   !Urut = 4
477                   .Update
478               End With
                  
                  
479               For i = 0 To 10
480                   nilaiTempTotal(i) = 0
481               Next
482               rsLaporan.filter = ""
                  
483               If Not rsLaporan.EOF Then
484                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
485                   If Not rsLaporan.EOF Then
486                       If rsLaporan!K1 > 0 Then
487                           nilaiTempTotal(0) = rsLaporan!K1
488                       End If
489                       If rsLaporan!K10 > 0 Then
490                           nilaiTempTotal(1) = rsLaporan!K10
491                       End If
492                       If rsLaporan!K20 > 0 Then
493                           nilaiTempTotal(2) = rsLaporan!K20
494                       End If
495                       If rsLaporan!K30 > 0 Then
496                           nilaiTempTotal(3) = rsLaporan!K30
497                       End If
498                       If rsLaporan!K40 > 0 Then
499                           nilaiTempTotal(4) = rsLaporan!K40
500                       End If
501                       If rsLaporan!K50 > 0 Then
502                           nilaiTempTotal(5) = rsLaporan!K50
503                       End If
504                       If rsLaporan!K60 > 0 Then
505                           nilaiTempTotal(6) = rsLaporan!K60
506                       End If
507                       If rsLaporan!K70 > 0 Then
508                           nilaiTempTotal(7) = rsLaporan!K70
509                       End If
510                       If rsLaporan!K80 > 0 Then
511                           nilaiTempTotal(8) = rsLaporan!K80
512                       End If
513                   End If
514               End If
                  
515               With rsLaporan
516                   .AddNew
517                   If nilaiTempDemand(0) > 0 Then
518                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
519                   Else
520                       !K1 = 0 & " %"
521                   End If
522                   If nilaiTempDemand(1) > 0 Then
523                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
524                   Else
525                       !K10 = 0 & " %"
526                   End If
527                   If nilaiTempDemand(2) > 0 Then
528                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
529                   Else
530                       !K20 = 0 & " %"
531                   End If
532                   If nilaiTempDemand(3) > 0 Then
533                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
534                   Else
535                       !K30 = 0 & " %"
536                   End If
537                   If nilaiTempDemand(4) > 0 Then
538                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
539                   Else
540                       !K40 = 0 & " %"
541                   End If
542                   If nilaiTempDemand(5) > 0 Then
543                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
544                   Else
545                       !K50 = 0 & " %"
546                   End If
547                   If nilaiTempDemand(6) > 0 Then
548                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
549                   Else
550                       !K60 = 0 & " %"
551                   End If
552                   If nilaiTempDemand(7) > 0 Then
553                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
554                   Else
555                       !K70 = 0 & " %"
556                   End If
557                   If nilaiTempDemand(8) > 0 Then
558                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
559                   Else
560                       !K80 = 0 & " %"
561                   End If
                      
                      
562                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
563                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
564                   Else
565                       !tot16 = 0 & " %"
566                   End If
                      
567                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
568                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
569                   Else
570                       !tot26 = 0 & " %"
571                   End If
                       
572                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
573                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
574                   Else
575                       !tot19 = 0 & " %"
576                   End If
                       
577                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
578                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
579                   Else
580                       !tot29 = 0 & " %"
581                   End If
                      
582                   !Urut = 5
583                   !judul = "Retention Rate Same Group (NCM)"
584                   !model = Nul(rsModel.Fields(0).value)
585                   .Update
586               End With
                   
587               With rsLaporan
588                   .AddNew
589                   If Not rsSBE.EOF Then
590                       rsSBE.filter = ""
591                       nilaiTemp = 0
592                       rsSBE.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
593                       If Not rsSBE.EOF Then
594                           Do While Not rsSBE.EOF
595                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
596                               rsSBE.MoveNext
597                           Loop
598                       End If
599                       !K1 = nilaiTemp
600                       rsSBE.filter = ""
601                       nilaiTemp = 0
                          
602                       rsSBE.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
603                       If Not rsSBE.EOF Then
604                           Do While Not rsSBE.EOF
605                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
606                               rsSBE.MoveNext
607                           Loop
608                       End If
609                       !K10 = nilaiTemp
610                       rsSBE.filter = ""
611                       nilaiTemp = 0
                          
612                       rsSBE.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
613                       If Not rsSBE.EOF Then
614                           Do While Not rsSBE.EOF
615                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
616                               rsSBE.MoveNext
617                           Loop
618                       End If
619                       !K20 = nilaiTemp
620                       rsSBE.filter = ""
621                       nilaiTemp = 0
                          
622                       rsSBE.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
623                       If Not rsSBE.EOF Then
624                           Do While Not rsSBE.EOF
625                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
626                               rsSBE.MoveNext
627                           Loop
628                       End If
629                       !K30 = nilaiTemp
630                       rsSBE.filter = ""
631                       nilaiTemp = 0
                          
632                       rsSBE.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
633                       If Not rsSBE.EOF Then
634                           Do While Not rsSBE.EOF
635                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
636                               rsSBE.MoveNext
637                           Loop
638                       End If
639                       !K40 = nilaiTemp
640                       rsSBE.filter = ""
641                       nilaiTemp = 0
                          
642                       rsSBE.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
643                       If Not rsSBE.EOF Then
644                           Do While Not rsSBE.EOF
645                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
646                               rsSBE.MoveNext
647                           Loop
648                       End If
649                       !K50 = nilaiTemp
650                       rsSBE.filter = ""
651                       nilaiTemp = 0
                          
652                       rsSBE.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
653                       If Not rsSBE.EOF Then
654                           Do While Not rsSBE.EOF
655                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
656                               rsSBE.MoveNext
657                           Loop
658                       End If
659                       !K60 = nilaiTemp
660                       rsSBE.filter = ""
661                       nilaiTemp = 0
                          
662                       rsSBE.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
663                       If Not rsSBE.EOF Then
664                           Do While Not rsSBE.EOF
665                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
666                               rsSBE.MoveNext
667                           Loop
668                       End If
669                       !K70 = nilaiTemp
670                       rsSBE.filter = ""
671                       nilaiTemp = 0
                          
672                       rsSBE.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
673                       If Not rsSBE.EOF Then
674                           Do While Not rsSBE.EOF
675                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
676                               rsSBE.MoveNext
677                           Loop
678                       End If
679                       !K80 = nilaiTemp
680                       rsSBE.filter = ""
681                       nilaiTemp = 0
                          
682                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
683                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
684                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
685                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
686                   Else
687                       !K1 = 0
688                       !K10 = 0
689                       !K20 = 0
690                       !K30 = 0
691                       !K40 = 0
692                       !K50 = 0
693                       !K60 = 0
694                       !K70 = 0
695                       !K80 = 0
696                       !tot16 = 0
697                       !tot19 = 0
698                       !tot29 = 0
699                       !tot26 = 0
700                   End If
                      
701                   nilaiTotal6(0) = nilaiTotal6(0) + !K1
702                   nilaiTotal6(1) = nilaiTotal6(1) + !K10
703                   nilaiTotal6(2) = nilaiTotal6(2) + !K20
704                   nilaiTotal6(3) = nilaiTotal6(3) + !K30
705                   nilaiTotal6(4) = nilaiTotal6(4) + !K40
706                   nilaiTotal6(5) = nilaiTotal6(5) + !K50
707                   nilaiTotal6(6) = nilaiTotal6(6) + !K60
708                   nilaiTotal6(7) = nilaiTotal6(7) + !K70
709                   nilaiTotal6(8) = nilaiTotal6(8) + !K80
                      
710                   !judul = "SB Total (NCM & Non NCM)"
711                   !model = Nul(rsModel.Fields(0).value)
712                   !Urut = 6
713                   .Update
714               End With
715               rsLaporan.filter = ""
                  
                  
716               For i = 0 To 10
717                   nilaiTempTotal(i) = 0
718               Next
719               rsLaporan.filter = ""
                  
720               If Not rsLaporan.EOF Then
721                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
722                   If Not rsLaporan.EOF Then
723                       If rsLaporan!K1 > 0 Then
724                           nilaiTempTotal(0) = rsLaporan!K1
725                       End If
726                       If rsLaporan!K10 > 0 Then
727                           nilaiTempTotal(1) = rsLaporan!K10
728                       End If
729                       If rsLaporan!K20 > 0 Then
730                           nilaiTempTotal(2) = rsLaporan!K20
731                       End If
732                       If rsLaporan!K30 > 0 Then
733                           nilaiTempTotal(3) = rsLaporan!K30
734                       End If
735                       If rsLaporan!K40 > 0 Then
736                           nilaiTempTotal(4) = rsLaporan!K40
737                       End If
738                       If rsLaporan!K50 > 0 Then
739                           nilaiTempTotal(5) = rsLaporan!K50
740                       End If
741                       If rsLaporan!K60 > 0 Then
742                           nilaiTempTotal(6) = rsLaporan!K60
743                       End If
744                       If rsLaporan!K70 > 0 Then
745                           nilaiTempTotal(7) = rsLaporan!K70
746                       End If
747                       If rsLaporan!K80 > 0 Then
748                           nilaiTempTotal(8) = rsLaporan!K80
749                       End If
750                   End If
751               End If
                  
752               With rsLaporan
753                   .AddNew
754                   If nilaiTempDemand(0) > 0 Then
755                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
756                   Else
757                       !K1 = 0 & " %"
758                   End If
759                   If nilaiTempDemand(1) > 0 Then
760                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
761                   Else
762                       !K10 = 0 & " %"
763                   End If
764                   If nilaiTempDemand(2) > 0 Then
765                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
766                   Else
767                       !K20 = 0 & " %"
768                   End If
769                   If nilaiTempDemand(3) > 0 Then
770                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
771                   Else
772                       !K30 = 0 & " %"
773                   End If
774                   If nilaiTempDemand(4) > 0 Then
775                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
776                   Else
777                       !K40 = 0 & " %"
778                   End If
779                   If nilaiTempDemand(5) > 0 Then
780                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
781                   Else
782                       !K50 = 0 & " %"
783                   End If
784                   If nilaiTempDemand(6) > 0 Then
785                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
786                   Else
787                       !K60 = 0 & " %"
788                   End If
789                   If nilaiTempDemand(7) > 0 Then
790                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
791                   Else
792                       !K70 = 0 & " %"
793                   End If
794                   If nilaiTempDemand(8) > 0 Then
795                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
796                   Else
797                       !K80 = 0 & " %"
798                   End If
                      
                      
799                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
800                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
801                   Else
802                       !tot16 = 0 & " %"
803                   End If
                      
804                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
805                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
806                   Else
807                       !tot26 = 0 & " %"
808                   End If
                       
809                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
810                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
811                   Else
812                       !tot19 = 0 & " %"
813                   End If
                       
814                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
815                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
816                   Else
817                       !tot29 = 0 & " %"
818                   End If
                      
819                   !Urut = 7
820                   !judul = "Retention Rate Total (NCM & Non NCM)"
821                   !model = Nul(rsModel.Fields(0).value)
822                   .Update
823               End With
824               rsLaporan.filter = ""
825               rsModel.MoveNext
826               ProgVin.value = ProgVin.value + 1
827               DoEvents
828           Loop
829       Else
830           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
831           Exit Sub
832       End If
          
833       Call addTotalRow
          
834       Report.database.SetDataSource rsLaporan
835       ProgVin.Visible = False
836       Report.txtHeaderLaporan.SetText "Laporan Retention " & gNmCbg
837       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
838       With CRViewer1
839           .ReportSource = Report
840           .ViewReport
841           .Zoom 100
842           .Visible = True
843           .Left = 300
844           .Height = Me.Height - 1200
845           .Width = Me.Width - 750
846       End With
         
847       cmdTutup.Visible = True
848       cmdTutup.Left = CRViewer1.Left + 20
849       cmdTutup.Top = CRViewer1.Top + 400
850       Screen.MousePointer = vbDefault
          
851       Call DisableExport
852       Call setLokasiTombol
End Sub

Private Sub ambilLaporanVinRetentionTglDelivery()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
          ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchData(sSQL)
          
5         If Log.RPLHistori = 0 Then
6             sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
7             sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
8             sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
9             sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
10            sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' and '" & Format(tglAkhirVin, "mm/dd/yyyy") & "')"
11            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
12            sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
13            sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
14        Else
15            sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
16            sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..summary s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
17            sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
18            sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
19            sSQL = sSQL & " and T.NO_CHASSIS in (select NoRangka from [PDS" & Right(gDb, 2) & "]..Intrn where Proses = 10 and Kode = 2 and TglSelesai between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' and '" & Format(tglAkhirVin, "mm/dd/yyyy") & "')"
20            sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
21            sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when C.FS_KD_DEALER = '" & Log.KodeDealer & "' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
22            sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
23        End If
          
24        Set rsSBE = objO.FetchData(sSQL)
25        If Not rsModel.EOF Then
26            rsModel.MoveFirst
27            ProgVin.Max = rsModel.RecordCount + 1
28            ProgVin.value = 1
29            ProgVin.Visible = True
30            DoEvents
              
31            For i = 0 To 10
32                nilaiTotal1(i) = 0
33                nilaiTotal2(i) = 0
34                nilaiTotal3(i) = 0
35                nilaiTotal4(i) = 0
36                nilaiTotal5(i) = 0
37                nilaiTotal6(i) = 0
38                nilaiTotal7(i) = 0
39            Next
40            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
41                With rsLaporan
42                    .AddNew
43                    rsSBE.filter = ""
44                    If Not rsSBE.EOF Then
45                        rsSBE.filter = ""
46                        nilaiTemp = 0
47                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
48                        If Not rsSBE.EOF Then
49                            Do While Not rsSBE.EOF
50                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
51                                rsSBE.MoveNext
52                            Loop
53                        End If
54                        !K1 = nilaiTemp
55                        rsSBE.filter = ""
56                        nilaiTemp = 0
                          
57                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
58                        If Not rsSBE.EOF Then
59                            Do While Not rsSBE.EOF
60                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
61                                rsSBE.MoveNext
62                            Loop
63                        End If
64                        !K10 = nilaiTemp
65                        rsSBE.filter = ""
66                        nilaiTemp = 0
                          
67                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
68                        If Not rsSBE.EOF Then
69                            Do While Not rsSBE.EOF
70                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
71                                rsSBE.MoveNext
72                            Loop
73                        End If
74                        !K20 = nilaiTemp
75                        rsSBE.filter = ""
76                        nilaiTemp = 0
                          
77                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
78                        If Not rsSBE.EOF Then
79                            Do While Not rsSBE.EOF
80                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
81                                rsSBE.MoveNext
82                            Loop
83                        End If
84                        !K30 = nilaiTemp
85                        rsSBE.filter = ""
86                        nilaiTemp = 0
                          
87                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
88                        If Not rsSBE.EOF Then
89                            Do While Not rsSBE.EOF
90                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
91                                rsSBE.MoveNext
92                            Loop
93                        End If
94                        !K40 = nilaiTemp
95                        rsSBE.filter = ""
96                        nilaiTemp = 0
                          
97                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
98                        If Not rsSBE.EOF Then
99                            Do While Not rsSBE.EOF
100                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
101                               rsSBE.MoveNext
102                           Loop
103                       End If
104                       !K50 = nilaiTemp
105                       rsSBE.filter = ""
106                       nilaiTemp = 0
                          
107                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
108                       If Not rsSBE.EOF Then
109                           Do While Not rsSBE.EOF
110                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
111                               rsSBE.MoveNext
112                           Loop
113                       End If
114                       !K60 = nilaiTemp
115                       rsSBE.filter = ""
116                       nilaiTemp = 0
                          
117                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
118                       If Not rsSBE.EOF Then
119                           Do While Not rsSBE.EOF
120                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
121                               rsSBE.MoveNext
122                           Loop
123                       End If
124                       !K70 = nilaiTemp
125                       rsSBE.filter = ""
126                       nilaiTemp = 0
                          
127                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
128                       If Not rsSBE.EOF Then
129                           Do While Not rsSBE.EOF
130                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
131                               rsSBE.MoveNext
132                           Loop
133                       End If
134                       !K80 = nilaiTemp
135                       rsSBE.filter = ""
136                       nilaiTemp = 0
                          
137                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
138                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
139                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
140                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
141                   Else
142                       !K1 = 0
143                       !K10 = 0
144                       !K20 = 0
145                       !K30 = 0
146                       !K40 = 0
147                       !K50 = 0
148                       !K60 = 0
149                       !K70 = 0
150                       !K80 = 0
151                       !tot16 = 0
152                       !tot19 = 0
                          
153                       !tot26 = 0
154                       !tot29 = 0
                          
155                   End If
                      
156                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
157                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
158                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
159                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
160                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
161                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
162                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
163                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
164                   nilaiTotal1(8) = nilaiTotal1(8) + !K80

165                   !judul = "SB Same Outlet"
166                   !model = Nul(rsModel.Fields(0).value)
167                   !Urut = 1
168                   .Update
169               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
170               With rsLaporan
171                   .AddNew
                  
172                   sSQL = "select round(isnull(sum(a.total),0)/2,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
173                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
174                   sSQL = "select round(isnull(sum(a.total),0)/3,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
175                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
176                   sSQL = "select round(isnull(sum(a.total),0)/4,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
177                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
178                   sSQL = "select round(isnull(sum(a.total),0)/5,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
179                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
180                   sSQL = "select round(isnull(sum(a.total),0)/6,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
181                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
182                   sSQL = "select round(isnull(sum(a.total),0)/7,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
183                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
184                   sSQL = "select round(isnull(sum(a.total),0)/8,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
185                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
186                   sSQL = "select round(isnull(sum(a.total),0)/9,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
187                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
188                   sSQL = "select round(isnull(sum(a.total),0)/10,0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type from Sihdr a INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 where a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
189                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
190                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
191                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
192                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
193                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
194                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
195                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
196                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
197                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
198                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
199                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
200                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
201                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
202                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
203                   !judul = "Demand SB Same Outlet"
204                   !model = Nul(rsModel.Fields(0).value)
205                   !Urut = 2
206                   .Update
207               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
208               For i = 0 To 10
209                   nilaiTempDemand(i) = 0
210                   nilaiTempTotal(i) = 0
211               Next
212               rsLaporan.filter = ""
213               If Not rsLaporan.EOF Then
214                   rsLaporan.filter = "judul = 'Demand SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
215                   If Not rsLaporan.EOF Then
216                       If rsLaporan!K1 > 0 Then
217                           nilaiTempDemand(0) = rsLaporan!K1
218                       End If
219                       If rsLaporan!K10 > 0 Then
220                           nilaiTempDemand(1) = rsLaporan!K10
221                       End If
222                       If rsLaporan!K20 > 0 Then
223                           nilaiTempDemand(2) = rsLaporan!K20
224                       End If
225                       If rsLaporan!K30 > 0 Then
226                           nilaiTempDemand(3) = rsLaporan!K30
227                       End If
228                       If rsLaporan!K40 > 0 Then
229                           nilaiTempDemand(4) = rsLaporan!K40
230                       End If
231                       If rsLaporan!K50 > 0 Then
232                           nilaiTempDemand(5) = rsLaporan!K50
233                       End If
234                       If rsLaporan!K60 > 0 Then
235                           nilaiTempDemand(6) = rsLaporan!K60
236                       End If
237                       If rsLaporan!K70 > 0 Then
238                           nilaiTempDemand(7) = rsLaporan!K70
239                       End If
240                       If rsLaporan!K80 > 0 Then
241                           nilaiTempDemand(8) = rsLaporan!K80
242                       End If
243                   End If
244               End If
245               rsLaporan.filter = ""
                  
246               If Not rsLaporan.EOF Then
247                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
248                   If Not rsLaporan.EOF Then
249                       If rsLaporan!K1 > 0 Then
250                           nilaiTempTotal(0) = rsLaporan!K1
251                       End If
252                       If rsLaporan!K10 > 0 Then
253                           nilaiTempTotal(1) = rsLaporan!K10
254                       End If
255                       If rsLaporan!K20 > 0 Then
256                           nilaiTempTotal(2) = rsLaporan!K20
257                       End If
258                       If rsLaporan!K30 > 0 Then
259                           nilaiTempTotal(3) = rsLaporan!K30
260                       End If
261                       If rsLaporan!K40 > 0 Then
262                           nilaiTempTotal(4) = rsLaporan!K40
263                       End If
264                       If rsLaporan!K50 > 0 Then
265                           nilaiTempTotal(5) = rsLaporan!K50
266                       End If
267                       If rsLaporan!K60 > 0 Then
268                           nilaiTempTotal(6) = rsLaporan!K60
269                       End If
270                       If rsLaporan!K70 > 0 Then
271                           nilaiTempTotal(7) = rsLaporan!K70
272                       End If
273                       If rsLaporan!K80 > 0 Then
274                           nilaiTempTotal(8) = rsLaporan!K80
275                       End If
276                   End If
277               End If
                  
278               With rsLaporan
279                   .AddNew
280                   If nilaiTempDemand(0) > 0 Then
281                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
282                   Else
283                       !K1 = 0 & " %"
284                   End If
285                   If nilaiTempDemand(1) > 0 Then
286                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
287                   Else
288                       !K10 = 0 & " %"
289                   End If
290                   If nilaiTempDemand(2) > 0 Then
291                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
292                   Else
293                       !K20 = 0 & " %"
294                   End If
295                   If nilaiTempDemand(3) > 0 Then
296                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
297                   Else
298                       !K30 = 0 & " %"
299                   End If
300                   If nilaiTempDemand(4) > 0 Then
301                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
302                   Else
303                       !K40 = 0 & " %"
304                   End If
305                   If nilaiTempDemand(5) > 0 Then
306                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
307                   Else
308                       !K50 = 0 & " %"
309                   End If
310                   If nilaiTempDemand(6) > 0 Then
311                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
312                   Else
313                       !K60 = 0 & " %"
314                   End If
315                   If nilaiTempDemand(7) > 0 Then
316                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
317                   Else
318                       !K70 = 0 & " %"
319                   End If
320                   If nilaiTempDemand(8) > 0 Then
321                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
322                   Else
323                       !K80 = 0 & " %"
324                   End If
                      
325                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
326                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
327                   Else
328                       !tot16 = 0 & " %"
329                   End If
                      
330                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
331                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
332                   Else
333                       !tot26 = 0 & " %"
334                   End If
                       
335                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
336                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
337                   Else
338                       !tot19 = 0 & " %"
339                   End If
                       
340                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
341                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
342                   Else
343                       !tot29 = 0 & " %"
344                   End If
                       
345                   !Urut = 3
346                   !judul = "Retention Rate Same Outlate"
347                   !model = Nul(rsModel.Fields(0).value)
348                   .Update
349               End With
                  
350               With rsLaporan
351                   .AddNew
352                   rsSBE.filter = ""
353                   If Not rsSBE.EOF Then
354                       rsSBE.filter = ""
355                       nilaiTemp = 0
356                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
357                       If Not rsSBE.EOF Then
358                           Do While Not rsSBE.EOF
359                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
360                               rsSBE.MoveNext
361                           Loop
362                       End If
363                       !K1 = nilaiTemp
364                       rsSBE.filter = ""
365                       nilaiTemp = 0
                          
366                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
367                       If Not rsSBE.EOF Then
368                           Do While Not rsSBE.EOF
369                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
370                               rsSBE.MoveNext
371                           Loop
372                       End If
373                       !K10 = nilaiTemp
374                       rsSBE.filter = ""
375                       nilaiTemp = 0
                          
376                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
377                       If Not rsSBE.EOF Then
378                           Do While Not rsSBE.EOF
379                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
380                               rsSBE.MoveNext
381                           Loop
382                       End If
383                       !K20 = nilaiTemp
384                       rsSBE.filter = ""
385                       nilaiTemp = 0
                          
386                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
387                       If Not rsSBE.EOF Then
388                           Do While Not rsSBE.EOF
389                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
390                               rsSBE.MoveNext
391                           Loop
392                       End If
393                       !K30 = nilaiTemp
394                       rsSBE.filter = ""
395                       nilaiTemp = 0
                          
396                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
397                       If Not rsSBE.EOF Then
398                           Do While Not rsSBE.EOF
399                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
400                               rsSBE.MoveNext
401                           Loop
402                       End If
403                       !K40 = nilaiTemp
404                       rsSBE.filter = ""
405                       nilaiTemp = 0
                          
406                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
407                       If Not rsSBE.EOF Then
408                           Do While Not rsSBE.EOF
409                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
410                               rsSBE.MoveNext
411                           Loop
412                       End If
413                       !K50 = nilaiTemp
414                       rsSBE.filter = ""
415                       nilaiTemp = 0
                          
416                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
417                       If Not rsSBE.EOF Then
418                           Do While Not rsSBE.EOF
419                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
420                               rsSBE.MoveNext
421                           Loop
422                       End If
423                       !K60 = nilaiTemp
424                       rsSBE.filter = ""
425                       nilaiTemp = 0
                          
426                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
427                       If Not rsSBE.EOF Then
428                           Do While Not rsSBE.EOF
429                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
430                               rsSBE.MoveNext
431                           Loop
432                       End If
433                       !K70 = nilaiTemp
434                       rsSBE.filter = ""
435                       nilaiTemp = 0
                          
436                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
437                       If Not rsSBE.EOF Then
438                           Do While Not rsSBE.EOF
439                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
440                               rsSBE.MoveNext
441                           Loop
442                       End If
443                       !K80 = nilaiTemp
444                       rsSBE.filter = ""
445                       nilaiTemp = 0
                           
446                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
447                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
448                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
449                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
450                   Else
451                       !K1 = 0
452                       !K10 = 0
453                       !K20 = 0
454                       !K30 = 0
455                       !K40 = 0
456                       !K50 = 0
457                       !K60 = 0
458                       !K70 = 0
459                       !K80 = 0
460                       !tot16 = 0
461                       !tot19 = 0
                          
462                       !tot26 = 0
463                       !tot29 = 0
464                   End If
                      
465                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
466                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
467                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
468                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
469                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
470                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
471                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
472                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
473                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
474                   !judul = "SB Same Group (NCM)"
475                   !model = Nul(rsModel.Fields(0).value)
476                   !Urut = 4
477                   .Update
478               End With
                  
                  
479               For i = 0 To 10
480                   nilaiTempTotal(i) = 0
481               Next
482               rsLaporan.filter = ""
                  
483               If Not rsLaporan.EOF Then
484                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
485                   If Not rsLaporan.EOF Then
486                       If rsLaporan!K1 > 0 Then
487                           nilaiTempTotal(0) = rsLaporan!K1
488                       End If
489                       If rsLaporan!K10 > 0 Then
490                           nilaiTempTotal(1) = rsLaporan!K10
491                       End If
492                       If rsLaporan!K20 > 0 Then
493                           nilaiTempTotal(2) = rsLaporan!K20
494                       End If
495                       If rsLaporan!K30 > 0 Then
496                           nilaiTempTotal(3) = rsLaporan!K30
497                       End If
498                       If rsLaporan!K40 > 0 Then
499                           nilaiTempTotal(4) = rsLaporan!K40
500                       End If
501                       If rsLaporan!K50 > 0 Then
502                           nilaiTempTotal(5) = rsLaporan!K50
503                       End If
504                       If rsLaporan!K60 > 0 Then
505                           nilaiTempTotal(6) = rsLaporan!K60
506                       End If
507                       If rsLaporan!K70 > 0 Then
508                           nilaiTempTotal(7) = rsLaporan!K70
509                       End If
510                       If rsLaporan!K80 > 0 Then
511                           nilaiTempTotal(8) = rsLaporan!K80
512                       End If
513                   End If
514               End If
                  
515               With rsLaporan
516                   .AddNew
517                   If nilaiTempDemand(0) > 0 Then
518                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
519                   Else
520                       !K1 = 0 & " %"
521                   End If
522                   If nilaiTempDemand(1) > 0 Then
523                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
524                   Else
525                       !K10 = 0 & " %"
526                   End If
527                   If nilaiTempDemand(2) > 0 Then
528                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
529                   Else
530                       !K20 = 0 & " %"
531                   End If
532                   If nilaiTempDemand(3) > 0 Then
533                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
534                   Else
535                       !K30 = 0 & " %"
536                   End If
537                   If nilaiTempDemand(4) > 0 Then
538                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
539                   Else
540                       !K40 = 0 & " %"
541                   End If
542                   If nilaiTempDemand(5) > 0 Then
543                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
544                   Else
545                       !K50 = 0 & " %"
546                   End If
547                   If nilaiTempDemand(6) > 0 Then
548                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
549                   Else
550                       !K60 = 0 & " %"
551                   End If
552                   If nilaiTempDemand(7) > 0 Then
553                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
554                   Else
555                       !K70 = 0 & " %"
556                   End If
557                   If nilaiTempDemand(8) > 0 Then
558                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
559                   Else
560                       !K80 = 0 & " %"
561                   End If
                      
                      
562                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
563                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
564                   Else
565                       !tot16 = 0 & " %"
566                   End If
                      
567                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
568                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
569                   Else
570                       !tot26 = 0 & " %"
571                   End If
                       
572                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
573                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
574                   Else
575                       !tot19 = 0 & " %"
576                   End If
                       
577                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
578                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
579                   Else
580                       !tot29 = 0 & " %"
581                   End If
                      
582                   !Urut = 5
583                   !judul = "Retention Rate Same Group (NCM)"
584                   !model = Nul(rsModel.Fields(0).value)
585                   .Update
586               End With
                   
587               With rsLaporan
588                   .AddNew
589                   If Not rsSBE.EOF Then
590                       rsSBE.filter = ""
591                       nilaiTemp = 0
592                       rsSBE.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
593                       If Not rsSBE.EOF Then
594                           Do While Not rsSBE.EOF
595                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
596                               rsSBE.MoveNext
597                           Loop
598                       End If
599                       !K1 = nilaiTemp
600                       rsSBE.filter = ""
601                       nilaiTemp = 0
                          
602                       rsSBE.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
603                       If Not rsSBE.EOF Then
604                           Do While Not rsSBE.EOF
605                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
606                               rsSBE.MoveNext
607                           Loop
608                       End If
609                       !K10 = nilaiTemp
610                       rsSBE.filter = ""
611                       nilaiTemp = 0
                          
612                       rsSBE.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
613                       If Not rsSBE.EOF Then
614                           Do While Not rsSBE.EOF
615                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
616                               rsSBE.MoveNext
617                           Loop
618                       End If
619                       !K20 = nilaiTemp
620                       rsSBE.filter = ""
621                       nilaiTemp = 0
                          
622                       rsSBE.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
623                       If Not rsSBE.EOF Then
624                           Do While Not rsSBE.EOF
625                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
626                               rsSBE.MoveNext
627                           Loop
628                       End If
629                       !K30 = nilaiTemp
630                       rsSBE.filter = ""
631                       nilaiTemp = 0
                          
632                       rsSBE.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
633                       If Not rsSBE.EOF Then
634                           Do While Not rsSBE.EOF
635                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
636                               rsSBE.MoveNext
637                           Loop
638                       End If
639                       !K40 = nilaiTemp
640                       rsSBE.filter = ""
641                       nilaiTemp = 0
                          
642                       rsSBE.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
643                       If Not rsSBE.EOF Then
644                           Do While Not rsSBE.EOF
645                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
646                               rsSBE.MoveNext
647                           Loop
648                       End If
649                       !K50 = nilaiTemp
650                       rsSBE.filter = ""
651                       nilaiTemp = 0
                          
652                       rsSBE.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
653                       If Not rsSBE.EOF Then
654                           Do While Not rsSBE.EOF
655                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
656                               rsSBE.MoveNext
657                           Loop
658                       End If
659                       !K60 = nilaiTemp
660                       rsSBE.filter = ""
661                       nilaiTemp = 0
                          
662                       rsSBE.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
663                       If Not rsSBE.EOF Then
664                           Do While Not rsSBE.EOF
665                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
666                               rsSBE.MoveNext
667                           Loop
668                       End If
669                       !K70 = nilaiTemp
670                       rsSBE.filter = ""
671                       nilaiTemp = 0
                          
672                       rsSBE.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
673                       If Not rsSBE.EOF Then
674                           Do While Not rsSBE.EOF
675                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
676                               rsSBE.MoveNext
677                           Loop
678                       End If
679                       !K80 = nilaiTemp
680                       rsSBE.filter = ""
681                       nilaiTemp = 0
                          
682                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
683                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
684                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
685                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
686                   Else
687                       !K1 = 0
688                       !K10 = 0
689                       !K20 = 0
690                       !K30 = 0
691                       !K40 = 0
692                       !K50 = 0
693                       !K60 = 0
694                       !K70 = 0
695                       !K80 = 0
696                       !tot16 = 0
697                       !tot19 = 0
698                       !tot29 = 0
699                       !tot26 = 0
700                   End If
                      
                      
701                   nilaiTotal6(0) = nilaiTotal6(0) + !K1
702                   nilaiTotal6(1) = nilaiTotal6(1) + !K10
703                   nilaiTotal6(2) = nilaiTotal6(2) + !K20
704                   nilaiTotal6(3) = nilaiTotal6(3) + !K30
705                   nilaiTotal6(4) = nilaiTotal6(4) + !K40
706                   nilaiTotal6(5) = nilaiTotal6(5) + !K50
707                   nilaiTotal6(6) = nilaiTotal6(6) + !K60
708                   nilaiTotal6(7) = nilaiTotal6(7) + !K70
709                   nilaiTotal6(8) = nilaiTotal6(8) + !K80
                      
710                   !judul = "SB Total (NCM & Non NCM)"
711                   !model = Nul(rsModel.Fields(0).value)
712                   !Urut = 6
713                   .Update
714               End With
715               rsLaporan.filter = ""
                  
                  
716               For i = 0 To 10
717                   nilaiTempTotal(i) = 0
718               Next
719               rsLaporan.filter = ""
                  
720               If Not rsLaporan.EOF Then
721                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
722                   If Not rsLaporan.EOF Then
723                       If rsLaporan!K1 > 0 Then
724                           nilaiTempTotal(0) = rsLaporan!K1
725                       End If
726                       If rsLaporan!K10 > 0 Then
727                           nilaiTempTotal(1) = rsLaporan!K10
728                       End If
729                       If rsLaporan!K20 > 0 Then
730                           nilaiTempTotal(2) = rsLaporan!K20
731                       End If
732                       If rsLaporan!K30 > 0 Then
733                           nilaiTempTotal(3) = rsLaporan!K30
734                       End If
735                       If rsLaporan!K40 > 0 Then
736                           nilaiTempTotal(4) = rsLaporan!K40
737                       End If
738                       If rsLaporan!K50 > 0 Then
739                           nilaiTempTotal(5) = rsLaporan!K50
740                       End If
741                       If rsLaporan!K60 > 0 Then
742                           nilaiTempTotal(6) = rsLaporan!K60
743                       End If
744                       If rsLaporan!K70 > 0 Then
745                           nilaiTempTotal(7) = rsLaporan!K70
746                       End If
747                       If rsLaporan!K80 > 0 Then
748                           nilaiTempTotal(8) = rsLaporan!K80
749                       End If
750                   End If
751               End If
                  
752               With rsLaporan
753                   .AddNew
754                   If nilaiTempDemand(0) > 0 Then
755                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
756                   Else
757                       !K1 = 0 & " %"
758                   End If
759                   If nilaiTempDemand(1) > 0 Then
760                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
761                   Else
762                       !K10 = 0 & " %"
763                   End If
764                   If nilaiTempDemand(2) > 0 Then
765                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
766                   Else
767                       !K20 = 0 & " %"
768                   End If
769                   If nilaiTempDemand(3) > 0 Then
770                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
771                   Else
772                       !K30 = 0 & " %"
773                   End If
774                   If nilaiTempDemand(4) > 0 Then
775                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
776                   Else
777                       !K40 = 0 & " %"
778                   End If
779                   If nilaiTempDemand(5) > 0 Then
780                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
781                   Else
782                       !K50 = 0 & " %"
783                   End If
784                   If nilaiTempDemand(6) > 0 Then
785                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
786                   Else
787                       !K60 = 0 & " %"
788                   End If
789                   If nilaiTempDemand(7) > 0 Then
790                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
791                   Else
792                       !K70 = 0 & " %"
793                   End If
794                   If nilaiTempDemand(8) > 0 Then
795                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
796                   Else
797                       !K80 = 0 & " %"
798                   End If
                      
                      
799                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
800                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
801                   Else
802                       !tot16 = 0 & " %"
803                   End If
                      
804                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
805                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
806                   Else
807                       !tot26 = 0 & " %"
808                   End If
                       
809                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
810                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
811                   Else
812                       !tot19 = 0 & " %"
813                   End If
                       
814                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
815                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
816                   Else
817                       !tot29 = 0 & " %"
818                   End If
                      
819                   !Urut = 7
820                   !judul = "Retention Rate Total (NCM & Non NCM)"
821                   !model = Nul(rsModel.Fields(0).value)
822                   .Update
823               End With
824               rsLaporan.filter = ""
825               rsModel.MoveNext
826               ProgVin.value = ProgVin.value + 1
827               DoEvents
828           Loop
829       Else
830           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
831           Exit Sub
832       End If
          
833       Call addTotalRow
          
834       Report.database.SetDataSource rsLaporan
835       ProgVin.Visible = False
836       Report.txtHeaderLaporan.SetText "Laporan Retention " & gNmCbg
837       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
838       With CRViewer1
839           .ReportSource = Report
840           .ViewReport
841           .Zoom 100
842           .Visible = True
843           .Left = 300
844           .Height = Me.Height - 1200
845           .Width = Me.Width - 750
846       End With
         
847       cmdTutup.Visible = True
848       cmdTutup.Left = CRViewer1.Left + 20
849       cmdTutup.Top = CRViewer1.Top + 400
850       Screen.MousePointer = vbDefault
          
851       Call DisableExport
852       Call setLokasiTombol
End Sub

Private Sub ProsesExport(onV As String)
          Dim oApp    As Excel.Application
          Dim oWB     As Excel.Workbook
          Dim nAsc    As Integer
1         Set oApp = New Excel.Application
2         oApp.Visible = False
3         Set oWB = oApp.Workbooks.Add
4         With oWB
              '---------------------------
              'SA FINISH
              '---------------------------
5             nAsc = 64
6             For C = 0 To rsDownload.Fields.Count - 1
7                 .Sheets("Sheet1").Cells(1, C + 1).value = rsDownload.Fields(C).Name
8                 nAsc = nAsc + 1
9             Next
10            .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
11            .Sheets("Sheet1").Range("1:1").Font.Bold = True
              
              
12            .Sheets("Sheet1").Select
13            .Sheets("Sheet1").Name = "Download Data Non Retention"
14            .Sheets("Download Data Non Retention").Select
      '        .Sheets.Add
      '
      '        '---------------------------
      '        'SA ENTRY
      '        '---------------------------
      '        nAsc = 64
      '        For C = 0 To rsB.Fields.Count - 1
      '            .Sheets("Sheet2").Cells(1, C + 1).value = rsB.Fields(C).Name
      '            nAsc = nAsc + 1
      '        Next
      '        .Sheets("Sheet2").Cells(2, 1).CopyFromRecordset rsB
      '        .Sheets("Sheet2").Range("1:1").Font.Bold = True
      '
      '        .Charts.Add
      '        .ActiveChart.ChartType = xlLineMarkers
      '        .ActiveChart.SetSourceData .Sheets("Sheet2").Range("A1:" & Chr(nAsc) & rsB.RecordCount + 1), PlotBy:=xlColumns
      '        .ActiveChart.Location where:=xlLocationAsObject, Name:="Sheet2"
      '        With .ActiveChart
      '            .HasTitle = True
      '            .ChartTitle.Characters.Text = "KPI UNIT COVER - UNIT FINISH"
      '            .Axes(xlCategory, xlPrimary).HasTitle = True
      '            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Unit Cover"
      '            .Axes(xlValue, xlPrimary).HasTitle = True
      '            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Jumlah"
      '        End With
      '        .Sheets("Sheet2").Select
      '        .Sheets("Sheet2").Name = "Cover-Entry-Bulan"
      '        .Sheets("Cover-Entry-Bulan").Select
      '        .Sheets.Add
      '
      '        '---------------------------
      '        'SA ENTRY TOTAL
      '        '---------------------------
      '        nAsc = 64
      '        For C = 0 To rsUE.Fields.Count - 1
      '            .Sheets("Sheet3").Cells(1, C + 1).value = rsUE.Fields(C).Name
      '            nAsc = nAsc + 1
      '        Next
      '        .Sheets("Sheet3").Cells(2, 1).CopyFromRecordset rsUE
      '        .Sheets("Sheet3").Range("1:1").Font.Bold = True
      '
      '        .Charts.Add
      '        .ActiveChart.ChartType = xlColumnClustered
      '        .ActiveChart.SetSourceData .Sheets("Sheet3").Range("A1:" & Chr(nAsc) & rsUE.RecordCount + 1), PlotBy:=xlColumns
      '        .ActiveChart.Location where:=xlLocationAsObject, Name:="Sheet3"
      '        With .ActiveChart
      '            .HasTitle = True
      '            .ChartTitle.Characters.Text = "KPI UNIT COVER - UNIT ENTRY (TOTAL)"
      '            .Axes(xlCategory, xlPrimary).HasTitle = True
      '            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Unit Cover"
      '            .Axes(xlValue, xlPrimary).HasTitle = True
      '            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Jumlah"
      '        End With
      '        .Sheets("Sheet3").Select
      '        .Sheets("Sheet3").Name = "Cover-Entry-Total"
      '        .Sheets("Cover-Entry-Total").Select
      '        .Sheets.Add
      '
      '        '---------------------------
      '        'SA FINISH TOTAL
      '        '---------------------------
      '        nAsc = 64
      '        For C = 0 To rsUF.Fields.Count - 1
      '            .Sheets("Sheet4").Cells(1, C + 1).value = rsUF.Fields(C).Name
      '            nAsc = nAsc + 1
      '        Next
      '        .Sheets("Sheet4").Cells(2, 1).CopyFromRecordset rsUF
      '        .Sheets("Sheet4").Range("1:1").Font.Bold = True
      '
      '        .Charts.Add
      '        .ActiveChart.ChartType = xlColumnClustered
      '        .ActiveChart.SetSourceData .Sheets("Sheet4").Range("A1:" & Chr(nAsc) & rsUF.RecordCount + 1), PlotBy:=xlColumns
      '        .ActiveChart.Location where:=xlLocationAsObject, Name:="Sheet4"
      '        With .ActiveChart
      '            .HasTitle = True
      '            .ChartTitle.Characters.Text = "KPI UNIT COVER - UNIT FINISH (TOTAL)"
      '            .Axes(xlCategory, xlPrimary).HasTitle = True
      '            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Unit Cover"
      '            .Axes(xlValue, xlPrimary).HasTitle = True
      '            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Jumlah"
      '        End With
      '        .Sheets("Sheet4").Select
      '        .Sheets("Sheet4").Name = "Cover-Finish-Total"
      '        .Sheets("Cover-Finish-Total").Select

15            .Close SaveChanges:=True, FileName:=onV
16        End With
17        Set oWB = Nothing
18        oApp.Quit
19        Set oApp = Nothing
End Sub

Private Sub addTotalRow()
    With rsLaporan
        .AddNew
        !K1 = nilaiTotal1(0)
        !K10 = nilaiTotal1(1)
        !K20 = nilaiTotal1(2)
        !K30 = nilaiTotal1(3)
        !K40 = nilaiTotal1(4)
        !K50 = nilaiTotal1(5)
        !K60 = nilaiTotal1(6)
        !K70 = nilaiTotal1(7)
        !K80 = nilaiTotal1(8)
        
        !tot16 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
        !tot19 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
        !tot26 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
        !tot29 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
                 
        !judul = "SB Same Outlet"
        !model = "Total"
        !Urut = 1
        .Update
    End With
    
    With rsLaporan
        .AddNew
        !K1 = nilaiTotal2(0)
        !K10 = nilaiTotal2(1)
        !K20 = nilaiTotal2(2)
        !K30 = nilaiTotal2(3)
        !K40 = nilaiTotal2(4)
        !K50 = nilaiTotal2(5)
        !K60 = nilaiTotal2(6)
        !K70 = nilaiTotal2(7)
        !K80 = nilaiTotal2(8)
        
        !tot16 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
        !tot19 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
        !tot26 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
        !tot29 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
                 
        !judul = "Demand SB Same Outlet"
        !model = "Total"
        !Urut = 2
        .Update
    End With
    
    With rsLaporan
        .AddNew
        If nilaiTotal2(0) > 0 Then
            !K1 = Round(nilaiTotal1(0) * 100 / nilaiTotal2(0), 0) & " %"
        Else
            !K1 = "0 %"
        End If
        
        If nilaiTotal2(1) > 0 Then
            !K10 = Round(nilaiTotal1(1) * 100 / nilaiTotal2(1), 0) & " %"
        Else
            !K10 = "0 %"
        End If
                
        If nilaiTotal2(2) > 0 Then
            !K20 = Round(nilaiTotal1(2) * 100 / nilaiTotal2(2), 0) & " %"
        Else
            !K20 = "0 %"
        End If
        
        If nilaiTotal2(3) > 0 Then
            !K30 = Round(nilaiTotal1(3) * 100 / nilaiTotal2(3), 0) & " %"
        Else
            !K30 = "0 %"
        End If
                 
        If nilaiTotal2(4) > 0 Then
            !K40 = Round(nilaiTotal1(4) * 100 / nilaiTotal2(4), 0) & " %"
        Else
            !K40 = "0 %"
        End If

        If nilaiTotal2(5) > 0 Then
            !K50 = Round(nilaiTotal1(5) * 100 / nilaiTotal2(5), 0) & " %"
        Else
            !K50 = "0 %"
        End If
        
        If nilaiTotal2(6) > 0 Then
            !K60 = Round(nilaiTotal1(6) * 100 / nilaiTotal2(6), 0) & " %"
        Else
            !K60 = "0 %"
        End If
        
        If nilaiTotal2(7) > 0 Then
            !K70 = Round(nilaiTotal1(7) * 100 / nilaiTotal2(7), 0) & " %"
        Else
            !K70 = "0 %"
        End If
        
        If nilaiTotal2(8) > 0 Then
            !K80 = Round(nilaiTotal1(8) * 100 / nilaiTotal2(8), 0) & " %"
        Else
            !K80 = "0 %"
        End If
        
        If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
            !tot16 = Round(((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
        Else
            !tot16 = "0 %"
        End If
        
        If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
            !tot19 = Round((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
        Else
            !tot19 = "0 %"
        End If
        
        If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
            !tot26 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
        Else
            !tot26 = "0 %"
        End If
        
        If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
            !tot29 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
        Else
            !tot29 = "0 %"
        End If
               
        !Urut = 3
        !judul = "Retention Rate Same Outlate"
        !model = "Total"
        .Update
    End With
    
    With rsLaporan
        .AddNew
        !K1 = nilaiTotal4(0)
        !K10 = nilaiTotal4(1)
        !K20 = nilaiTotal4(2)
        !K30 = nilaiTotal4(3)
        !K40 = nilaiTotal4(4)
        !K50 = nilaiTotal4(5)
        !K60 = nilaiTotal4(6)
        !K70 = nilaiTotal4(7)
        !K80 = nilaiTotal4(8)
        
        !tot16 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
        !tot19 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
        !tot26 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
        !tot29 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
             
        !judul = "SB Same Group (NCM)"
        !model = "Total"
        !Urut = 4
        .Update
    End With
    
    With rsLaporan
        .AddNew
        If nilaiTotal2(0) > 0 Then
            !K1 = Round(nilaiTotal4(0) * 100 / nilaiTotal2(0), 0) & " %"
        Else
            !K1 = "0 %"
        End If
        
        If nilaiTotal2(1) > 0 Then
            !K10 = Round(nilaiTotal4(1) * 100 / nilaiTotal2(1), 0) & " %"
        Else
            !K10 = "0 %"
        End If
                
        If nilaiTotal2(2) > 0 Then
            !K20 = Round(nilaiTotal4(2) * 100 / nilaiTotal2(2), 0) & " %"
        Else
            !K20 = "0 %"
        End If
        
        If nilaiTotal2(3) > 0 Then
            !K30 = Round(nilaiTotal4(3) * 100 / nilaiTotal2(3), 0) & " %"
        Else
            !K30 = "0 %"
        End If
                 
        If nilaiTotal2(4) > 0 Then
            !K40 = Round(nilaiTotal4(4) * 100 / nilaiTotal2(4), 0) & " %"
        Else
            !K40 = "0 %"
        End If

        If nilaiTotal2(5) > 0 Then
            !K50 = Round(nilaiTotal4(5) * 100 / nilaiTotal2(5), 0) & " %"
        Else
            !K50 = "0 %"
        End If
        
        If nilaiTotal2(6) > 0 Then
            !K60 = Round(nilaiTotal4(6) * 100 / nilaiTotal2(6), 0) & " %"
        Else
            !K60 = "0 %"
        End If
        
        If nilaiTotal2(7) > 0 Then
            !K70 = Round(nilaiTotal4(7) * 100 / nilaiTotal2(7), 0) & " %"
        Else
            !K70 = "0 %"
        End If
        
        If nilaiTotal2(8) > 0 Then
            !K80 = Round(nilaiTotal4(8) * 100 / nilaiTotal2(8), 0) & " %"
        Else
            !K80 = "0 %"
        End If
        
        If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
            !tot16 = Round(((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
        Else
            !tot16 = "0 %"
        End If
        
        If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
            !tot19 = Round((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
        Else
            !tot19 = "0 %"
        End If
        
        If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
            !tot26 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
        Else
            !tot26 = "0 %"
        End If
        
        If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
            !tot29 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
        Else
            !tot29 = "0 %"
        End If
        
        !Urut = 5
        !judul = "Retention Rate Same Group (NCM)"
        !model = "Total"
        .Update
    End With
    
    With rsLaporan
        .AddNew
        !K1 = nilaiTotal6(0)
        !K10 = nilaiTotal6(1)
        !K20 = nilaiTotal6(2)
        !K30 = nilaiTotal6(3)
        !K40 = nilaiTotal6(4)
        !K50 = nilaiTotal6(5)
        !K60 = nilaiTotal6(6)
        !K70 = nilaiTotal6(7)
        !K80 = nilaiTotal6(8)
        
        !tot16 = nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)
        !tot19 = nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)
        !tot26 = nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)
        !tot29 = nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)
             
        !judul = "SB Total (NCM & Non NCM)"
        !model = "Total"
        !Urut = 6
        .Update
    End With
    
    With rsLaporan
        .AddNew
        If nilaiTotal2(0) > 0 Then
            !K1 = Round(nilaiTotal6(0) * 100 / nilaiTotal2(0), 0) & " %"
        Else
            !K1 = "0 %"
        End If
        
        If nilaiTotal2(1) > 0 Then
            !K10 = Round(nilaiTotal6(1) * 100 / nilaiTotal2(1), 0) & " %"
        Else
            !K10 = "0 %"
        End If
                
        If nilaiTotal2(2) > 0 Then
            !K20 = Round(nilaiTotal6(2) * 100 / nilaiTotal2(2), 0) & " %"
        Else
            !K20 = "0 %"
        End If
        
        If nilaiTotal2(3) > 0 Then
            !K30 = Round(nilaiTotal6(3) * 100 / nilaiTotal2(3), 0) & " %"
        Else
            !K30 = "0 %"
        End If
                 
        If nilaiTotal2(4) > 0 Then
            !K40 = Round(nilaiTotal6(4) * 100 / nilaiTotal2(4), 0) & " %"
        Else
            !K40 = "0 %"
        End If

        If nilaiTotal2(5) > 0 Then
            !K50 = Round(nilaiTotal6(5) * 100 / nilaiTotal2(5), 0) & " %"
        Else
            !K50 = "0 %"
        End If
        
        If nilaiTotal2(6) > 0 Then
            !K60 = Round(nilaiTotal6(6) * 100 / nilaiTotal2(6), 0) & " %"
        Else
            !K60 = "0 %"
        End If
        
        If nilaiTotal2(7) > 0 Then
            !K70 = Round(nilaiTotal6(7) * 100 / nilaiTotal2(7), 0) & " %"
        Else
            !K70 = "0 %"
        End If
        
        If nilaiTotal2(8) > 0 Then
            !K80 = Round(nilaiTotal6(8) * 100 / nilaiTotal2(8), 0) & " %"
        Else
            !K80 = "0 %"
        End If
        
        If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
            !tot16 = Round(((nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
        Else
            !tot16 = "0 %"
        End If
        
        If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
            !tot19 = Round((nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
        Else
            !tot19 = "0 %"
        End If
        
        If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
            !tot26 = Round((nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
        Else
            !tot26 = "0 %"
        End If
        
        If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
            !tot29 = Round((nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
        Else
            !tot29 = "0 %"
        End If
        !Urut = 7
        !judul = "Retention Rate Total (NCM & Non NCM)"
        !model = "Total"
        .Update
    End With
End Sub
