VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoSBRateRetentionRev1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form Ratantion Rate"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
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
      Left            =   4950
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2685
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
      Left            =   1740
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2685
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
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2685
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   8280
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
      TabIndex        =   19
      Top             =   480
      Width           =   12375
      Begin VB.Frame frameKolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daftar Kolom"
         Height          =   4695
         Left            =   8040
         TabIndex        =   20
         Top             =   960
         Width           =   4215
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Reminder"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama Customer"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama Model"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Polisi"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Rangka"
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tahun"
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delivery Date"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   36
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Retail Sales Date"
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Model No"
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   34
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alamat Customer"
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   33
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Telpon"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   32
            Top             =   3840
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "History SBI / SBE"
            Height          =   375
            Index           =   11
            Left            =   2160
            TabIndex        =   31
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Next Service"
            Height          =   375
            Index           =   12
            Left            =   2160
            TabIndex        =   30
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dealer Penjual"
            Height          =   375
            Index           =   13
            Left            =   2160
            TabIndex        =   29
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dealer Terakhir SB"
            Height          =   375
            Index           =   14
            Left            =   2160
            TabIndex        =   28
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Wiraniaga"
            Height          =   375
            Index           =   15
            Left            =   2160
            TabIndex        =   27
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cara Bayar"
            Height          =   375
            Index           =   16
            Left            =   2160
            TabIndex        =   26
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama SPK"
            Height          =   375
            Index           =   17
            Left            =   2160
            TabIndex        =   25
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alamat SPK"
            Height          =   375
            Index           =   18
            Left            =   2160
            TabIndex        =   24
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No Telp SPK"
            Height          =   375
            Index           =   19
            Left            =   2160
            TabIndex        =   23
            Top             =   3120
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nama STNK"
            Height          =   375
            Index           =   20
            Left            =   2160
            TabIndex        =   22
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox kolom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alamat STNK"
            Height          =   375
            Index           =   21
            Left            =   2160
            TabIndex        =   21
            Top             =   3840
            Width           =   1815
         End
         Begin AIFCmp1.asxPowerButton cmdTutupDealer 
            Height          =   360
            Left            =   0
            TabIndex        =   43
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
         TabIndex        =   50
         Top             =   1035
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   46
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal retail"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   49
            Top             =   120
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal invoice"
            Height          =   255
            Index           =   7
            Left            =   105
            TabIndex        =   48
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal delivery"
            Height          =   255
            Index           =   8
            Left            =   5520
            TabIndex        =   47
            Top             =   120
            Visible         =   0   'False
            Width           =   2415
         End
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
         TabIndex        =   45
         Top             =   1035
         Width           =   3255
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
         TabIndex        =   44
         Top             =   1080
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker tglAwalVin 
         Height          =   375
         Left            =   1410
         TabIndex        =   51
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
         Format          =   105775105
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker tglAkhirVin 
         Height          =   375
         Left            =   3000
         TabIndex        =   52
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
         Format          =   105775105
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgVin 
         Height          =   615
         Left            =   480
         TabIndex        =   53
         Top             =   1560
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
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
         TabIndex        =   56
         Top             =   1800
         Visible         =   0   'False
         Width           =   6255
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
         TabIndex        =   55
         Top             =   1080
         Width           =   840
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
         TabIndex        =   54
         Top             =   1080
         Width           =   60
      End
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5550
      TabIndex        =   15
      Top             =   2715
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   3225
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2685
      Visible         =   0   'False
      Width           =   1185
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   2700
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
      MICON           =   "frmRepoSBRateRetentionRev1.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   375
         Left            =   1650
         TabIndex        =   5
         Top             =   210
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
         Format          =   105775105
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   210
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
         Format          =   105775105
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   720
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
         Top             =   240
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
         Top             =   240
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
      FormatString    =   "   Laporan Ratention Rate"
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
   Begin CRVIEWERLibCtl.CRViewer CRViewer2 
      Height          =   2175
      Left            =   0
      TabIndex        =   57
      Top             =   0
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
End
Attribute VB_Name = "frmRepoSBRateRetentionRev1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report                  As Object
Private rsTemp              As ADODB.Recordset
Private rsLaporan           As New ADODB.Recordset
Private rsDownload          As New ADODB.Recordset
Private rsCari              As New ADODB.Recordset
Private i                   As Integer
Private nilaiTotal1(20)     As Double
Private nilaiTotal2(20)     As Double
Private nilaiTotal3(20)     As Double
Private nilaiTotal4(20)     As Double
Private nilaiTotal5(20)     As Double
Private nilaiTotal6(20)     As Double
Private nilaiTotal7(20)     As Double
Private nilaiTotal8(20)     As Double


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


Private Sub cmdDownloadVin_Click()
1         On Error Resume Next
          Dim rsNorangka      As ADODB.Recordset
          Dim rsData          As ADODB.Recordset
          Dim isDownloadAll   As Boolean
          
2         If DateDiff("D", Format(tglAwalVin, "dd/mm/yyyy"), Format(tglAkhirVin, "dd/mm/yyyy")) < 0 Then
3             MsgBox "Tanggal awal laporan tidak boleh lebih besar dari tanggal akhir laporan", vbInformation
4             Exit Sub
5         End If
          
      '    If Month(tglAwalVin) <> Month(tglAkhirVin) Or year(tglAwalVin) <> year(tglAkhirVin) Then
      '        MsgBox "Range periode dalam 1 bulan", vbInformation
      '        Exit Sub
      '    End If
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

         
          
20        sSQL = "select b.NoRangka,b.kode_sales,case when b.Syarat_Bayar = 0 then 'TUNAI' else 'KREDIT' end bayarcust,b.Nama_Faktur,b.Alamat1_Faktur,b.Kota,b.Kecamatan,b.Kode_Pos,b.HP1,b.HP2,f.propinsi,b.telp_rumah as cdma"
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
          
          ''20201215 daftar noka yang seharusnya kembali service (reminder) tapi belum service kembali
          ''--------------------------------------------------------------------------------------------
33        sSQL = "select distinct m.NoRangka,b.kode_sales,case when b.Syarat_Bayar = 0 then 'TUNAI' when b.Syarat_Bayar = 1 then 'KREDIT' else '' end bayarcust,b.Nama_Faktur,b.Alamat1_Faktur,b.Kota,b.Kecamatan,b.Kode_Pos,b.HP1,b.HP2,f.propinsi"
34        sSQL = sSQL & " from M_ESTIMASI m"
35        sSQL = sSQL & " left JOIN [SPK" & Right(gDb, 2) & "]..spk b ON m.NoRangka = b.NoRangka and b.SPK_Status = 1"
36        sSQL = sSQL & " left join SingleIDOri..Bridging e on b.Kode_Lgn = e.KodeLgn and e.KodeDealer = '" & Right(gDb, 2) & "'"
37        sSQL = sSQL & " left join SingleIDOri..NewCustomer f on e.IDCust = f.IDCust"
38        sSQL = sSQL & " left join SingleIDOri..UnitCustomer g on b.NoRangka = g.NoRangka"
39        sSQL = sSQL & " left join SingleIDOri..NewCustomer H on h.IDCust = g.IDCustSTNK"
40        sSQL = sSQL & " where (CASE WHEN m.JT_Tgl > m.JT_Bln then  m.JT_Bln WHEN m.JT_Tgl < m.JT_Bln then  m.JT_Tgl WHEN m.JT_Tgl = m.JT_Bln then  m.Jt_Bln END) between '" & Format(tglAwalVin, "yyyy-mm-dd") & "' and '" & Format(tglAkhirVin, "yyyy-mm-dd") & "'"
41        sSQL = sSQL & " and m.SBENext < 85000 and m.NoWo = ''"
42        sSQL = sSQL & " and m.NoRangka  not in ("
43        sSQL = sSQL & " select no_chassis from T_INVOICE where  NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
44        sSQL = sSQL & " and FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
45        sSQL = sSQL & " and TPSVC <> 'w' and PDS = 0)"
46        Set rsNorangka = objO.FetchData(sSQL)
47        If rsNorangka.RecordCount = 0 Then
48            MsgBox "Tidak ada data untuk di download", vbInformation
49            Exit Sub
50        Else
51            rsNorangka.MoveFirst
52            ProgVin.Max = rsNorangka.RecordCount + 1
53            ProgVin.value = 1
54            ProgVin.Visible = True
55            DoEvents
56            Do While Not rsNorangka.EOF
57                sSQL = "select"
58                sSQL = sSQL & " ISNULL((select top 1 tglfollow from m_listremindercall where norangka = a.NoRangka order by tglfollow desc),'1900-01-01') lastreminder,"
59                sSQL = sSQL & " b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3 as namacustomer,"
60                sSQL = sSQL & " a.Fmdl as namamodel,a.NoPolisi,a.NoRangka,a.Thmobil,a.Ddate,"
                  ''sSQL = sSQL & " ISNULL((select top 1 TglFaktur from [SPK" & Right(gDb, 2) & "]..Sihdr where NoRangka = a.NoRangka order by TglFaktur desc),'1900-01-01') as tglretail,"
61                sSQL = sSQL & " ISNULL((select top 1 Tanggal from [ReplBengkel99]..Summary where NoRangka = a.NoRangka and TipeTr = 'ND'  and [Status] = 1),'1900-01-01') AS tglretail,"
62                sSQL = sSQL & " a.Mdl,b.Alamat,b.Kelurahan,b.Kecamatan,b.Kota,b.Propinsi,b.KodePos,b.IdTelp+b.Telp AS TELP,b.HP1,b.HP2,a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV  as usename,"
63                sSQL = sSQL & " a.HP1SV as usehp,ISNULL((select Top 1 fs_rc from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''  order by invdate desc),'') tipelastsb,"
64                sSQL = sSQL & " ISNULL((select Top 1 invdate from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> '' order by invdate desc),'') tgllastsb,"
65                sSQL = sSQL & " ISNULL((select Top 1 FS_NM_DEALER from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''  order by invdate desc),'') dealerlassb,"
66                sSQL = sSQL & " ISNULL((select Top 1 FRONT from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' and fs_rc <> ''  order by invdate desc),'') lastsa,"
                  ''sSQL = sSQL & " ISNULL((select top 1 (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) from m_estimasi where NoRangka = a.NoRangka and nowo = '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) asc ),'1900-01-01') tglnextservice,"
67                sSQL = sSQL & " ISNULL((select top 1 (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) from [ReplBengkel99]..m_estimasi where NoRangka = a.NoRangka and nowo <> '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) desc ),'1900-01-01') tglnextservice,"
                  ''sSQL = sSQL & " ISNULL((select top 1 sbenext from m_estimasi where NoRangka = a.NoRangka and nowo = '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) asc ),0) as sbnext,"
68                sSQL = sSQL & " ISNULL((select top 1 sbenext from [ReplBengkel99]..m_estimasi where NoRangka = a.NoRangka and nowo <> '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) desc ),0) as sbnext,"
69                sSQL = sSQL & " a.NmDlrPenjual,"
70                sSQL = sSQL & " ISNULL((select Top 1 FS_NM_DEALER from [ReplBengkel99]..T_INVOICEnew where no_chassis = a.NoRangka and fs_log_del = '' order by invdate desc),'') dealerterakhirsv,"
71                sSQL = sSQL & " a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK as namastnk,"
72                sSQL = sSQL & " a.AlamatSTNK , a.PropinsiSTNK, a.KotaSTNK, a.KecamatanSTNK, a.KodePosSTNK,"
73                sSQL = sSQL & " DATEDIFF(month,  isnull((select top 1 DATE from ("
74                sSQL = sSQL & " select top 2 DATE from [ReplBengkel99]..T_INVOICEnew  where FS_RC <> '' and FS_LOG_DEL = '' and NO_CHASSIS = a.norangka and NOINV not like '%b' order by INVDATE desc"
75                sSQL = sSQL & " ) z order by DATE asc),GETDATE()),ISNULL((select top 1 DATE from [ReplBengkel99]..T_INVOICEnew  where FS_RC <> '' and FS_LOG_DEL = '' and NO_CHASSIS = a.norangka and NOINV not like '%b' order by INVDATE desc),GETDATE())) selisihSB"
76                sSQL = sSQL & " from SingleIDOri..UnitCustomer a"
77                sSQL = sSQL & " left join SingleIDOri..NewCustomer b on b.IDCust = case when a.IDCustSales <> '' then a.IDCustSales when a.idcustservice <> '' then a.idcustservice   end"
78                sSQL = sSQL & " where NoRangka = '" & NulSID(rsNorangka!noRangka) & "'"
79                Set rsData = objO.FetchData(sSQL)
80                If Not rsData.EOF Then
81                    rsData.MoveFirst
82                    If Log.RPLHistori = 0 Then
83                        Set rsCari = objO.FetchDataHOUnlimitedTime("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM T_INVOICENEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(rsNorangka!noRangka) & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B'  order by INVDATE desc")
84                    Else
85                        Set rsCari = objO.FetchData("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM [" & dbReplikasiHO & "]..T_INVOICENEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(rsNorangka!noRangka) & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B' order by INVDATE desc")
86                    End If
87                    With rsDownload
88                        .AddNew
89                        If kolom(0).value = vbChecked Then
90                            !lastReminder = Format(rsData!lastReminder, "dd-mm-yyyy")
91                        End If
92                        If kolom(1).value = vbChecked Then
93                            !customer = Left(NulSID(rsData!namaCustomer), 200)
94                        End If
95                        If kolom(2).value = vbChecked Then
96                            !Fmdl = Left(NulSID(rsData!namaModel), 100)
97                        End If
98                        If kolom(3).value = vbChecked Then
99                            !NoPolisi = Left(NulSID(rsData!NoPolisi), 40)
100                       End If
101                       If kolom(4).value = vbChecked Then
102                           !noRangka = Left(NulSID(rsData!noRangka), 100)
103                       End If
104                       If kolom(5).value = vbChecked Then
105                           !Tahun = Left(NulSID(rsData!thMobil), 10)
106                       End If
107                       If kolom(6).value = vbChecked Then
108                           !dDate = Format(rsData!dDate, "dd-mm-yyyy")
109                       End If
110                       If kolom(7).value = vbChecked Then
111                           !retailSales = Left(Format(NulSID(rsData!tglretail), "DD-MM-YYYY"), 50)
112                       End If
113                       If kolom(8).value = vbChecked Then
114                           !MDL = Left(NulSID(rsData!MDL), 100)
115                       End If
116                       If kolom(9).value = vbChecked Then
117                           If isDownloadAll Then
118                               !Alamat = Left(NulSID(rsData!Alamat), 300)
119                           End If
120                           !kelurahan = Left(NulSID(rsData!kelurahan), 100)
121                           !kecamatan = Left(NulSID(rsData!kecamatan), 100)
122                           !Kota = Left(NulSID(rsData!Kota), 100)
123                           !Propinsi = Left(NulSID(rsData!Propinsi), 100)
124                           !KodePos = Left(NulSID(rsData!KodePos), 100)
125                       End If
                          
126                       If kolom(10).value = vbChecked Then
127                           !PTSN = Left(NulSID(rsData!telp), 100)
128                           !hp1 = Left(NulSID(rsData!hp1), 100)
129                           !hp2 = Left(NulSID(rsData!hp2), 100)
130                           !UseName = Left(NulSID(rsData!UseName), 200)
131                           !USEHP = Left(NulSID(rsData!hp1), 100)
132                       End If
                          
133                       If kolom(11).value = vbChecked Then
134                           !K1 = GetService("R001", True) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
135                           !dealerK1 = GetServiceDealer("R001", True)
                              'sbe
136                           !k5 = GetService("R005", True) 'IIf(IsNull(Trim(rsTagih("5K"))), "", Trim(rsTagih("5K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
137                           !dealerK5 = GetServiceDealer("R005", True)
138                           !K10 = GetService("R010", True) 'IIf(IsNull(Trim((rsTagih("10K")))), "", Trim(rsTagih("10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
139                           !dealerK10 = GetServiceDealer("R010", True)
140                           !k15 = GetService("R015", True) 'IIf(IsNull(Trim((rsTagih("15K")))), "", Trim(rsTagih("15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
141                           !dealerK15 = GetServiceDealer("R015", True)
142                           !K20 = GetService("R020", True) 'IIf(IsNull(Trim((rsTagih("20K")))), "", Trim(rsTagih("20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
143                           !dealerK20 = GetServiceDealer("R020", True)
144                           !k25 = GetService("R025", True) 'IIf(IsNull(Trim((rsTagih("25K")))), "", Trim(rsTagih("25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
145                           !dealerK25 = GetServiceDealer("R025", True)
146                           !K30 = GetService("R030", True) 'IIf(IsNull(Trim((rsTagih("30K")))), "", Trim(rsTagih("30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
147                           !dealerK30 = GetServiceDealer("R030", True)
148                           !k35 = GetService("R035", True) 'IIf(IsNull(Trim((rsTagih("35K")))), "", Trim(rsTagih("35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
149                           !dealerK35 = GetServiceDealer("R035", True)
150                           !K40 = GetService("R040", True) 'IIf(IsNull(Trim((rsTagih("40K")))), "", Trim(rsTagih("40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
151                           !dealerK40 = GetServiceDealer("R040", True)
152                           !k45 = GetService("R045", True) 'IIf(IsNull(Trim((rsTagih("45K")))), "", Trim(rsTagih("45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
153                           !dealerK45 = GetServiceDealer("R045", True)
154                           !K50 = GetService("R050", True) 'IIf(IsNull(Trim((rsTagih("50K")))), "", Trim(rsTagih("50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
155                           !dealerK50 = GetServiceDealer("R050", True)
156                           !k55 = GetService("R055", True) 'IIf(IsNull(Trim((rsTagih("55K")))), "", Trim(rsTagih("55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
157                           !dealerK55 = GetServiceDealer("R055", True)
158                           !K60 = GetService("R060", True) 'IIf(IsNull(Trim((rsTagih("60K")))), "", Trim(rsTagih("60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
159                           !dealerK60 = GetServiceDealer("R060", True)
160                           !k65 = GetService("R065", True) 'IIf(IsNull(Trim((rsTagih("65K")))), "", Trim(rsTagih("65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
161                           !dealerK65 = GetServiceDealer("R065", True)
162                           !K70 = GetService("R070", True) 'IIf(IsNull(Trim((rsTagih("70K")))), "", Trim(rsTagih("70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
163                           !dealerK70 = GetServiceDealer("R070", True)
164                           !k75 = GetService("R075", True) 'IIf(IsNull(Trim((rsTagih("75K")))), "", Trim(rsTagih("75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
165                           !dealerK75 = GetServiceDealer("R075", True)
166                           !K80 = GetService("R080", True) 'IIf(IsNull(Trim((rsTagih("80K")))), "", Trim(rsTagih("80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
167                           !dealerK80 = GetServiceDealer("R080", True)
168                           !k85 = GetService("R085", True) 'IIf(IsNull(Trim((rsTagih("85K")))), "", Trim(rsTagih("85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
169                           !dealerK85 = GetServiceDealer("R085", True)
170                           !k90 = GetService("R090", True) 'IIf(IsNull(Trim((rsTagih("90K")))), "", Trim(rsTagih("90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
171                           !dealerK90 = GetServiceDealer("R090", True)
172                           !k95 = GetService("R095", True) 'IIf(IsNull(Trim((rsTagih("95K")))), "", Trim(rsTagih("95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
173                           !dealerK95 = GetServiceDealer("R095", True)
174                           !k100 = GetService("R100", True) 'IIf(IsNull(Trim((rsTagih("100K")))), "", Trim(rsTagih("100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
175                           !dealerK100 = GetServiceDealer("R100", True)
176                           !k105 = GetService("R105", True) 'IIf(IsNull(Trim((rsTagih("105K")))), "", Trim(rsTagih("105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
177                           !dealerK105 = GetServiceDealer("R105", True)
178                           !k110 = GetService("R110", True) 'IIf(IsNull(Trim((rsTagih("110K")))), "", Trim(rsTagih("110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
179                           !dealerK110 = GetServiceDealer("R110", True)
180                           !k115 = GetService("R115", True) 'IIf(IsNull(Trim((rsTagih("115K")))), "", Trim(rsTagih("115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
181                           !dealerK115 = GetServiceDealer("R115", True)
182                           !k120 = GetService("R120", True) 'IIf(IsNull(Trim((rsTagih("120K")))), "", Trim(rsTagih("120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
183                           !dealerK120 = GetServiceDealer("R120", True)
184                           !k125 = GetService("R125", True) 'IIf(IsNull(Trim((rsTagih("125K")))), "", Trim(rsTagih("125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
185                           !dealerK125 = GetServiceDealer("R125", True)
186                           !k130 = GetService("R130", True) 'IIf(IsNull(Trim((rsTagih("130K")))), "", Trim(rsTagih("130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
187                           !dealerK130 = GetServiceDealer("R130", True)
188                           !k135 = GetService("R135", True) 'IIf(IsNull(Trim((rsTagih("135K")))), "", Trim(rsTagih("135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
189                           !dealerK135 = GetServiceDealer("R135", True)
190                           !k140 = GetService("R140", True) 'IIf(IsNull(Trim((rsTagih("140K")))), "", Trim(rsTagih("140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
191                           !dealerK140 = GetServiceDealer("R140", True)
192                           !k145 = GetService("R145", True) 'IIf(IsNull(Trim((rsTagih("145K")))), "", Trim(rsTagih("145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
193                           !dealerK145 = GetServiceDealer("R145", True)
194                           !k150 = GetService("R150", True) 'IIf(IsNull(Trim((rsTagih("150K")))), "", Trim(rsTagih("150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
195                           !dealerK150 = GetServiceDealer("R150", True)
196                           !k155 = GetService("R155", True) 'IIf(IsNull(Trim((rsTagih("155K")))), "", Trim(rsTagih("155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
197                           !dealerK155 = GetServiceDealer("R155", True)
198                           !k160 = GetService("R160", True) 'IIf(IsNull(Trim((rsTagih("160K")))), "", Trim(rsTagih("160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
199                           !dealerK160 = GetServiceDealer("R160", True)
200                           !k165 = GetService("R165", True) 'IIf(IsNull(Trim((rsTagih("165K")))), "", Trim(rsTagih("165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
201                           !dealerK165 = GetServiceDealer("R165", True)
202                           !k170 = GetService("R170", True) 'IIf(IsNull(Trim((rsTagih("170K")))), "", Trim(rsTagih("170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
203                           !dealerK170 = GetServiceDealer("R170", True)
204                           !k175 = GetService("R175", True) 'IIf(IsNull(Trim((rsTagih("175K")))), "", Trim(rsTagih("175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
205                           !dealerK175 = GetServiceDealer("R175", True)
206                           !k180 = GetService("R180", True) 'IIf(IsNull(Trim((rsTagih("180K")))), "", Trim(rsTagih("180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
207                           !dealerK180 = GetServiceDealer("R180", True)
208                           !k185 = GetService("R185", True) 'IIf(IsNull(Trim((rsTagih("185K")))), "", Trim(rsTagih("185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
209                           !dealerK185 = GetServiceDealer("R185", True)
210                           !k190 = GetService("R190", True) 'IIf(IsNull(Trim((rsTagih("190K")))), "", Trim(rsTagih("190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
211                           !dealerK190 = GetServiceDealer("R190", True)
212                           !k195 = GetService("R195", True) 'IIf(IsNull(Trim((rsTagih("195K")))), "", Trim(rsTagih("195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
213                           !dealerK195 = GetServiceDealer("R195", True)
214                           !k200 = GetService("R200", True) 'IIf(IsNull(Trim((rsTagih("200K")))), "", Trim(rsTagih("200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
215                           !dealerK200 = GetServiceDealer("R200", True)
                              
216                           !tipelastsb = IIf(NulSID(rsData!tipelastsb) <> "", CInt(Right(NulSID(rsData!tipelastsb), 3)) & " K", "") ''Left(NulSID(rsData!tipelastsb), 100)
217                           !LASTSVCSB = Left(NulSID(rsData!tgllastsb), 100)
218                           !UMURLASTSB = Left(NulSID(rsData!selisihsb), 100) 'Left(NulSID(rsData!tipelastsb), 100)
219                           !dealerLastSB = Left(NulSID(rsData!dealerlassb), 100)
220                           !SALAST = Left(NulSID(rsData!LastSA), 100)
                               
                              
221                       End If
222                       If kolom(12).value = vbChecked Then
223                           !NEXTSB = Left(NulSID(rsData!SBNEXT), 100)
224                           !TGLNEXTSB = Left(NulSID(rsData!tglnextservice), 100)
225                       End If
226                       If kolom(13).value = vbChecked Then
227                           !DealerPenjual = Left(NulSID(rsData!NmDlrPenjual), 200)
228                       End If
                                              
229                       If kolom(14).value = vbChecked Then
230                           !DEALERTERAKHIRSB = Left(NulSID(rsData!dealerlassb), 100)
231                       End If
                          
232                       If kolom(15).value = vbChecked Then
233                           !WIRANIAGA = Left(NulSID(rsNorangka!kode_sales), 100)
234                       End If
                           
235                       If isDownloadAll Then
236                           If kolom(16).value = vbChecked Then
237                               !CARABAYAR = Left(NulSID(rsNorangka!bayarcust), 100)
238                           End If
239                           If kolom(17).value = vbChecked Then
240                               !NAMASPK = Left(NulSID(rsNorangka!Nama_Faktur), 100)
241                           End If
242                           If kolom(18).value = vbChecked Then
243                               !ALAMATSPK = Left(NulSID(rsNorangka!Alamat1_Faktur), 100)
244                               !PROPINSISPK = Left(NulSID(rsData!Propinsi), 100)
245                               !KOTASPK = Left(NulSID(rsNorangka!Kota), 100)
246                               !KECSPK = Left(NulSID(rsNorangka!kecamatan), 100)
247                               !KODEPOSSPK = Left(NulSID(rsNorangka!Kode_Pos), 100)
248                               !CDMASPK = Left(NulSID(rsNorangka!hp2), 100)
249                           End If
250                           If kolom(19).value = vbChecked Then
251                               !CDMASPK = Left(NulSID(rsNorangka!CDMA), 100)
252                               !HP1SPK = Left(NulSID(rsNorangka!hp1), 100)
253                               !HP2SPK = Left(NulSID(rsNorangka!hp2), 100)
254                           End If
255                           If kolom(20).value = vbChecked Then
256                               !namastnk = Left(NulSID(rsData!namastnk), 100)
257                           End If
258                           If kolom(21).value = vbChecked Then
259                               !AlamatSTNK = Left(NulSID(rsData!AlamatSTNK), 100)
260                               !PropinsiSTNK = Left(NulSID(rsData!PropinsiSTNK), 100)
261                               !Kotastnk = Left(NulSID(rsData!Kotastnk), 100)
262                               !KECSTNK = Left(NulSID(rsData!KecamatanSTNK), 100)
263                               !KodePosSTNK = Left(NulSID(rsData!KodePosSTNK), 100)
264                           End If
265                       End If
266                       .Update
267                   End With
268               End If
                  
269               rsNorangka.MoveNext
270               ProgVin.value = ProgVin.value + 1
271               DoEvents
272           Loop
273       End If
274       ProgVin.Visible = False
          Dim Er As New ErrInfo
275       If rsDownload.RecordCount = 0 Then
276           Er.HeadCap = "Export To Excel !"
277           Er.InfoCap = "List Data Reminder Service Kosong !"
278           Beep
279           Er.show (vbModal)
280           Exit Sub
281       End If
          Dim MyPath As String
          Dim onVal As String
282       MyPath = ""
          Dim f As New FormExplorer
283       f.show (vbModal)
284       MyPath = f.PathSaving
285       If MyPath = "" Then Exit Sub
286       If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
287       If f.FileName = "" Then Exit Sub
288       onVal = MyPath & f.FileName & ".XLS"
          
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
289       sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
290       sSQL = sSQL & " values ('" & gSA & "','" & Me.Name & "',getdate(),'" & Format(tglAwalVin.value, "MM/DD/YYYY") & "','" & Format(tglAkhirVin.value, "MM/DD/YYYY") & "','','" & rsDownload.RecordCount & "','" & onVal & "','" & ComptName & "') "
291       Call objO.InsertData(sSQL)
          
292       Call ProsesExport(onVal, isDownloadAll)
293       Er.HeadCap = "Export to excel !"
294       Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                       "File : " & onVal
295       Er.show (vbModal)
End Sub


Private Sub ProsesExport(onV As String, isDownloadAll)
          Dim oApp    As Excel.Application
          Dim oWB     As Excel.Workbook
          Dim nAsc    As Integer
          Dim C       As Integer
1         Set oApp = New Excel.Application
2         oApp.Visible = False
3         Set oWB = oApp.Workbooks.Add
4         With oWB
              '---------------------------
              'SA FINISH
              '---------------------------
5             nAsc = 64
      '        For C = 0 To rsDownload.Fields.Count - 1
      '            .Sheets("Sheet1").Cells(1, C + 1).value = rsDownload.Fields(C).Name
      '            nAsc = nAsc + 1
      '        Next
6             C = 1
7             If kolom(0).value = vbChecked Then
8                 .Sheets("Sheet1").Cells(1, C).value = "Last Reminder"
9                 C = C + 1
10            End If
11            If kolom(1).value = vbChecked Then
12                .Sheets("Sheet1").Cells(1, C).value = "Nama Customer"
13                C = C + 1
14            End If
15            If kolom(2).value = vbChecked Then
16                .Sheets("Sheet1").Cells(1, C).value = "Nama Model"
17                C = C + 1
18            End If
19            If kolom(3).value = vbChecked Then
20                .Sheets("Sheet1").Cells(1, C).value = "No Polisi"
21                C = C + 1
22            End If
23            If kolom(4).value = vbChecked Then
24                .Sheets("Sheet1").Cells(1, C).value = "No Rangka"
25                C = C + 1
26            End If
27            If kolom(5).value = vbChecked Then
28                .Sheets("Sheet1").Cells(1, C).value = "Tahun"
29                C = C + 1
30            End If
31            If kolom(6).value = vbChecked Then
32                .Sheets("Sheet1").Cells(1, C).value = "Delivery Date"
33                C = C + 1
34            End If
35            If kolom(7).value = vbChecked Then
36                .Sheets("Sheet1").Cells(1, C).value = "Retail Sales Date"
37                C = C + 1
38            End If
39            If kolom(8).value = vbChecked Then
40                .Sheets("Sheet1").Cells(1, C).value = "Model No"
41                C = C + 1
42            End If
43            If kolom(9).value = vbChecked Then
44                If isDownloadAll Then
45                    .Sheets("Sheet1").Cells(1, C).value = "Alamat Customer"
46                    C = C + 1
47                End If
48                .Sheets("Sheet1").Cells(1, C).value = "Kelurahan"
49                C = C + 1
50                .Sheets("Sheet1").Cells(1, C).value = "Kecamatan"
51                C = C + 1
52                .Sheets("Sheet1").Cells(1, C).value = "Kota / Kabupaten"
53                C = C + 1
54                .Sheets("Sheet1").Cells(1, C).value = "Provinsi"
55                C = C + 1
56                .Sheets("Sheet1").Cells(1, C).value = "Kode Pos"
57                C = C + 1
58            End If
              
59            If kolom(10).value = vbChecked Then
                  
60                .Sheets("Sheet1").Cells(1, C).value = "PTSN"
61                C = C + 1
                 
62                .Sheets("Sheet1").Cells(1, C).value = "HP 1"
63                C = C + 1
                  
64                .Sheets("Sheet1").Cells(1, C).value = "HP 2"
65                C = C + 1
                  
66                .Sheets("Sheet1").Cells(1, C).value = "Usename"
67                C = C + 1
                  
68                .Sheets("Sheet1").Cells(1, C).value = "Usehp"
69                C = C + 1
70            End If
              
71            If kolom(11).value = vbChecked Then
                  
72                .Sheets("Sheet1").Cells(1, C).value = "SBI 1K"
73                C = C + 1
74                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBI 1K"
75                C = C + 1
                  
76                .Sheets("Sheet1").Cells(1, C).value = "SBE 5K"
77                C = C + 1
78                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 5K"
79                C = C + 1
                  
80                .Sheets("Sheet1").Cells(1, C).value = "SBE 10K"
81                C = C + 1
82                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 10K"
83                C = C + 1
                  
84                .Sheets("Sheet1").Cells(1, C).value = "SBE 15K"
85                C = C + 1
86                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 15K"
87                C = C + 1
                  
88                .Sheets("Sheet1").Cells(1, C).value = "SBE 20K"
89                C = C + 1
90                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 20K"
91                C = C + 1
                  
92                .Sheets("Sheet1").Cells(1, C).value = "SBE 25K"
93                C = C + 1
94                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 25K"
95                C = C + 1
                  
96                .Sheets("Sheet1").Cells(1, C).value = "SBE 30K"
97                C = C + 1
98                .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 30K"
99                C = C + 1
                  
100               .Sheets("Sheet1").Cells(1, C).value = "SBE 35K"
101               C = C + 1
102               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 35K"
103               C = C + 1
                  
104               .Sheets("Sheet1").Cells(1, C).value = "SBE 40K"
105               C = C + 1
106               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 40K"
107               C = C + 1
                  
108               .Sheets("Sheet1").Cells(1, C).value = "SBE 45K"
109               C = C + 1
110               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 45K"
111               C = C + 1
                  
112               .Sheets("Sheet1").Cells(1, C).value = "SBE 50K"
113               C = C + 1
114               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 50K"
115               C = C + 1
                  
116               .Sheets("Sheet1").Cells(1, C).value = "SBE 55K"
117               C = C + 1
118               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 55K"
119               C = C + 1
                  
120               .Sheets("Sheet1").Cells(1, C).value = "SBE 60K"
121               C = C + 1
122               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 60K"
123               C = C + 1
                  
124               .Sheets("Sheet1").Cells(1, C).value = "SBE 65K"
125               C = C + 1
126               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 65K"
127               C = C + 1
                  
128               .Sheets("Sheet1").Cells(1, C).value = "SBE 70K"
129               C = C + 1
130               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 70K"
131               C = C + 1
                  
132               .Sheets("Sheet1").Cells(1, C).value = "SBE 75K"
133               C = C + 1
134               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 75K"
135               C = C + 1
                  
136               .Sheets("Sheet1").Cells(1, C).value = "SBE 80K"
137               C = C + 1
138               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 80K"
139               C = C + 1
                  
140               .Sheets("Sheet1").Cells(1, C).value = "SBE 85K"
141               C = C + 1
142               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 85K"
143               C = C + 1
                  
144               .Sheets("Sheet1").Cells(1, C).value = "SBE 90K"
145               C = C + 1
146               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 90K"
147               C = C + 1
                  
148               .Sheets("Sheet1").Cells(1, C).value = "SBE 95K"
149               C = C + 1
150               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 95K"
151               C = C + 1
                  
152               .Sheets("Sheet1").Cells(1, C).value = "SBE 100K"
153               C = C + 1
154               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 100K"
155               C = C + 1
                  
156               .Sheets("Sheet1").Cells(1, C).value = "SBE 105K"
157               C = C + 1
158               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 105K"
159               C = C + 1
                  
160               .Sheets("Sheet1").Cells(1, C).value = "SBE 110K"
161               C = C + 1
162               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 110K"
163               C = C + 1
                  
164               .Sheets("Sheet1").Cells(1, C).value = "SBE 115K"
165               C = C + 1
166               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 115K"
167               C = C + 1
                  
168               .Sheets("Sheet1").Cells(1, C).value = "SBE 120K"
169               C = C + 1
170               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 120K"
171               C = C + 1
                  
172               .Sheets("Sheet1").Cells(1, C).value = "SBE 125K"
173               C = C + 1
174               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 125K"
175               C = C + 1
                  
176               .Sheets("Sheet1").Cells(1, C).value = "SBE 130K"
177               C = C + 1
178               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 130K"
179               C = C + 1
                  
180               .Sheets("Sheet1").Cells(1, C).value = "SBE 135K"
181               C = C + 1
182               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 135K"
183               C = C + 1
                  
184               .Sheets("Sheet1").Cells(1, C).value = "SBE 140K"
185               C = C + 1
186               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 140K"
187               C = C + 1
                  
188               .Sheets("Sheet1").Cells(1, C).value = "SBE 145K"
189               C = C + 1
190               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 145K"
191               C = C + 1
                  
192               .Sheets("Sheet1").Cells(1, C).value = "SBE 150K"
193               C = C + 1
194               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 150K"
195               C = C + 1
                  
196               .Sheets("Sheet1").Cells(1, C).value = "SBE 155K"
197               C = C + 1
198               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 155K"
199               C = C + 1
                  
200               .Sheets("Sheet1").Cells(1, C).value = "SBE 160K"
201               C = C + 1
202               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 160K"
203               C = C + 1
                  
204               .Sheets("Sheet1").Cells(1, C).value = "SBE 165K"
205               C = C + 1
206               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 165K"
207               C = C + 1
                  
208               .Sheets("Sheet1").Cells(1, C).value = "SBE 170K"
209               C = C + 1
210               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 170K"
211               C = C + 1
                  
212               .Sheets("Sheet1").Cells(1, C).value = "SBE 175K"
213               C = C + 1
214               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 175K"
215               C = C + 1
                  
216               .Sheets("Sheet1").Cells(1, C).value = "SBE 180K"
217               C = C + 1
218               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 180K"
219               C = C + 1
                  
220               .Sheets("Sheet1").Cells(1, C).value = "SBE 185K"
221               C = C + 1
222               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 185K"
223               C = C + 1
                  
224               .Sheets("Sheet1").Cells(1, C).value = "SBE 190K"
225               C = C + 1
226               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 190K"
227               C = C + 1
                  
228               .Sheets("Sheet1").Cells(1, C).value = "SBE 195K"
229               C = C + 1
230               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 195K"
231               C = C + 1
                  
232               .Sheets("Sheet1").Cells(1, C).value = "SBE 200K"
233               C = C + 1
234               .Sheets("Sheet1").Cells(1, C).value = "Dealer SBE 200K"
235               C = C + 1
                  
236               .Sheets("Sheet1").Cells(1, C).value = "Tipe Last SB"
237               C = C + 1
                  
238               .Sheets("Sheet1").Cells(1, C).value = "Last SVC SB"
239               C = C + 1
                  
240               .Sheets("Sheet1").Cells(1, C).value = "Umur Last SB"
241               C = C + 1
                  
242               .Sheets("Sheet1").Cells(1, C).value = "Dealer Last SVC SB"
243               C = C + 1
                  
244               .Sheets("Sheet1").Cells(1, C).value = "SA Last"
245               C = C + 1
246           End If
247           If kolom(12).value = vbChecked Then
248               .Sheets("Sheet1").Cells(1, C).value = "Suggest Next Service"
249               C = C + 1
                  
250               .Sheets("Sheet1").Cells(1, C).value = "Tanggal Next Service"
251               C = C + 1
252           End If
              
              
253           If kolom(13).value = vbChecked Then
254               .Sheets("Sheet1").Cells(1, C).value = "Dealer Penjual"
255               C = C + 1
256           End If
              
257           If kolom(14).value = vbChecked Then
258               .Sheets("Sheet1").Cells(1, C).value = "Dealer Terakhir SV"
259               C = C + 1
260           End If
261           If kolom(15).value = vbChecked Then
262               .Sheets("Sheet1").Cells(1, C).value = "Wiraniaga"
263               C = C + 1
264           End If
265           If isDownloadAll Then
266               If kolom(16).value = vbChecked Then
267                   .Sheets("Sheet1").Cells(1, C).value = "Cara Bayar"
268                   C = C + 1
269               End If
270               If kolom(17).value = vbChecked Then
271                   .Sheets("Sheet1").Cells(1, C).value = "Nama SPK"
272                   C = C + 1
273               End If
274               If kolom(18).value = vbChecked Then
275                   .Sheets("Sheet1").Cells(1, C).value = "Alamat SPK"
276                   C = C + 1
                      
277                   .Sheets("Sheet1").Cells(1, C).value = "Provinsi SPK"
278                   C = C + 1
                      
279                   .Sheets("Sheet1").Cells(1, C).value = "Kota SPK"
280                   C = C + 1
                      
281                   .Sheets("Sheet1").Cells(1, C).value = "Kecamatan SPK"
282                   C = C + 1
                      
283                   .Sheets("Sheet1").Cells(1, C).value = "Kodepos SPK"
284                   C = C + 1
285               End If
                  
286               If kolom(19).value = vbChecked Then
287                   .Sheets("Sheet1").Cells(1, C).value = "CDMA SPK"
288                   C = C + 1
                      
289                   .Sheets("Sheet1").Cells(1, C).value = "HP1 SPK"
290                   C = C + 1
                      
291                   .Sheets("Sheet1").Cells(1, C).value = "HP2 SPK"
292                   C = C + 1
293               End If
294               If kolom(20).value = vbChecked Then
295                   .Sheets("Sheet1").Cells(1, C).value = "Nama STNK"
296                   C = C + 1
297               End If
298               If kolom(21).value = vbChecked Then
299                   .Sheets("Sheet1").Cells(1, C).value = "Alamat STNK"
300                   C = C + 1
                      
301                   .Sheets("Sheet1").Cells(1, C).value = "Provinsi STNK"
302                   C = C + 1
                      
303                   .Sheets("Sheet1").Cells(1, C).value = "Kota STNK"
304                   C = C + 1
                      
305                   .Sheets("Sheet1").Cells(1, C).value = "Kecamatan STNK"
306                   C = C + 1
                      
307                   .Sheets("Sheet1").Cells(1, C).value = "Kodepos STNK"
308                   C = C + 1
309               End If
310           End If

              
311           .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
312           .Sheets("Sheet1").Range("1:1").Font.Bold = True
              
              
313           .Sheets("Sheet1").Select
314           .Sheets("Sheet1").Name = "Download Data Non Retention"
315           .Sheets("Download Data Non Retention").Select
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

316           .Close SaveChanges:=True, FileName:=onV
317       End With
318       Set oWB = Nothing
319       oApp.Quit
320       Set oApp = Nothing
End Sub


Private Sub cmdPrintOut_Click()
    Dim pr As Object
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

Private Sub cmdProsesVin_Click()
    If Month(tglAwalVin) <> Month(tglAkhirVin) Or year(tglAwalVin) <> year(tglAkhirVin) Then
        MsgBox "Range periode dalam 1 bulan", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If opTipe(7).value = True Then
        Call ambilLaporanVinRetentionTglInvoice
    ElseIf opTipe(6).value = True Then
        'Call ambilLaporanVinRetentionTglRetail
    ElseIf opTipe(8).value = True Then
        'Call ambilLaporanVinRetentionTglDelivery
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
            If semua Then
                .Fields.Append "ALAMAT", adVarChar, 300
            End If
            .Fields.Append "KELURAHAN", adVarChar, 100
            .Fields.Append "KECAMATAN", adVarChar, 100
            .Fields.Append "KOTA", adVarChar, 100
            .Fields.Append "PROPINSI", adVarChar, 100
            .Fields.Append "KODEPOS", adVarChar, 100
        End If
        
        If kolom(10).value = vbChecked Then
            .Fields.Append "PTSN", adVarChar, 100
            .Fields.Append "HP1", adVarChar, 100
            .Fields.Append "HP2", adVarChar, 100
            .Fields.Append "USENAME", adVarChar, 200
            .Fields.Append "USEHP", adVarChar, 100
        End If
        
        If kolom(11).value = vbChecked Then
            .Fields.Append "k1", adVarChar, 100
            .Fields.Append "dealerK1", adVarChar, 100
            
            .Fields.Append "k5", adVarChar, 100
            .Fields.Append "dealerK5", adVarChar, 100
            
            .Fields.Append "k10", adVarChar, 100
            .Fields.Append "dealerK10", adVarChar, 100
            
            .Fields.Append "k15", adVarChar, 100
            .Fields.Append "dealerK15", adVarChar, 100
            
            .Fields.Append "k20", adVarChar, 100
            .Fields.Append "dealerK20", adVarChar, 100
            
            .Fields.Append "k25", adVarChar, 100
            .Fields.Append "dealerK25", adVarChar, 100
            
            .Fields.Append "k30", adVarChar, 100
            .Fields.Append "dealerK30", adVarChar, 100
            
            .Fields.Append "k35", adVarChar, 100
            .Fields.Append "dealerK35", adVarChar, 100
            
            .Fields.Append "k40", adVarChar, 100
            .Fields.Append "dealerK40", adVarChar, 100
            
            .Fields.Append "k45", adVarChar, 100
            .Fields.Append "dealerK45", adVarChar, 100
            
            .Fields.Append "k50", adVarChar, 100
            .Fields.Append "dealerK50", adVarChar, 100
            
            .Fields.Append "k55", adVarChar, 100
            .Fields.Append "dealerK55", adVarChar, 100
            
            .Fields.Append "k60", adVarChar, 100
            .Fields.Append "dealerK60", adVarChar, 100
            
            .Fields.Append "k65", adVarChar, 100
            .Fields.Append "dealerK65", adVarChar, 100
            
            .Fields.Append "k70", adVarChar, 100
            .Fields.Append "dealerK70", adVarChar, 100
            
            .Fields.Append "k75", adVarChar, 100
            .Fields.Append "dealerK75", adVarChar, 100
            
            .Fields.Append "k80", adVarChar, 100
            .Fields.Append "dealerK80", adVarChar, 100
            
            .Fields.Append "k85", adVarChar, 100
            .Fields.Append "dealerK85", adVarChar, 100
            
            .Fields.Append "k90", adVarChar, 100
            .Fields.Append "dealerK90", adVarChar, 100
            
            .Fields.Append "k95", adVarChar, 100
            .Fields.Append "dealerK95", adVarChar, 100
            
            .Fields.Append "k100", adVarChar, 100
            .Fields.Append "dealerK100", adVarChar, 100
            
            .Fields.Append "k105", adVarChar, 100
            .Fields.Append "dealerK105", adVarChar, 100
            
            .Fields.Append "k110", adVarChar, 100
            .Fields.Append "dealerK110", adVarChar, 100
            
            .Fields.Append "k115", adVarChar, 100
            .Fields.Append "dealerK115", adVarChar, 100
            
            .Fields.Append "k120", adVarChar, 100
            .Fields.Append "dealerK120", adVarChar, 100
            
            .Fields.Append "k125", adVarChar, 100
            .Fields.Append "dealerK125", adVarChar, 100
            
            .Fields.Append "k130", adVarChar, 100
            .Fields.Append "dealerK130", adVarChar, 100
            
            .Fields.Append "k135", adVarChar, 100
            .Fields.Append "dealerK135", adVarChar, 100
            
            .Fields.Append "k140", adVarChar, 100
            .Fields.Append "dealerK140", adVarChar, 100
            
            .Fields.Append "k145", adVarChar, 100
            .Fields.Append "dealerK145", adVarChar, 100
            
            .Fields.Append "k150", adVarChar, 100
            .Fields.Append "dealerK150", adVarChar, 100
            
            .Fields.Append "k155", adVarChar, 100
            .Fields.Append "dealerK155", adVarChar, 100
            
            .Fields.Append "k160", adVarChar, 100
            .Fields.Append "dealerK160", adVarChar, 100
            
            .Fields.Append "k165", adVarChar, 100
            .Fields.Append "dealerK165", adVarChar, 100
            
            .Fields.Append "k170", adVarChar, 100
            .Fields.Append "dealerK170", adVarChar, 100
            
            .Fields.Append "k175", adVarChar, 100
            .Fields.Append "dealerK175", adVarChar, 100
            
            .Fields.Append "k180", adVarChar, 100
            .Fields.Append "dealerK180", adVarChar, 100
            
            .Fields.Append "k185", adVarChar, 100
            .Fields.Append "dealerK185", adVarChar, 100
            
            .Fields.Append "k190", adVarChar, 100
            .Fields.Append "dealerK190", adVarChar, 100
            
            .Fields.Append "k195", adVarChar, 100
            .Fields.Append "dealerK195", adVarChar, 100
            
            .Fields.Append "k200", adVarChar, 100
            .Fields.Append "dealerK200", adVarChar, 100
            
            .Fields.Append "TIPELASTSB", adVarChar, 100
            .Fields.Append "LASTSVCSB", adVarChar, 100
            .Fields.Append "UMURLASTSB", adVarChar, 100
            .Fields.Append "DEALERLASTSB", adVarChar, 100
            .Fields.Append "SALAST", adVarChar, 100
        
        End If
        If kolom(12).value = vbChecked Then
            .Fields.Append "NEXTSB", adVarChar, 100
            .Fields.Append "TGLNEXTSB", adVarChar, 100
        End If
        
        
        If kolom(13).value = vbChecked Then
            .Fields.Append "DEALERPENJUAL", adVarChar, 200
        End If
        
        
        If kolom(14).value = vbChecked Then
            .Fields.Append "DEALERTERAKHIRSB", adVarChar, 100
        End If
        If kolom(15).value = vbChecked Then
            .Fields.Append "WIRANIAGA", adVarChar, 100
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
                .Fields.Append "PROPINSISPK", adVarChar, 100
                .Fields.Append "KOTASPK", adVarChar, 100
                .Fields.Append "KECSPK", adVarChar, 100
                .Fields.Append "KODEPOSSPK", adVarChar, 100
            End If
            
            If kolom(19).value = vbChecked Then
                .Fields.Append "CDMASPK", adVarChar, 200
                .Fields.Append "HP1SPK", adVarChar, 100
                .Fields.Append "HP2SPK", adVarChar, 100
            End If
            If kolom(20).value = vbChecked Then
                .Fields.Append "NAMASTNK", adVarChar, 100
            End If
            If kolom(21).value = vbChecked Then
                .Fields.Append "ALAMATSTNK", adVarChar, 100
                .Fields.Append "PROPINSISTNK", adVarChar, 100
                .Fields.Append "KOTASTNK", adVarChar, 100
                .Fields.Append "KECSTNK", adVarChar, 100
                .Fields.Append "KODEPOSSTNK", adVarChar, 100
            End If
        End If
        
        
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

Private Sub ambilLaporanVinRetentionTglInvoice()
          Dim rsModel                 As ADODB.Recordset
          Dim rsSBE                   As ADODB.Recordset
          Dim rsDemand(10)            As ADODB.Recordset
          Dim rsSBETotal              As ADODB.Recordset
          Dim nilaiTemp               As Double
          Dim nilaiTempDemand(10)     As Double
          Dim nilaiTempDemandAll(10)  As Double
          Dim nilaiTempTotal(10)      As Double
          
1         If Log.RPLHistori = 0 Then
2             MsgBox "Proses dibatalkan karena flag replikasi belum aktif.", vbInformation
3             Exit Sub
4         End If
          
5         Set Report = New rptSBERateByModelVinRetention
6         Call createTableLaporanVin
          
          ''sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
7         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from [" & dbReplikasiHO & "]..model with (nolock) where FS_LOG_DEL = '' order by FS_GRP_MDL"
8         Set rsModel = objO.FetchData(sSQL)
           
9         If Not rsModel.EOF Then
10            rsModel.MoveFirst
11            ProgVin.Max = rsModel.RecordCount + 1
12            ProgVin.value = 1
13            ProgVin.Visible = True
14            DoEvents
              
15            For i = 0 To 10
16                nilaiTotal1(i) = 0
17                nilaiTotal2(i) = 0
18                nilaiTotal3(i) = 0
19                nilaiTotal4(i) = 0
20                nilaiTotal5(i) = 0
21                nilaiTotal6(i) = 0
22                nilaiTotal7(i) = 0
23                nilaiTotal8(i) = 0
24            Next
              
25            Do While Not rsModel.EOF
                  
                  ''1.  SB Same Outlet by VIN
26                With rsLaporan
27                    .AddNew
28                    nilaiTemp = 0
29                    !K1 = getDataTotalSBEbyVin(DateAdd("M", -1, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -1, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R001", NulSID(rsModel!fs_grp_mdl), False)
30                    !K10 = getDataTotalSBEbyVin(DateAdd("M", -7, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -5, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R010", NulSID(rsModel!fs_grp_mdl), False)
31                    !K20 = getDataTotalSBEbyVin(DateAdd("M", -14, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -10, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R020", NulSID(rsModel!fs_grp_mdl), False)
32                    !K30 = getDataTotalSBEbyVin(DateAdd("M", -21, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -15, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R030", NulSID(rsModel!fs_grp_mdl), False)
33                    !K40 = getDataTotalSBEbyVin(DateAdd("M", -28, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -20, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R040", NulSID(rsModel!fs_grp_mdl), False)
34                    !K50 = getDataTotalSBEbyVin(DateAdd("M", -35, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -25, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R050", NulSID(rsModel!fs_grp_mdl), False)
35                    !K60 = getDataTotalSBEbyVin(DateAdd("M", -42, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -30, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R060", NulSID(rsModel!fs_grp_mdl), False)
36                    !K70 = getDataTotalSBEbyVin(DateAdd("M", -49, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -35, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R070", NulSID(rsModel!fs_grp_mdl), False)
37                    !K80 = getDataTotalSBEbyVin(DateAdd("M", -56, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -40, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R080", NulSID(rsModel!fs_grp_mdl), False)
                      
38                    !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
39                    !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
40                    !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
41                    !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                         
                      
42                    nilaiTotal1(0) = nilaiTotal1(0) + !K1
43                    nilaiTotal1(1) = nilaiTotal1(1) + !K10
44                    nilaiTotal1(2) = nilaiTotal1(2) + !K20
45                    nilaiTotal1(3) = nilaiTotal1(3) + !K30
46                    nilaiTotal1(4) = nilaiTotal1(4) + !K40
47                    nilaiTotal1(5) = nilaiTotal1(5) + !K50
48                    nilaiTotal1(6) = nilaiTotal1(6) + !K60
49                    nilaiTotal1(7) = nilaiTotal1(7) + !K70
50                    nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
51                    !judul = "SB Same Outlet By Vin"
52                    !model = Nul(rsModel.Fields(0).value)
53                    !Urut = 1
54                    .Update
55                End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-12-14 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
56                With rsLaporan
57                    .AddNew
                  
                      
58                    !K1 = getDataDemandSameOutlet(DateAdd("M", -1, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -1, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 1)
59                    !K10 = getDataDemandSameOutlet(DateAdd("M", -7, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -5, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 3)
60                    !K20 = getDataDemandSameOutlet(DateAdd("M", -14, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -10, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 5)
61                    !K30 = getDataDemandSameOutlet(DateAdd("M", -21, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -15, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 7)
62                    !K40 = getDataDemandSameOutlet(DateAdd("M", -28, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -20, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 9)
63                    !K50 = getDataDemandSameOutlet(DateAdd("M", -35, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -25, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 11)
64                    !K60 = getDataDemandSameOutlet(DateAdd("M", -42, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -30, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 13)
65                    !K70 = getDataDemandSameOutlet(DateAdd("M", -49, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -35, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 15)
66                    !K80 = getDataDemandSameOutlet(DateAdd("M", -56, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -40, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 17)
                      
67                    !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
68                    !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
69                    !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
70                    !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
71                    nilaiTotal2(0) = nilaiTotal2(0) + !K1
72                    nilaiTotal2(1) = nilaiTotal2(1) + !K10
73                    nilaiTotal2(2) = nilaiTotal2(2) + !K20
74                    nilaiTotal2(3) = nilaiTotal2(3) + !K30
75                    nilaiTotal2(4) = nilaiTotal2(4) + !K40
76                    nilaiTotal2(5) = nilaiTotal2(5) + !K50
77                    nilaiTotal2(6) = nilaiTotal2(6) + !K60
78                    nilaiTotal2(7) = nilaiTotal2(7) + !K70
79                    nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
80                    !judul = "Demand SB Same Outlet By Vin"
81                    !model = Nul(rsModel.Fields(0).value)
82                    !Urut = 2
83                    .Update
84                End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
85                For i = 0 To 10
86                    nilaiTempDemand(i) = 0
87                    nilaiTempTotal(i) = 0
88                Next
89                rsLaporan.filter = ""
90                If Not rsLaporan.EOF Then
91                    rsLaporan.filter = "judul = 'Demand SB Same Outlet By Vin' and model = '" & Nul(rsModel.Fields(0).value) & "'"
92                    If Not rsLaporan.EOF Then
93                        If rsLaporan!K1 > 0 Then
94                            nilaiTempDemand(0) = rsLaporan!K1
95                        End If
96                        If rsLaporan!K10 > 0 Then
97                            nilaiTempDemand(1) = rsLaporan!K10
98                        End If
99                        If rsLaporan!K20 > 0 Then
100                           nilaiTempDemand(2) = rsLaporan!K20
101                       End If
102                       If rsLaporan!K30 > 0 Then
103                           nilaiTempDemand(3) = rsLaporan!K30
104                       End If
105                       If rsLaporan!K40 > 0 Then
106                           nilaiTempDemand(4) = rsLaporan!K40
107                       End If
108                       If rsLaporan!K50 > 0 Then
109                           nilaiTempDemand(5) = rsLaporan!K50
110                       End If
111                       If rsLaporan!K60 > 0 Then
112                           nilaiTempDemand(6) = rsLaporan!K60
113                       End If
114                       If rsLaporan!K70 > 0 Then
115                           nilaiTempDemand(7) = rsLaporan!K70
116                       End If
117                       If rsLaporan!K80 > 0 Then
118                           nilaiTempDemand(8) = rsLaporan!K80
119                       End If
120                   End If
121               End If
122               rsLaporan.filter = ""
                  
123               If Not rsLaporan.EOF Then
124                   rsLaporan.filter = "judul = 'SB Same Outlet By Vin' and model = '" & Nul(rsModel.Fields(0).value) & "'"
125                   If Not rsLaporan.EOF Then
126                       If rsLaporan!K1 > 0 Then
127                           nilaiTempTotal(0) = rsLaporan!K1
128                       End If
129                       If rsLaporan!K10 > 0 Then
130                           nilaiTempTotal(1) = rsLaporan!K10
131                       End If
132                       If rsLaporan!K20 > 0 Then
133                           nilaiTempTotal(2) = rsLaporan!K20
134                       End If
135                       If rsLaporan!K30 > 0 Then
136                           nilaiTempTotal(3) = rsLaporan!K30
137                       End If
138                       If rsLaporan!K40 > 0 Then
139                           nilaiTempTotal(4) = rsLaporan!K40
140                       End If
141                       If rsLaporan!K50 > 0 Then
142                           nilaiTempTotal(5) = rsLaporan!K50
143                       End If
144                       If rsLaporan!K60 > 0 Then
145                           nilaiTempTotal(6) = rsLaporan!K60
146                       End If
147                       If rsLaporan!K70 > 0 Then
148                           nilaiTempTotal(7) = rsLaporan!K70
149                       End If
150                       If rsLaporan!K80 > 0 Then
151                           nilaiTempTotal(8) = rsLaporan!K80
152                       End If
153                   End If
154               End If
                  
155               With rsLaporan
156                   .AddNew
157                   If nilaiTempDemand(0) > 0 Then
158                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
159                   Else
160                       !K1 = 0 & " %"
161                   End If
162                   If nilaiTempDemand(1) > 0 Then
163                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
164                   Else
165                       !K10 = 0 & " %"
166                   End If
167                   If nilaiTempDemand(2) > 0 Then
168                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
169                   Else
170                       !K20 = 0 & " %"
171                   End If
172                   If nilaiTempDemand(3) > 0 Then
173                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
174                   Else
175                       !K30 = 0 & " %"
176                   End If
177                   If nilaiTempDemand(4) > 0 Then
178                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
179                   Else
180                       !K40 = 0 & " %"
181                   End If
182                   If nilaiTempDemand(5) > 0 Then
183                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
184                   Else
185                       !K50 = 0 & " %"
186                   End If
187                   If nilaiTempDemand(6) > 0 Then
188                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
189                   Else
190                       !K60 = 0 & " %"
191                   End If
192                   If nilaiTempDemand(7) > 0 Then
193                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
194                   Else
195                       !K70 = 0 & " %"
196                   End If
197                   If nilaiTempDemand(8) > 0 Then
198                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
199                   Else
200                       !K80 = 0 & " %"
201                   End If
                      
202                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
203                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
204                   Else
205                       !tot16 = 0 & " %"
206                   End If
                      
207                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
208                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
209                   Else
210                       !tot26 = 0 & " %"
211                   End If
                       
212                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
213                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
214                   Else
215                       !tot19 = 0 & " %"
216                   End If
                       
217                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
218                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
219                   Else
220                       !tot29 = 0 & " %"
221                   End If
                       
222                   !Urut = 3
223                   !judul = "Retention Rate Same Outlet"
224                   !model = Nul(rsModel.Fields(0).value)
225                   .Update
226               End With
                  
227               With rsLaporan
228                   .AddNew
                      
229                   !K1 = getDataTotalSBEbyVin(DateAdd("M", -1, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -1, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R001", NulSID(rsModel!fs_grp_mdl), True)
230                   !K10 = getDataTotalSBEbyVin(DateAdd("M", -7, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -5, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R010", NulSID(rsModel!fs_grp_mdl), True)
231                   !K20 = getDataTotalSBEbyVin(DateAdd("M", -14, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -10, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R020", NulSID(rsModel!fs_grp_mdl), True)
232                   !K30 = getDataTotalSBEbyVin(DateAdd("M", -21, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -15, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R030", NulSID(rsModel!fs_grp_mdl), True)
233                   !K40 = getDataTotalSBEbyVin(DateAdd("M", -28, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -20, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R040", NulSID(rsModel!fs_grp_mdl), True)
234                   !K50 = getDataTotalSBEbyVin(DateAdd("M", -35, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -25, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R050", NulSID(rsModel!fs_grp_mdl), True)
235                   !K60 = getDataTotalSBEbyVin(DateAdd("M", -42, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -30, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R060", NulSID(rsModel!fs_grp_mdl), True)
236                   !K70 = getDataTotalSBEbyVin(DateAdd("M", -49, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -35, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R070", NulSID(rsModel!fs_grp_mdl), True)
237                   !K80 = getDataTotalSBEbyVin(DateAdd("M", -56, Format(tglAwalVin.value, "dd/mm/yyyy")), DateAdd("M", -40, Format(tglAkhirVin.value, "dd/mm/yyyy")), "R080", NulSID(rsModel!fs_grp_mdl), True)
                      
                           
238                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
239                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
240                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
241                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                           
242                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
243                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
244                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
245                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
246                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
247                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
248                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
249                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
250                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
251                   !judul = "SB Same Group (NCM) By Vin"
252                   !model = Nul(rsModel.Fields(0).value)
253                   !Urut = 4
254                   .Update
255               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-12-14 || DEMAND SAME GROUP
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
256               For i = 0 To 10
257                   nilaiTempDemandAll(i) = 0
258               Next
259               With rsLaporan
260                   .AddNew
                  
                      
261                   !K1 = getDataDemandSameGroup(DateAdd("M", -1, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -1, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 1)
262                   nilaiTempDemandAll(0) = !K1
263                   !K10 = getDataDemandSameGroup(DateAdd("M", -7, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -5, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 3)
264                   nilaiTempDemandAll(1) = !K10
265                   !K20 = getDataDemandSameGroup(DateAdd("M", -14, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -10, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 5)
266                   nilaiTempDemandAll(2) = !K20
267                   !K30 = getDataDemandSameGroup(DateAdd("M", -21, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -15, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 7)
268                   nilaiTempDemandAll(3) = !K30
269                   !K40 = getDataDemandSameGroup(DateAdd("M", -28, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -20, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 9)
270                   nilaiTempDemandAll(4) = !K40
271                   !K50 = getDataDemandSameGroup(DateAdd("M", -35, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -25, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 11)
272                   nilaiTempDemandAll(5) = !K50
273                   !K60 = getDataDemandSameGroup(DateAdd("M", -42, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -30, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 13)
274                   nilaiTempDemandAll(6) = !K60
275                   !K70 = getDataDemandSameGroup(DateAdd("M", -49, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -35, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 15)
276                   nilaiTempDemandAll(7) = !K70
277                   !K80 = getDataDemandSameGroup(DateAdd("M", -56, Format(tglAwalVin, "dd/mm/yyyy")), DateAdd("M", -40, Format(tglAkhirVin, "dd/mm/yyyy")), NulSID(rsModel.Fields(0).value), 17)
278                   nilaiTempDemandAll(8) = !K80
                      
279                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
280                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
281                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
282                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
283                   nilaiTotal5(0) = nilaiTotal5(0) + !K1
284                   nilaiTotal5(1) = nilaiTotal5(1) + !K10
285                   nilaiTotal5(2) = nilaiTotal5(2) + !K20
286                   nilaiTotal5(3) = nilaiTotal5(3) + !K30
287                   nilaiTotal5(4) = nilaiTotal5(4) + !K40
288                   nilaiTotal5(5) = nilaiTotal5(5) + !K50
289                   nilaiTotal5(6) = nilaiTotal5(6) + !K60
290                   nilaiTotal5(7) = nilaiTotal5(7) + !K70
291                   nilaiTotal5(8) = nilaiTotal5(8) + !K80
                       
292                   !judul = "Demand SB Same Group By Vin"
293                   !model = Nul(rsModel.Fields(0).value)
294                   !Urut = 5
295                   .Update
296               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
297               For i = 0 To 10
298                   nilaiTempTotal(i) = 0
299               Next
300               rsLaporan.filter = ""
                  
301               If Not rsLaporan.EOF Then
302                   rsLaporan.filter = "judul = 'SB Same Group (NCM) By Vin' and model = '" & Nul(rsModel.Fields(0).value) & "'"
303                   If Not rsLaporan.EOF Then
304                       If rsLaporan!K1 > 0 Then
305                           nilaiTempTotal(0) = rsLaporan!K1
306                       End If
307                       If rsLaporan!K10 > 0 Then
308                           nilaiTempTotal(1) = rsLaporan!K10
309                       End If
310                       If rsLaporan!K20 > 0 Then
311                           nilaiTempTotal(2) = rsLaporan!K20
312                       End If
313                       If rsLaporan!K30 > 0 Then
314                           nilaiTempTotal(3) = rsLaporan!K30
315                       End If
316                       If rsLaporan!K40 > 0 Then
317                           nilaiTempTotal(4) = rsLaporan!K40
318                       End If
319                       If rsLaporan!K50 > 0 Then
320                           nilaiTempTotal(5) = rsLaporan!K50
321                       End If
322                       If rsLaporan!K60 > 0 Then
323                           nilaiTempTotal(6) = rsLaporan!K60
324                       End If
325                       If rsLaporan!K70 > 0 Then
326                           nilaiTempTotal(7) = rsLaporan!K70
327                       End If
328                       If rsLaporan!K80 > 0 Then
329                           nilaiTempTotal(8) = rsLaporan!K80
330                       End If
331                   End If
332               End If
                  
333               With rsLaporan
334                   .AddNew
335                   If nilaiTempDemandAll(0) > 0 Then
336                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemandAll(0)) * 100, 0) & " %"
337                   Else
338                       !K1 = 0 & " %"
339                   End If
340                   If nilaiTempDemandAll(1) > 0 Then
341                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemandAll(1)) * 100, 0) & " %"
342                   Else
343                       !K10 = 0 & " %"
344                   End If
345                   If nilaiTempDemandAll(2) > 0 Then
346                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemandAll(2)) * 100, 0) & " %"
347                   Else
348                       !K20 = 0 & " %"
349                   End If
350                   If nilaiTempDemandAll(3) > 0 Then
351                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemandAll(3)) * 100, 0) & " %"
352                   Else
353                       !K30 = 0 & " %"
354                   End If
355                   If nilaiTempDemandAll(4) > 0 Then
356                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemandAll(4)) * 100, 0) & " %"
357                   Else
358                       !K40 = 0 & " %"
359                   End If
360                   If nilaiTempDemandAll(5) > 0 Then
361                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemandAll(5)) * 100, 0) & " %"
362                   Else
363                       !K50 = 0 & " %"
364                   End If
365                   If nilaiTempDemandAll(6) > 0 Then
366                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemandAll(6)) * 100, 0) & " %"
367                   Else
368                       !K60 = 0 & " %"
369                   End If
370                   If nilaiTempDemandAll(7) > 0 Then
371                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemandAll(7)) * 100, 0) & " %"
372                   Else
373                       !K70 = 0 & " %"
374                   End If
375                   If nilaiTempDemandAll(8) > 0 Then
376                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemandAll(8)) * 100, 0) & " %"
377                   Else
378                       !K80 = 0 & " %"
379                   End If
                      
                      
380                   If nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) > 0 Then
381                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5))) * 100, 0) & " %"
382                   Else
383                       !tot16 = 0 & " %"
384                   End If
                      
385                   If nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) > 0 Then
386                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5))) * 100, 0) & " %"
387                   Else
388                       !tot26 = 0 & " %"
389                   End If
                       
390                   If nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8) > 0 Then
391                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8))) * 100, 0) & " %"
392                   Else
393                       !tot19 = 0 & " %"
394                   End If
                       
395                   If nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8) > 0 Then
396                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8))) * 100, 0) & " %"
397                   Else
398                       !tot29 = 0 & " %"
399                   End If
                      
400                   !Urut = 6
401                   !judul = "Retention Rate Same Group (NCM) By Vin"
402                   !model = Nul(rsModel.Fields(0).value)
403                   .Update
404               End With
                   
405               With rsLaporan
406                   .AddNew
                      
407                   !K1 = getDataTotalSBEAll("R001", NulSID(rsModel!fs_grp_mdl))
408                   !K10 = getDataTotalSBEAll("R010", NulSID(rsModel!fs_grp_mdl))
409                   !K20 = getDataTotalSBEAll("R020", NulSID(rsModel!fs_grp_mdl))
410                   !K30 = getDataTotalSBEAll("R030", NulSID(rsModel!fs_grp_mdl))
411                   !K40 = getDataTotalSBEAll("R040", NulSID(rsModel!fs_grp_mdl))
412                   !K50 = getDataTotalSBEAll("R050", NulSID(rsModel!fs_grp_mdl))
413                   !K60 = getDataTotalSBEAll("R060", NulSID(rsModel!fs_grp_mdl))
414                   !K70 = getDataTotalSBEAll("R070", NulSID(rsModel!fs_grp_mdl))
415                   !K80 = getDataTotalSBEAll("R080", NulSID(rsModel!fs_grp_mdl))
                      
416                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
417                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
418                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
419                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
420                   nilaiTotal6(0) = nilaiTotal6(0) + !K1
421                   nilaiTotal6(1) = nilaiTotal6(1) + !K10
422                   nilaiTotal6(2) = nilaiTotal6(2) + !K20
423                   nilaiTotal6(3) = nilaiTotal6(3) + !K30
424                   nilaiTotal6(4) = nilaiTotal6(4) + !K40
425                   nilaiTotal6(5) = nilaiTotal6(5) + !K50
426                   nilaiTotal6(6) = nilaiTotal6(6) + !K60
427                   nilaiTotal6(7) = nilaiTotal6(7) + !K70
428                   nilaiTotal6(8) = nilaiTotal6(8) + !K80
                      
429                   !judul = "SB Total (NCM & Non NCM)"
430                   !model = Nul(rsModel.Fields(0).value)
431                   !Urut = 7
432                   .Update
433               End With
434               rsLaporan.filter = ""
                  
                  
435               For i = 0 To 10
436                   nilaiTempTotal(i) = 0
437               Next
438               rsLaporan.filter = ""
                  
439               If Not rsLaporan.EOF Then
440                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
441                   If Not rsLaporan.EOF Then
442                       If rsLaporan!K1 > 0 Then
443                           nilaiTempTotal(0) = rsLaporan!K1
444                       End If
445                       If rsLaporan!K10 > 0 Then
446                           nilaiTempTotal(1) = rsLaporan!K10
447                       End If
448                       If rsLaporan!K20 > 0 Then
449                           nilaiTempTotal(2) = rsLaporan!K20
450                       End If
451                       If rsLaporan!K30 > 0 Then
452                           nilaiTempTotal(3) = rsLaporan!K30
453                       End If
454                       If rsLaporan!K40 > 0 Then
455                           nilaiTempTotal(4) = rsLaporan!K40
456                       End If
457                       If rsLaporan!K50 > 0 Then
458                           nilaiTempTotal(5) = rsLaporan!K50
459                       End If
460                       If rsLaporan!K60 > 0 Then
461                           nilaiTempTotal(6) = rsLaporan!K60
462                       End If
463                       If rsLaporan!K70 > 0 Then
464                           nilaiTempTotal(7) = rsLaporan!K70
465                       End If
466                       If rsLaporan!K80 > 0 Then
467                           nilaiTempTotal(8) = rsLaporan!K80
468                       End If
469                   End If
470               End If
                  
471               With rsLaporan
472                   .AddNew
473                   If nilaiTempDemandAll(0) > 0 Then
474                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemandAll(0)) * 100, 0) & " %"
475                   Else
476                       !K1 = 0 & " %"
477                   End If
478                   If nilaiTempDemandAll(1) > 0 Then
479                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemandAll(1)) * 100, 0) & " %"
480                   Else
481                       !K10 = 0 & " %"
482                   End If
483                   If nilaiTempDemandAll(2) > 0 Then
484                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemandAll(2)) * 100, 0) & " %"
485                   Else
486                       !K20 = 0 & " %"
487                   End If
488                   If nilaiTempDemandAll(3) > 0 Then
489                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemandAll(3)) * 100, 0) & " %"
490                   Else
491                       !K30 = 0 & " %"
492                   End If
493                   If nilaiTempDemandAll(4) > 0 Then
494                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemandAll(4)) * 100, 0) & " %"
495                   Else
496                       !K40 = 0 & " %"
497                   End If
498                   If nilaiTempDemandAll(5) > 0 Then
499                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemandAll(5)) * 100, 0) & " %"
500                   Else
501                       !K50 = 0 & " %"
502                   End If
503                   If nilaiTempDemandAll(6) > 0 Then
504                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemandAll(6)) * 100, 0) & " %"
505                   Else
506                       !K60 = 0 & " %"
507                   End If
508                   If nilaiTempDemandAll(7) > 0 Then
509                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemandAll(7)) * 100, 0) & " %"
510                   Else
511                       !K70 = 0 & " %"
512                   End If
513                   If nilaiTempDemandAll(8) > 0 Then
514                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemandAll(8)) * 100, 0) & " %"
515                   Else
516                       !K80 = 0 & " %"
517                   End If
                      
                      
518                   If nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) > 0 Then
519                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5))) * 100, 0) & " %"
520                   Else
521                       !tot16 = 0 & " %"
522                   End If
                      
523                   If nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) > 0 Then
524                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5))) * 100, 0) & " %"
525                   Else
526                       !tot26 = 0 & " %"
527                   End If
                       
528                   If nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8) > 0 Then
529                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemandAll(0) + nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8))) * 100, 0) & " %"
530                   Else
531                       !tot19 = 0 & " %"
532                   End If
                       
533                   If nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8) > 0 Then
534                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemandAll(1) + nilaiTempDemandAll(2) + nilaiTempDemandAll(3) + nilaiTempDemandAll(4) + nilaiTempDemandAll(5) + nilaiTempDemandAll(6) + nilaiTempDemandAll(7) + nilaiTempDemandAll(8))) * 100, 0) & " %"
535                   Else
536                       !tot29 = 0 & " %"
537                   End If
                      
538                   !Urut = 8
539                   !judul = "Retention Rate Total (NCM & Non NCM)"
540                   !model = Nul(rsModel.Fields(0).value)
541                   .Update
542               End With
543               rsLaporan.filter = ""
544               rsModel.MoveNext
545               ProgVin.value = ProgVin.value + 1
546               DoEvents
547           Loop
548       Else
549           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
550           Exit Sub
551       End If
          
552       Call addTotalRow
          
553       Report.database.SetDataSource rsLaporan
554       ProgVin.Visible = False
555       Report.txtHeaderLaporan.SetText "Laporan Retention " & gNmCbg
556       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
557       With CRViewer1
558           .ReportSource = Report
559           .ViewReport
560           .Zoom 100
561           .Visible = True
562           .Left = 300
563           .Height = Me.Height - 1200
564           .Width = Me.Width - 750
565       End With
         
566       cmdTutup.Visible = True
567       cmdTutup.Left = CRViewer1.Left + 20
568       cmdTutup.Top = CRViewer1.Top + 400
569       Screen.MousePointer = vbDefault
          
570       Call DisableExport
571       Call setLokasiTombol
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
    Dim pr As Object
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

Private Function getDataTotalSBEAll(FSRC As String, fs_grp_mdl As String) As Double
    sSQL = "select COUNT(noinv) total"
    sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
    sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
    sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
    sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
    sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
    sSQL = sSQL & " and t.FS_RC = '" & FSRC & "'"
    sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0 and e.fs_grp_mdl = '" & fs_grp_mdl & "'"
    getDataTotalSBEAll = objO.FetchData_Double(sSQL)
End Function

Private Function getDataTotalSBEbyVin(tanggalRetailAwal As Date, tanggalRetailAkhir As Date, FSRC As String, fs_grp_mdl As String, isPenjualanAll As Boolean) As Double
    sSQL = "select COUNT(noinv) total"
    sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK)"
    sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' AND (s.tanggal between '" & Format(tanggalRetailAwal, "yyyy-mm-dd") & " 00:00:00' and '" & Format(tanggalRetailAkhir, "yyyy-mm-dd") & " 23:59:59') " & IIf(isPenjualanAll, "", " and s.kodedealer = '" & Log.kodeCabang & "'")
    sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
    sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.FS_KD_DEALER and d.FS_LOG_DEL = ''"
    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
    sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
    sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
    sSQL = sSQL & " and t.FS_RC = '" & FSRC & "'"
    sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0 and e.fs_grp_mdl = '" & fs_grp_mdl & "'"
    getDataTotalSBEbyVin = objO.FetchData_Double(sSQL)
End Function

Private Function getDataDemandSameOutlet(tanggalRetailAwal As Date, tanggalRetailAkhir As Date, model As String, pembagi As Integer) As Double
    sSQL = "select round(isnull(sum(a.total),0)/" & IIf(pembagi > 1, pembagi, 1) & ",0) total,isnull(b.fs_grp_mdl,'') model "
    ''20210624 Pertimbangan : ditemukan ada beberapa pembagian yang nilainya dibulatkan turun oleh SQL
    sSQL = "select round(isnull(cast(sum(a.total) as decimal),0)/" & IIf(pembagi > 1, pembagi, 1) & ",0)  total"
    sSQL = sSQL & " from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kode_type "
    sSQL = sSQL & " from Sihdr a "
    sSQL = sSQL & " INNER JOIN spk b ON a.nospk = b.SPK_No and b.SPK_Status = 1 "
    sSQL = sSQL & " where a.TglFaktur between '" & Format(tanggalRetailAwal, "yyyy-mm-dd") & "' and '" & Format(tanggalRetailAkhir, "yyyy-mm-dd") & "' "
    sSQL = sSQL & " group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kode_type ) a "
    ''sSQL = sSQL & " left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.kode_type,CHARINDEX(' ',Kode_Type)) and b.FS_LOG_DEL = '' "
    ''20210625 Pertimbangan : jika di retail ada suffix maka diabaikan suffixnya dikarenakan master di bengkel seringkali tanpa suffix sedangkan kategory modelnya sama.
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..model b on b.mdl = case when CHARINDEX(' ',a.kode_type) > 0 then left(a.kode_type,CHARINDEX(' ',a.kode_type)) else a.kode_type end and b.FS_LOG_DEL = '' "
    sSQL = sSQL & " where isnull(b.fs_grp_mdl,'') = '" & NulSID(model) & "' group by b.fs_grp_mdl"
    getDataDemandSameOutlet = objO.FetchDataSPK_Double(sSQL)
End Function

Private Function getDataDemandSameGroup(tanggalRetailAwal As Date, tanggalRetailAkhir As Date, model As String, pembagi As Integer) As Double
    sSQL = "select round(COUNT(norangka)/" & IIf(pembagi > 1, pembagi, 1) & ",0) total"
    ''20210624 Pertimbangan : ditemukan ada beberapa pembagian yang nilainya dibulatkan turun oleh SQL
    sSQL = "select round(isnull(cast(COUNT(norangka) as decimal),0)/" & IIf(pembagi > 1, pembagi, 1) & ",0)  total"
    sSQL = sSQL & " from [" & dbReplikasiHO & "]..Summary a "
    ''sSQL = sSQL & " left join [" & dbReplikasiHO & "]..model b on b.mdl = left(a.model,CHARINDEX(' ',a.model)) and b.FS_LOG_DEL = '' "
    ''20210625 Pertimbangan : jika di retail ada suffix maka diabaikan suffixnya dikarenakan master di bengkel seringkali tanpa suffix sedangkan kategory modelnya sama.
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..model b on b.mdl = case when CHARINDEX(' ',a.model) > 0 then left(a.model,CHARINDEX(' ',a.model)) else a.Model end and b.FS_LOG_DEL = '' "
    sSQL = sSQL & " where a.Status = 1 and a.TipeTr  = 'nd' and isnull(b.fs_grp_mdl,'') = '" & NulSID(model) & "'"
    sSQL = sSQL & " and a.tanggal between '" & Format(tanggalRetailAwal, "yyyy-mm-dd") & " 00:00:00' and '" & Format(tanggalRetailAkhir, "yyyy-mm-dd") & " 23:59:59' "
    
     sSQL = "select round(COUNT(norangka)/" & IIf(pembagi > 1, pembagi, 1) & ",0) total from ("
    ''20210624 Pertimbangan : ditemukan ada beberapa pembagian yang nilainya dibulatkan turun oleh SQL
    sSQL = "select round(isnull(cast(COUNT(norangka) as decimal),0)/" & IIf(pembagi > 1, pembagi, 1) & ",0)  total from ("
    sSQL = sSQL & " select a.NoRangka, MAX(a.TglEntry) Tanggal from [" & dbReplikasiHO & "]..Summary a "
    ''sSQL = sSQL & " left join [BENGKEL99]..model b on b.mdl = left(a.model,CHARINDEX(' ',a.model)) and b.FS_LOG_DEL = '' "
    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..model b on b.mdl = case when CHARINDEX(' ',a.model) > 0 then left(a.model,CHARINDEX(' ',a.model)) else a.Model end and b.FS_LOG_DEL = '' "
    sSQL = sSQL & " where a.TipeTr  = 'nd' and isnull(b.fs_grp_mdl,'') = '" & NulSID(model) & "' group by a.NoRangka"
    sSQL = sSQL & " )z where z.tanggal between '" & Format(tanggalRetailAwal, "yyyy-mm-dd") & " 00:00:00' and '" & Format(tanggalRetailAkhir, "yyyy-mm-dd") & " 23:59:59' "
    
    getDataDemandSameGroup = objO.FetchData_Double(sSQL)
End Function

Private Sub addTotalRow()
1         With rsLaporan
2             .AddNew
3             !K1 = nilaiTotal1(0)
4             !K10 = nilaiTotal1(1)
5             !K20 = nilaiTotal1(2)
6             !K30 = nilaiTotal1(3)
7             !K40 = nilaiTotal1(4)
8             !K50 = nilaiTotal1(5)
9             !K60 = nilaiTotal1(6)
10            !K70 = nilaiTotal1(7)
11            !K80 = nilaiTotal1(8)
              
12            !tot16 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
13            !tot19 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
14            !tot26 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
15            !tot29 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
                       
16            !judul = "SB Same Outlet By Vin"
17            !model = "Total"
18            !Urut = 1
19            .Update
20        End With
          
21        With rsLaporan
22            .AddNew
23            !K1 = nilaiTotal2(0)
24            !K10 = nilaiTotal2(1)
25            !K20 = nilaiTotal2(2)
26            !K30 = nilaiTotal2(3)
27            !K40 = nilaiTotal2(4)
28            !K50 = nilaiTotal2(5)
29            !K60 = nilaiTotal2(6)
30            !K70 = nilaiTotal2(7)
31            !K80 = nilaiTotal2(8)
              
32            !tot16 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
33            !tot19 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
34            !tot26 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
35            !tot29 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
                       
36            !judul = "Demand SB Same Outlet By Vin"
37            !model = "Total"
38            !Urut = 2
39            .Update
40        End With
          
41        With rsLaporan
42            .AddNew
43            If nilaiTotal2(0) > 0 Then
44                !K1 = Round(nilaiTotal1(0) * 100 / nilaiTotal2(0), 0) & " %"
45            Else
46                !K1 = "0 %"
47            End If
              
48            If nilaiTotal2(1) > 0 Then
49                !K10 = Round(nilaiTotal1(1) * 100 / nilaiTotal2(1), 0) & " %"
50            Else
51                !K10 = "0 %"
52            End If
                      
53            If nilaiTotal2(2) > 0 Then
54                !K20 = Round(nilaiTotal1(2) * 100 / nilaiTotal2(2), 0) & " %"
55            Else
56                !K20 = "0 %"
57            End If
              
58            If nilaiTotal2(3) > 0 Then
59                !K30 = Round(nilaiTotal1(3) * 100 / nilaiTotal2(3), 0) & " %"
60            Else
61                !K30 = "0 %"
62            End If
                       
63            If nilaiTotal2(4) > 0 Then
64                !K40 = Round(nilaiTotal1(4) * 100 / nilaiTotal2(4), 0) & " %"
65            Else
66                !K40 = "0 %"
67            End If

68            If nilaiTotal2(5) > 0 Then
69                !K50 = Round(nilaiTotal1(5) * 100 / nilaiTotal2(5), 0) & " %"
70            Else
71                !K50 = "0 %"
72            End If
              
73            If nilaiTotal2(6) > 0 Then
74                !K60 = Round(nilaiTotal1(6) * 100 / nilaiTotal2(6), 0) & " %"
75            Else
76                !K60 = "0 %"
77            End If
              
78            If nilaiTotal2(7) > 0 Then
79                !K70 = Round(nilaiTotal1(7) * 100 / nilaiTotal2(7), 0) & " %"
80            Else
81                !K70 = "0 %"
82            End If
              
83            If nilaiTotal2(8) > 0 Then
84                !K80 = Round(nilaiTotal1(8) * 100 / nilaiTotal2(8), 0) & " %"
85            Else
86                !K80 = "0 %"
87            End If
              
88            If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
89                !tot16 = Round(((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
90            Else
91                !tot16 = "0 %"
92            End If
              
93            If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
94                !tot19 = Round((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
95            Else
96                !tot19 = "0 %"
97            End If
              
98            If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
99                !tot26 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
100           Else
101               !tot26 = "0 %"
102           End If
              
103           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
104               !tot29 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
105           Else
106               !tot29 = "0 %"
107           End If
                     
108           !Urut = 3
109           !judul = "Retention Rate Same Outlet"
110           !model = "Total"
111           .Update
112       End With
          
113       With rsLaporan
114           .AddNew
115           !K1 = nilaiTotal4(0)
116           !K10 = nilaiTotal4(1)
117           !K20 = nilaiTotal4(2)
118           !K30 = nilaiTotal4(3)
119           !K40 = nilaiTotal4(4)
120           !K50 = nilaiTotal4(5)
121           !K60 = nilaiTotal4(6)
122           !K70 = nilaiTotal4(7)
123           !K80 = nilaiTotal4(8)
              
124           !tot16 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
125           !tot19 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
126           !tot26 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
127           !tot29 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
                   
128           !judul = "SB Same Group (NCM) By Vin"
129           !model = "Total"
130           !Urut = 4
131           .Update
132       End With
          
133       With rsLaporan
134           .AddNew
135           !K1 = nilaiTotal5(0)
136           !K10 = nilaiTotal5(1)
137           !K20 = nilaiTotal5(2)
138           !K30 = nilaiTotal5(3)
139           !K40 = nilaiTotal5(4)
140           !K50 = nilaiTotal5(5)
141           !K60 = nilaiTotal5(6)
142           !K70 = nilaiTotal5(7)
143           !K80 = nilaiTotal5(8)
              
144           !tot16 = nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5)
145           !tot19 = nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8)
146           !tot26 = nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5)
147           !tot29 = nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8)
                   
148           !judul = "Demand SB Same Group By Vin"
149           !model = "Total"
150           !Urut = 5
151           .Update
152       End With
          
153       With rsLaporan
154           .AddNew
155           If nilaiTotal5(0) > 0 Then
156               !K1 = Round(nilaiTotal4(0) * 100 / nilaiTotal5(0), 0) & " %"
157           Else
158               !K1 = "0 %"
159           End If
              
160           If nilaiTotal5(1) > 0 Then
161               !K10 = Round(nilaiTotal4(1) * 100 / nilaiTotal5(1), 0) & " %"
162           Else
163               !K10 = "0 %"
164           End If
                      
165           If nilaiTotal5(2) > 0 Then
166               !K20 = Round(nilaiTotal4(2) * 100 / nilaiTotal5(2), 0) & " %"
167           Else
168               !K20 = "0 %"
169           End If
              
170           If nilaiTotal5(3) > 0 Then
171               !K30 = Round(nilaiTotal4(3) * 100 / nilaiTotal5(3), 0) & " %"
172           Else
173               !K30 = "0 %"
174           End If
                       
175           If nilaiTotal5(4) > 0 Then
176               !K40 = Round(nilaiTotal4(4) * 100 / nilaiTotal5(4), 0) & " %"
177           Else
178               !K40 = "0 %"
179           End If

180           If nilaiTotal5(5) > 0 Then
181               !K50 = Round(nilaiTotal4(5) * 100 / nilaiTotal5(5), 0) & " %"
182           Else
183               !K50 = "0 %"
184           End If
              
185           If nilaiTotal5(6) > 0 Then
186               !K60 = Round(nilaiTotal4(6) * 100 / nilaiTotal5(6), 0) & " %"
187           Else
188               !K60 = "0 %"
189           End If
              
190           If nilaiTotal5(7) > 0 Then
191               !K70 = Round(nilaiTotal4(7) * 100 / nilaiTotal5(7), 0) & " %"
192           Else
193               !K70 = "0 %"
194           End If
              
195           If nilaiTotal5(8) > 0 Then
196               !K80 = Round(nilaiTotal4(8) * 100 / nilaiTotal5(8), 0) & " %"
197           Else
198               !K80 = "0 %"
199           End If
              
200           If nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) > 0 Then
201               !tot16 = Round(((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100) / (nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5)), 0) & " %"
202           Else
203               !tot16 = "0 %"
204           End If
              
205           If nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8) > 0 Then
206               !tot19 = Round((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8)), 0) & " %"
207           Else
208               !tot19 = "0 %"
209           End If
              
210           If nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) > 0 Then
211               !tot26 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100 / (nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5)), 0) & " %"
212           Else
213               !tot26 = "0 %"
214           End If
              
215           If nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8) > 0 Then
216               !tot29 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8)), 0) & " %"
217           Else
218               !tot29 = "0 %"
219           End If
              
220           !Urut = 6
221           !judul = "Retention Rate Same Group (NCM) By Vin"
222           !model = "Total"
223           .Update
224       End With
          
225       With rsLaporan
226           .AddNew
227           !K1 = nilaiTotal6(0)
228           !K10 = nilaiTotal6(1)
229           !K20 = nilaiTotal6(2)
230           !K30 = nilaiTotal6(3)
231           !K40 = nilaiTotal6(4)
232           !K50 = nilaiTotal6(5)
233           !K60 = nilaiTotal6(6)
234           !K70 = nilaiTotal6(7)
235           !K80 = nilaiTotal6(8)
              
236           !tot16 = nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)
237           !tot19 = nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)
238           !tot26 = nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)
239           !tot29 = nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)
                   
240           !judul = "SB Total (NCM & Non NCM)"
241           !model = "Total"
242           !Urut = 7
243           .Update
244       End With
          
245       With rsLaporan
246           .AddNew
247           If nilaiTotal5(0) > 0 Then
248               !K1 = Round(nilaiTotal6(0) * 100 / nilaiTotal5(0), 0) & " %"
249           Else
250               !K1 = "0 %"
251           End If
              
252           If nilaiTotal5(1) > 0 Then
253               !K10 = Round(nilaiTotal6(1) * 100 / nilaiTotal5(1), 0) & " %"
254           Else
255               !K10 = "0 %"
256           End If
                      
257           If nilaiTotal5(2) > 0 Then
258               !K20 = Round(nilaiTotal6(2) * 100 / nilaiTotal5(2), 0) & " %"
259           Else
260               !K20 = "0 %"
261           End If
              
262           If nilaiTotal5(3) > 0 Then
263               !K30 = Round(nilaiTotal6(3) * 100 / nilaiTotal5(3), 0) & " %"
264           Else
265               !K30 = "0 %"
266           End If
                       
267           If nilaiTotal5(4) > 0 Then
268               !K40 = Round(nilaiTotal6(4) * 100 / nilaiTotal5(4), 0) & " %"
269           Else
270               !K40 = "0 %"
271           End If

272           If nilaiTotal5(5) > 0 Then
273               !K50 = Round(nilaiTotal6(5) * 100 / nilaiTotal5(5), 0) & " %"
274           Else
275               !K50 = "0 %"
276           End If
              
277           If nilaiTotal5(6) > 0 Then
278               !K60 = Round(nilaiTotal6(6) * 100 / nilaiTotal5(6), 0) & " %"
279           Else
280               !K60 = "0 %"
281           End If
              
282           If nilaiTotal5(7) > 0 Then
283               !K70 = Round(nilaiTotal6(7) * 100 / nilaiTotal5(7), 0) & " %"
284           Else
285               !K70 = "0 %"
286           End If
              
287           If nilaiTotal5(8) > 0 Then
288               !K80 = Round(nilaiTotal6(8) * 100 / nilaiTotal5(8), 0) & " %"
289           Else
290               !K80 = "0 %"
291           End If
              
292           If nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) > 0 Then
293               !tot16 = Round(((nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)) * 100) / (nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5)), 0) & " %"
294           Else
295               !tot16 = "0 %"
296           End If
              
297           If nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8) > 0 Then
298               !tot19 = Round((nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)) * 100 / (nilaiTotal5(0) + nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8)), 0) & " %"
299           Else
300               !tot19 = "0 %"
301           End If
              
302           If nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) > 0 Then
303               !tot26 = Round((nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)) * 100 / (nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5)), 0) & " %"
304           Else
305               !tot26 = "0 %"
306           End If
              
307           If nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8) > 0 Then
308               !tot29 = Round((nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)) * 100 / (nilaiTotal5(1) + nilaiTotal5(2) + nilaiTotal5(3) + nilaiTotal5(4) + nilaiTotal5(5) + nilaiTotal5(6) + nilaiTotal5(7) + nilaiTotal5(8)), 0) & " %"
309           Else
310               !tot29 = "0 %"
311           End If
312           !Urut = 8
313           !judul = "Retention Rate Total (NCM & Non NCM)"
314           !model = "Total"
315           .Update
316       End With
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
    
    tglAwalVin = isnow
    tglAkhirVin = isnow
    
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
    
    ckKolom.value = vbChecked
    Call aturKolomLevel
End Sub


Public Function GetService(kriteria As String, Optional isHO As Boolean = False)
    rsCari.filter = ""
    rsCari.filter = "fs_rc = '" & kriteria & "'"
    If Not rsCari.EOF Then
        If isHO Then
            GetService = Format(rsCari("invdate"), "DD/MM/YYYY") '' & vbCrLf & Trim(rsCari("Fs_nm_dealer"))
        Else
            GetService = Format(rsCari("invdate"), "DD/MM/YYYY")
        End If
        
        Exit Function
    End If
    GetService = ""
End Function


Public Function GetServiceDealer(kriteria As String, Optional isHO As Boolean = False)
    rsCari.filter = ""
    rsCari.filter = "fs_rc = '" & kriteria & "'"
    If Not rsCari.EOF Then
        If isHO Then
            GetServiceDealer = Trim(rsCari("Fs_nm_dealer"))
         
        End If
        
        Exit Function
    End If
    GetServiceDealer = ""
End Function
