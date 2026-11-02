VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmInfoPelangganRecoveryMappingCall 
   Caption         =   ".:: Recovery Call Mapping ::."
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   2640
      Left            =   165
      TabIndex        =   49
      Top             =   705
      Width           =   11655
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbPilihKeterangan 
         Height          =   330
         Left            =   6120
         TabIndex        =   50
         Top             =   90
         Width           =   5415
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   9551
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   2400
         TabIndex        =   70
         Top             =   80
         Width           =   4215
         Begin VB.OptionButton OptTelp 
            Caption         =   "Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton OptTelp 
            Caption         =   "Tidak Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   71
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   66
         Top             =   2160
         Width           =   7455
         Begin VB.OptionButton OptSMS 
            Caption         =   "Belum Dikonfirmasi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   69
            Top             =   0
            Width           =   2415
         End
         Begin VB.OptionButton OptSMS 
            Caption         =   "Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton OptSMS 
            Caption         =   "Tidak Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   67
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   62
         Top             =   1800
         Width           =   7575
         Begin VB.OptionButton OptDM 
            Caption         =   "Tidak Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   65
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton OptDM 
            Caption         =   "Diterima"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton OptDM 
            Caption         =   "Belum Dikonfirmasi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   63
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   58
         Top             =   1080
         Width           =   7575
         Begin VB.OptionButton OptPel 
            Caption         =   "Belum Dikonfirmasi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   61
            Top             =   0
            Width           =   2655
         End
         Begin VB.OptionButton OptPel 
            Caption         =   "Tidak Akurat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   60
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton OptPel 
            Caption         =   "Akurat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   54
         Top             =   1440
         Width           =   7335
         Begin VB.OptionButton OptKen 
            Caption         =   "Belum Dikonfirmasi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   57
            Top             =   0
            Width           =   2415
         End
         Begin VB.OptionButton OptKen 
            Caption         =   "Sesuai"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton OptKen 
            Caption         =   "Tidak Sesuai"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   55
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.CheckBox ChUlang 
         Caption         =   "Konfirmasi Ulang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtKetTelp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6720
         MaxLength       =   100
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CheckBox chJual 
         Caption         =   "Unit Dijual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   51
         Top             =   600
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc AdodcKonfirmasi 
         Height          =   330
         Left            =   8925
         Top             =   630
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
      Begin MSComCtl2.DTPicker JamNewConfirm 
         Height          =   300
         Left            =   5520
         TabIndex        =   73
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111149058
         CurrentDate     =   39905
      End
      Begin TDBDate6Ctl.TDBDate TglNewKonfirm 
         Height          =   315
         Left            =   3360
         TabIndex        =   74
         Top             =   593
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmInfoPelangganRecoveryMappingCall.frx":0000
         Caption         =   "frmInfoPelangganRecoveryMappingCall.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPelangganRecoveryMappingCall.frx":0198
         Keys            =   "frmInfoPelangganRecoveryMappingCall.frx":01B6
         Spin            =   "frmInfoPelangganRecoveryMappingCall.frx":0214
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "dd/mm/yyyy"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   87
         Top             =   130
         Width           =   75
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirm Telp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   140
         Width           =   1245
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirm SMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   85
         Top             =   2220
         Width           =   1245
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   84
         Top             =   2205
         Width           =   75
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirm DM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   82
         Top             =   1845
         Width           =   75
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   2530
         Y2              =   2530
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   81
         Top             =   1485
         Width           =   75
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirm Pelanggan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   80
         Top             =   1140
         Width           =   1800
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   79
         Top             =   1140
         Width           =   75
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirm Kendaraan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   78
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   77
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label5 
         Caption         =   "Tanggal "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   76
         Top             =   623
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Jam"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   75
         Top             =   623
         Width           =   495
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame FrameGRP 
      BorderStyle     =   0  'None
      Caption         =   ".:: SELECTED ITEM REMAINDER::."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   150
      TabIndex        =   15
      Top             =   3345
      Width           =   11655
      Begin VB.TextBox txtPos 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         MaxLength       =   5
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtKota 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtCDMA 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         MaxLength       =   15
         TabIndex        =   24
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtTelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         MaxLength       =   15
         TabIndex        =   23
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtGSM 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         MaxLength       =   15
         TabIndex        =   22
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtNoChassis 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtNoPol 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNama 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox txtAlamat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   200
         TabIndex        =   18
         Top             =   1680
         Width           =   6015
      End
      Begin VB.TextBox txtNoEngine 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   17
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox ch 
         Caption         =   "Simpan data lama ke Log"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   16
         Top             =   990
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.Label Label40 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   ".:: SELECTED CUSTOMER AS USER ::."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3960
         TabIndex        =   48
         Top             =   0
         Width           =   4095
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   47
         Top             =   2040
         Width           =   75
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   46
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   45
         Top             =   2040
         Width           =   75
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2040
         Width           =   465
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   3840
         X2              =   8160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8520
         TabIndex        =   43
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9360
         TabIndex        =   42
         Top             =   2040
         Width           =   75
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8520
         TabIndex        =   41
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8520
         TabIndex        =   40
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9360
         TabIndex        =   39
         Top             =   1320
         Width           =   75
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9360
         TabIndex        =   38
         Top             =   1680
         Width           =   75
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Chassis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   36
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   35
         Top             =   1680
         Width           =   75
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   1320
         Width           =   75
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5160
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   30
         Top             =   600
         Width           =   75
      End
      Begin VB.Label lblMerk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Kendaraan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label74 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   75
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11520
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11520
         Y1              =   2460
         Y2              =   2460
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   5940
      Width           =   11655
      Begin VB.OptionButton OptApp 
         Caption         =   "Tidak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton OptApp 
         Caption         =   "Ya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox txtApp 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   149
         TabIndex        =   5
         Top             =   360
         Width           =   6015
      End
      Begin VB.CommandButton cmdBooking 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Booking"
         Enabled         =   0   'False
         Height          =   732
         Left            =   2520
         Picture         =   "frmInfoPelangganRecoveryMappingCall.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10680
         Picture         =   "frmInfoPelangganRecoveryMappingCall.frx":0891
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdRecord 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Record"
         Height          =   732
         Left            =   9960
         Picture         =   "frmInfoPelangganRecoveryMappingCall.frx":1021
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtPesan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   149
         TabIndex        =   1
         Top             =   960
         Width           =   7815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Appointment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   60
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   13
         Top             =   45
         Width           =   75
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   9840
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   9600
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   9600
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Promosikan Bengkel   -   Sistem Appointment (sarankan pelanggan untuk melakukan appointment)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   9615
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         X1              =   9600
         X2              =   120
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lCall1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   10080
         TabIndex        =   11
         Top             =   0
         Width           =   120
      End
      Begin VB.Label lCall2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10095
         TabIndex        =   10
         Top             =   15
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   75
      End
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM REMINDER RECOVERY MAPPING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1365
      TabIndex        =   88
      Top             =   195
      Width           =   10215
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   300
      X2              =   3765
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1365
      X2              =   3765
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2340
      X2              =   3765
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   11685
   End
End
Attribute VB_Name = "frmInfoPelangganRecoveryMappingCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TglNext As Date
'Private Remainder As String
Public noWO As String
Private Editing As Boolean
Private objR As New ClsRemainder
Private Rec As Boolean
Private rsRem As ADODB.Recordset
Private rsAll As ADODB.Recordset
Private rsCall As ADODB.Recordset

Private is1000 As Boolean
Private isNoPol As String

Private Nama As String
Private Alamat As String
Private KodePos As String
Private Kota As String
Private Propinsi As String
Private Rumah As String
Private GSM As String
Private CDMA As String
Private NoPolisi As String

Private kmLast As Double
Private TglLast As Date
Private tglJTBln As Date
Private flagTdkDiundang As Boolean
Public mdlKend As String


Public Property Let xTglNext(value As Date)
   TglNext = value
End Property

Public Property Let xRsRem(value As ADODB.Recordset)
   Set rsRem = value
End Property

Public Property Let xRsAll(value As ADODB.Recordset)
   Set rsAll = value
End Property

Public Property Let IsSvc1000(value As Boolean)
   is1000 = value
End Property

Public Property Let xNoPol(value As String)
   isNoPol = value
End Property

Private Sub ChUlang_Click()
   If ChUlang.value = 1 Then
      TglNewKonfirm.Enabled = True
      JamNewConfirm.Enabled = True
      TglNewKonfirm.Text = Format(isnow, "dd/mm/yyyy")
      JamNewConfirm.value = Now
   Else
      TglNewKonfirm.Enabled = False
      JamNewConfirm.Enabled = False
      TglNewKonfirm.Text = "__/__/____"
      JamNewConfirm.value = "0:00:00"
   End If
End Sub

 

Private Sub cmdBooking_Click()
Dim rsSearch As ADODB.Recordset

On Error Resume Next
    isBookByMrs = True
    If rsRem.Fields(0).Name = "NoSPK" Then
        rsAll.Find "SPK_No ='" & rsRem.Fields(0).value & "'", , adSearchForward, 1
    Else
        rsAll.Find "NoInv ='" & rsRem.Fields(0).value & "'", , adSearchForward, 1
    End If
    
    With frmBooking
        Dim nDpn1 As String
        Dim nTgh1 As String
        Dim nBlk1 As String
        Call pecahNoPol(txtNoPol.Text, nDpn1, nTgh1, nBlk1)
        .Text(0) = Trim(nDpn1)
        .Text(56) = Trim(nTgh1)
        .Text(57) = Trim(nBlk1)
                     
        .show vbModal
    End With
End Sub

 
 
Private Sub cmdExit_Click()
   If Rec = False Then
      If MsgBox("Data Remainder belum di'record !, yakin untuk keluar ?", vbCritical + vbYesNo, "Remainder Call") = vbYes Then
         Unload Me
      End If
   Else
      Unload Me
   End If
End Sub

 
 
 

Private Sub Form_Load()
   Rec = False
   Editing = False
   Call ListComboWaktu
   'Call BindToControl
   OptTelp(0).value = True
   Call ItemFormRecoveryRemainderCall
   Call Eta
End Sub

Private Sub ListComboWaktu()
   
End Sub
 
Private Sub OptApp_Click(Index As Integer)

      OptApp(0).ForeColor = vbBlack
      OptApp(1).ForeColor = vbBlack
      
   If OptApp(1).value = True Then
      cmdBooking.Enabled = True
      txtApp.Enabled = False
      OptApp(1).ForeColor = vbBlue
   ElseIf OptApp(0).value = True Then
      cmdBooking.Enabled = False
      txtApp.Enabled = True
      OptApp(0).ForeColor = vbBlue
   End If
End Sub

Private Sub OptDM_Click(Index As Integer)
Dim i As Integer
For i = 0 To 2
   OptDM(i).ForeColor = vbBlack
   If OptDM(i).value = True Then OptDM(i).ForeColor = vbBlue
Next i
End Sub

Private Sub OptKen_Click(Index As Integer)
Dim i As Integer
For i = 0 To 2
   OptKen(i).ForeColor = vbBlack
   If OptKen(i).value = True Then OptKen(i).ForeColor = vbBlue
Next i
End Sub

Private Sub OptPel_Click(Index As Integer)
Dim i As Integer
For i = 0 To 2
   OptPel(i).ForeColor = vbBlack
   If OptPel(i).value = True Then OptPel(i).ForeColor = vbBlue
Next i
End Sub

Private Sub OptSMS_Click(Index As Integer)
Dim i As Integer
For i = 0 To 2
   OptSMS(i).ForeColor = vbBlack
   If OptSMS(i).value = True Then OptSMS(i).ForeColor = vbBlue
Next i
End Sub

Private Sub OptTelp_Click(Index As Integer)
On Error Resume Next
      Dim i As Integer
OptTelp(0).ForeColor = vbBlack
OptTelp(1).ForeColor = vbBlack
With Me
   If OptTelp(0).value = True Then
        ''tidak diterima
        sSQL = "select keterangantipe as Tipe, pesan as keterangan,isRecovery FROM M_VoiceOfCustomer with (nolock) where tipe = 0"
        
      .OptTelp(0).ForeColor = vbBlue
      '.cmdEdit.Enabled = False
     ' .cmdCancel.Enabled = False
     ' .cmdSave.Enabled = False
      .txtKetTelp.Enabled = False
      .txtPesan.Enabled = False
      .txtKetTelp.SetFocus
      .chJual.Enabled = False
       
      For i = 0 To 2
         .OptPel(i).Enabled = False
         .OptKen(i).Enabled = False
         .OptDM(i).Enabled = False
         .OptSMS(i).Enabled = False
      Next i
      .OptApp(0).value = False
      .OptApp(0).Enabled = False
      .OptApp(1).value = False
      .OptApp(1).Enabled = False
      .txtApp.Enabled = False
      .cmdBooking.Enabled = False
      .cmdExit.Enabled = True
   ElseIf OptTelp(1).value = True Then
        ''diterima
        sSQL = "select keterangantipe as Tipe, pesan as keterangan,isRecovery FROM M_VoiceOfCustomer with (nolock) where tipe = 1"
        
      .OptTelp(1).ForeColor = vbBlue
    '  .cmdEdit.Enabled = False
    '  .cmdCancel.Enabled = False
     ' .cmdSave.Enabled = False
      .txtKetTelp.Enabled = False
      .txtKetTelp.Text = ""
      .chJual.Enabled = True
 
      For i = 0 To 2
         .OptPel(i).Enabled = True
         .OptKen(i).Enabled = True
         .OptDM(i).Enabled = True
         .OptSMS(i).Enabled = True
      Next i
      .OptApp(0).value = True
      .OptApp(0).Enabled = True
      .OptApp(1).value = False
      .OptApp(1).Enabled = True
      .txtApp.Enabled = True
      .txtPesan.Enabled = True
      .OptApp(0).ForeColor = vbBack
      .OptApp(1).ForeColor = vbBack
      .cmdBooking.Enabled = False
      .cmdExit.Enabled = True
   End If
   
    AdodcKonfirmasi.ConnectionString = SQLCONBengkel
    AdodcKonfirmasi.RecordSource = sSQL
    AdodcKonfirmasi.Refresh
    cbPilihKeterangan.DataSourceList = AdodcKonfirmasi
    cbPilihKeterangan.DataFieldList = "Keterangan"
    cbPilihKeterangan.Columns(0).Width = 2000
    cbPilihKeterangan.Columns(1).Width = 3000
    cbPilihKeterangan.Columns(2).Visible = False
    cbPilihKeterangan.Text = ""
    
   .txtNama.Enabled = False
   .txtAlamat.Enabled = False
   .txtKota.Enabled = False
   .txtPos.Enabled = False
   .txtTelp.Enabled = False
   .txtCDMA.Enabled = False
   .txtGSM.Enabled = False
   .txtNoPol.Enabled = False
End With
End Sub

Private Sub Eta()
'   Dim rsEta As ADODB.Recordset
'   Set rsEta = objR.FetchEta(noWO, "SM")
'   If rsEta.RecordCount = 0 Then Exit Sub
'   With Me
'      .txtHargaJasa = rsEta.Fields(0).value
'      .txtHargaPart = rsEta.Fields(1).value
'      .cboJam = Val(rsEta.Fields(2).value)
'      .cboMenit.text = rsEta.Fields(3).value
'   End With
End Sub

Private Sub ItemFormRecoveryRemainderCall()
'   Set rsCall = objR.FetchItemRemainderCall(noWO, txtNoChassis, NulA(txtKmNextSvc))
'   If rsCall.RecordCount = 0 Then
'      lCall1.Caption = "Call : 1"
'      lCall2.Caption = "Call : 1"
'      Exit Sub
'   End If
'   rsCall.MoveLast
'   lCall1.Caption = "Call : " & rsCall.RecordCount + 1
'   lCall2.Caption = "Call : " & rsCall.RecordCount + 1
'   Call BindingLastRemainder(rsCall)
End Sub

Private Sub BindingLastRemainder(rs As ADODB.Recordset)
 
End Sub

  

Private Sub cmdRecord_Click()
   If MsgBox("Record Data & " & lCall1.Caption & " ?", vbQuestion + vbYesNo) = vbYes Then
        If cbPilihKeterangan.value = "" Then
            MsgBox "Silahkan pilih keterangan konfirmasi Telp", vbInformation
            Exit Sub
        End If
'        If OptTelp(0).value Then
'            frmeTidakDiterima.Visible = True
'            Dim cRS As ADODB.Recordset
'            Set cRS = objR.FetchLastReminderCall(noWO)
'            If cRS.RecordCount <> 0 Then
'                If Trim(cRS.Fields(0).value) <> "" Then
'                    OptAlasan(CInt(cRS.Fields(0).value) - 1).value = True
'                    If CInt(cRS.Fields(0).value) = 4 Then
'                        txtAlasanLain.text = Trim(Replace(cRS.Fields(1).value, "4.", ""))
'                    End If
'                End If
'            End If
'        Else
'            Call simpanHasil
'        End If

        ''2017-09-27||update recovery reminder menjadi 1
        ''-----------------------------------------------
        sSQL = "update recoverymapping set isView = 1 where  norangka = '" & Trim(txtNoChassis.Text) & "'"
        Call objO.InsertData(sSQL)
        
        ''2017-08-15
        Call simpanHasil
   End If
End Sub

Private Sub simpanHasil()
'    Call objR.AddNew_LogRemainder_Call(noWO, Format(isnow, "dd/mm/yyyy"), DateDiff("d", Format(isnow, "dd/mm/yyyy"), _
'                           Format(CDate(txttglNextSvc.text), "dd/mm/yyyy")), txtNoChassis.text, _
'                           Format(CDate(txttglNextSvc.text), "dd/mm/yyyy"), Val(txtKmNextSvc.text), _
'                           IIf(OptTelp(1).value = True, 1, 0), txtKetTelp.text, IIf(ChUlang.value = 1, 1, 0), _
'                           Format(TglNewKonfirm.value, "dd/mm/yyyy"), Format(Me.JamNewConfirm.value, "hh:nn:ss"), _
'                           IIf(OptPel(1).value = True, 1, (IIf(OptPel(0).value = True, 0, 2))), _
'                           IIf(OptKen(1).value = True, 1, (IIf(OptKen(0).value = True, 0, 2))), _
'                           IIf(OptDM(1).value = True, 1, (IIf(OptDM(0).value = True, 0, 2))), _
'                           IIf(OptSMS(1).value = True, 1, (IIf(OptSMS(0).value = True, 0, 2))), _
'                           Val(Replace(txtHargaJasa, ".", "")), Val(Replace(txtHargaPart, ".", "")), Val(cboJam.text), Val(cboMenit.text), IIf(Editing = True, 1, 0), _
'                           IIf(OptApp(1).value = True, 1, 0), Trim(txtApp.text), Trim(txtPesan.text))
   
    Call objR.AddNew_LogRemainder_CallMapping(Format(isnow, "dd/mm/yyyy"), 1, txtNoChassis.Text, _
                           IIf(OptTelp(1).value = True, 1, 0), cbPilihKeterangan.value, IIf(ChUlang.value = 1, 1, 0), _
                           Format(TglNewKonfirm.value, "dd/mm/yyyy"), Format(Me.JamNewConfirm.value, "hh:nn:ss"), _
                           IIf(OptPel(1).value = True, 1, (IIf(OptPel(0).value = True, 0, 2))), _
                           IIf(OptKen(1).value = True, 1, (IIf(OptKen(0).value = True, 0, 2))), _
                           IIf(OptDM(1).value = True, 1, (IIf(OptDM(0).value = True, 0, 2))), _
                           IIf(OptSMS(1).value = True, 1, (IIf(OptSMS(0).value = True, 0, 2))), _
                           IIf(OptApp(1).value = True, 1, 0), Trim(txtApp.Text), Trim(txtPesan.Text))
   
      Rec = True
      
      If MsgBox("Keluar dari Remainder Recovery Call ?", vbQuestion + vbYesNo) = vbYes Then
         Unload Me
      End If
End Sub

Private Function estimasiWaktu(nKm As String, nTipeKend As String) As String
Dim rsPart As ADODB.Recordset
Dim nRate As Double

    estimasiWaktu = "00:00"
    Set rsPart = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE MDL = '" & Trim(nTipeKend) & "' AND FS_ITEM LIKE '%" & Format(nKm, "##,0") & "%'")
    If Not rsPart.EOF Then
        nRate = CDbl(rsPart("RATE")) * 60
        estimasiWaktu = Format(Round(nRate / 60), "0,#") & ":" & Format(Round(nRate Mod 60), "0,#")
    End If
End Function

Private Function estimasiPart(nKm As String, nTipeKend As String) As Double
Dim rsPart As ADODB.Recordset
Dim nHarga As Double

    estimasiPart = 0
    
    If Log.RPLHistori = 0 Then
        Set rsPart = objO.FetchData("SELECT * FROM HO_PARTPACK WITH(NOLOCK) WHERE MDL = '" & Trim(nTipeKend) & "' AND FS_ITEM LIKE '%" & Format(nKm, "##,0") & "%'")
    Else
        Set rsPart = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_PARTPACK WITH(NOLOCK) WHERE MDL = '" & Trim(nTipeKend) & "' AND FS_ITEM LIKE '%" & Format(nKm, "##,0") & "%'")
    End If
    If Not rsPart.EOF Then
        Do While Not rsPart.EOF
            estimasiPart = estimasiPart + objR.pricePartMonth(Trim(rsPart("ITEM_PART")), CDbl(rsPart("QTY")))
            rsPart.MoveNext
        Loop
    End If
End Function
 
Private Function dealerJual(noRangka As String) As Integer
Dim rsSearch As ADODB.Recordset
    
    Set rsSearch = objO.FetchData("SELECT INTERNAL FROM CUSTDAT A WITH(NOLOCK) LEFT JOIN T_DEALER B WITH(NOLOCK) ON A.FS_KD_DEALER = B.FS_KD_DEALER WHERE NO_CHASSIS = '" & Trim(noRangka) & "' AND FS_LOG_DEL = ''")
    If Not rsSearch.EOF Then
        dealerJual = Trim(rsSearch("INTERNAL"))
    Else
        dealerJual = 0
    End If
End Function

 
Private Function tampilLebihKecil(date1 As Date, date2 As Date) As Date

    tampilLebihKecil = IIf(DateDiff("d", date1, date2) > 0, date1, date2)
End Function


