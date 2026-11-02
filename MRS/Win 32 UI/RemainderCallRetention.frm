VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form RemainderCallRetention 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   Picture         =   "RemainderCallRetention.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmeHitungUlang 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   3540
      TabIndex        =   106
      Top             =   2190
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton countH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "c"
         Height          =   330
         Left            =   4000
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   667
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdBatal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   735
         Index           =   0
         Left            =   2640
         Picture         =   "RemainderCallRetention.frx":DA6C
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveEst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   735
         Left            =   1920
         Picture         =   "RemainderCallRetention.frx":E0CB
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtTglEst 
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
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   116
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtKmEst 
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
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   114
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtKMNow 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   111
         Top             =   667
         Width           =   1575
      End
      Begin TDBDate6Ctl.TDBDate tglNow 
         Height          =   315
         Left            =   2400
         TabIndex        =   110
         Top             =   1035
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         Calendar        =   "RemainderCallRetention.frx":E797
         Caption         =   "RemainderCallRetention.frx":E8C3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "RemainderCallRetention.frx":E92F
         Keys            =   "RemainderCallRetention.frx":E94D
         Spin            =   "RemainderCallRetention.frx":E9AB
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
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         X1              =   120
         X2              =   5160
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Estimasi :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   675
         TabIndex        =   115
         Top             =   2820
         Width           =   1545
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM Estimasi :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1080
         TabIndex        =   113
         Top             =   2460
         Width           =   1140
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   3000
         X2              =   5040
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line19 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   3480
         X2              =   5040
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Jatuh Tempo Service"
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
         Left            =   240
         TabIndex        =   112
         Top             =   1755
         Width           =   2580
      End
      Begin VB.Shape Shape9 
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         X1              =   120
         X2              =   5160
         Y1              =   2060
         Y2              =   2060
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         X1              =   120
         X2              =   5160
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1455
         TabIndex        =   109
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM Sekarang :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   990
         TabIndex        =   108
         Top             =   720
         Width           =   1230
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   2880
         X2              =   5040
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   3480
         X2              =   5040
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Masukkan KM Sekarang"
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
         Left            =   240
         TabIndex        =   107
         Top             =   200
         Width           =   2070
      End
      Begin VB.Shape Shape8 
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5055
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Height          =   4575
         Left            =   105
         Top             =   105
         Width           =   5055
      End
   End
   Begin VB.Frame frmeTidakDiterima 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   2760
      TabIndex        =   119
      Top             =   2400
      Visible         =   0   'False
      Width           =   7815
      Begin VB.OptionButton OptAlasan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4. Lainnya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   355
         Index           =   3
         Left            =   240
         TabIndex        =   127
         Top             =   1680
         Width           =   6615
      End
      Begin VB.TextBox txtAlasanLain 
         Appearance      =   0  'Flat
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
         Left            =   480
         MaxLength       =   100
         TabIndex        =   126
         Top             =   2040
         Width           =   7095
      End
      Begin VB.OptionButton OptAlasan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3. Tidak bertemu dengan yang bersangkutan (untuk Fix Phone)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   125
         Top             =   1320
         Width           =   6615
      End
      Begin VB.OptionButton OptAlasan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2. No Telp Tidak Aktif / Sibuk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   124
         Top             =   960
         Width           =   6615
      End
      Begin VB.OptionButton OptAlasan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1. No Telp Tidak Valid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   123
         Top             =   600
         Value           =   -1  'True
         Width           =   6615
      End
      Begin VB.CommandButton simpanAlasan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   735
         Left            =   3120
         Picture         =   "RemainderCallRetention.frx":E9D3
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdBatal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   735
         Index           =   1
         Left            =   3840
         Picture         =   "RemainderCallRetention.frx":F09F
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   2760
         Width           =   735
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   3615
         Left            =   120
         Top             =   120
         Width           =   7575
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alasan Telp Tidak diterima"
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
         Left            =   240
         TabIndex        =   122
         Top             =   195
         Width           =   2250
      End
      Begin VB.Line Line28 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   4800
         X2              =   7440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line27 
         BorderColor     =   &H000040C0&
         BorderWidth     =   2
         X1              =   4080
         X2              =   7440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Shape Shape13 
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   7575
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   480
      TabIndex        =   85
      Top             =   6840
      Width           =   11655
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
         TabIndex        =   100
         Top             =   960
         Width           =   7815
      End
      Begin VB.CommandButton cmdRecord 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Record"
         Height          =   732
         Left            =   9960
         Picture         =   "RemainderCallRetention.frx":F6FE
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   720
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
         Picture         =   "RemainderCallRetention.frx":FD6B
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdBooking 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Booking"
         Enabled         =   0   'False
         Height          =   732
         Left            =   2520
         Picture         =   "RemainderCallRetention.frx":104FB
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   0
         Width           =   735
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
         TabIndex        =   88
         Top             =   360
         Width           =   6015
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
         TabIndex        =   91
         Top             =   0
         Width           =   615
      End
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
         TabIndex        =   90
         Top             =   0
         Width           =   855
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
         TabIndex        =   102
         Top             =   960
         Width           =   75
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
         TabIndex        =   101
         Top             =   960
         Width           =   555
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
         TabIndex        =   99
         Top             =   15
         Width           =   120
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
         TabIndex        =   98
         Top             =   0
         Width           =   120
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         X1              =   9600
         X2              =   120
         Y1              =   1680
         Y2              =   1680
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
         TabIndex        =   93
         Top             =   1440
         Width           =   9615
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   9600
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   9600
         Y1              =   840
         Y2              =   840
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
         TabIndex        =   87
         Top             =   45
         Width           =   75
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
         TabIndex        =   86
         Top             =   60
         Width           =   1200
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
      Left            =   480
      TabIndex        =   30
      Top             =   4200
      Width           =   11655
      Begin VB.TextBox txtStatusData 
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
         TabIndex        =   133
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdMaster 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Master"
         Enabled         =   0   'False
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
         Left            =   8520
         Picture         =   "RemainderCallRetention.frx":10B50
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   120
         Width           =   735
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
         TabIndex        =   104
         Top             =   990
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Enabled         =   0   'False
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
         Left            =   9240
         Picture         =   "RemainderCallRetention.frx":111E2
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   735
         Left            =   9960
         Picture         =   "RemainderCallRetention.frx":11874
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   735
         Left            =   10680
         Picture         =   "RemainderCallRetention.frx":11F40
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   120
         Width           =   735
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
         TabIndex        =   40
         Top             =   960
         Width           =   2655
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
         TabIndex        =   39
         Top             =   1680
         Width           =   6015
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
         TabIndex        =   38
         Top             =   1320
         Width           =   6015
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
         TabIndex        =   37
         Top             =   600
         Width           =   1455
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
         TabIndex        =   36
         Top             =   600
         Width           =   2655
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
         TabIndex        =   35
         Top             =   1680
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
         TabIndex        =   34
         Top             =   1320
         Width           =   1815
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
         TabIndex        =   33
         Top             =   2040
         Width           =   1815
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
         TabIndex        =   32
         Top             =   2040
         Width           =   2655
      End
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
         TabIndex        =   31
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Data"
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
         Left            =   5160
         TabIndex        =   135
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label36 
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
         TabIndex        =   134
         Top             =   960
         Width           =   75
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   8400
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3135
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11520
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11520
         Y1              =   2520
         Y2              =   2520
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
         TabIndex        =   61
         Top             =   960
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
         TabIndex        =   60
         Top             =   960
         Width           =   1335
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
         TabIndex        =   59
         Top             =   360
         Width           =   75
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
         TabIndex        =   58
         Top             =   600
         Width           =   75
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
         TabIndex        =   57
         Top             =   600
         Width           =   855
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
         TabIndex        =   56
         Top             =   1320
         Width           =   585
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
         TabIndex        =   55
         Top             =   1680
         Width           =   690
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
         TabIndex        =   54
         Top             =   1320
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
         TabIndex        =   53
         Top             =   1680
         Width           =   75
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
         TabIndex        =   52
         Top             =   600
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
         TabIndex        =   51
         Top             =   600
         Width           =   1065
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
         TabIndex        =   50
         Top             =   1680
         Width           =   75
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
         TabIndex        =   49
         Top             =   1320
         Width           =   75
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
         TabIndex        =   48
         Top             =   1680
         Width           =   435
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
         TabIndex        =   47
         Top             =   1320
         Width           =   480
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
         TabIndex        =   46
         Top             =   2040
         Width           =   75
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
         TabIndex        =   45
         Top             =   2040
         Width           =   600
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   3840
         X2              =   8160
         Y1              =   240
         Y2              =   240
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   2040
         Width           =   900
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
         TabIndex        =   41
         Top             =   2040
         Width           =   75
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
         TabIndex        =   62
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   2640
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   11655
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbPilihKeterangan 
         Height          =   330
         Left            =   6120
         TabIndex        =   131
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
         TabIndex        =   130
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdCalculate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Hitung Ulang"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   10200
         Picture         =   "RemainderCallRetention.frx":1259F
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   84
         Top             =   120
         Visible         =   0   'False
         Width           =   4815
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
         TabIndex        =   83
         Top             =   600
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   1440
         Width           =   7335
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
            TabIndex        =   19
            Top             =   0
            Width           =   1935
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
            TabIndex        =   18
            Top             =   0
            Width           =   1335
         End
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
            TabIndex        =   17
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   1080
         Width           =   7575
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
            TabIndex        =   15
            Top             =   0
            Width           =   1335
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
            TabIndex        =   14
            Top             =   0
            Width           =   1935
         End
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
            TabIndex        =   13
            Top             =   0
            Width           =   2655
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1800
         Width           =   7575
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
            TabIndex        =   11
            Top             =   0
            Width           =   2415
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
            TabIndex        =   10
            Top             =   0
            Width           =   1335
         End
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
            TabIndex        =   9
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   2160
         Width           =   7455
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
            TabIndex        =   7
            Top             =   0
            Width           =   1935
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
            TabIndex        =   6
            Top             =   0
            Width           =   1335
         End
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
            TabIndex        =   5
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   2400
         TabIndex        =   1
         Top             =   80
         Width           =   4215
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
            TabIndex        =   3
            Top             =   0
            Width           =   1935
         End
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
            TabIndex        =   2
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSComCtl2.DTPicker JamNewConfirm 
         Height          =   300
         Left            =   5520
         TabIndex        =   78
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
         Format          =   342753282
         CurrentDate     =   39905
      End
      Begin TDBDate6Ctl.TDBDate TglNewKonfirm 
         Height          =   315
         Left            =   3360
         TabIndex        =   79
         Top             =   593
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "RemainderCallRetention.frx":14BD7
         Caption         =   "RemainderCallRetention.frx":14D03
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "RemainderCallRetention.frx":14D6F
         Keys            =   "RemainderCallRetention.frx":14D8D
         Spin            =   "RemainderCallRetention.frx":14DEB
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   10080
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   960
         Y2              =   960
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
         TabIndex        =   82
         Top             =   623
         Width           =   495
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
         TabIndex        =   81
         Top             =   623
         Width           =   855
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
         TabIndex        =   80
         Top             =   600
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
         TabIndex        =   29
         Top             =   1500
         Width           =   1815
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
         TabIndex        =   28
         Top             =   1140
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
         TabIndex        =   27
         Top             =   1140
         Width           =   1800
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
         TabIndex        =   26
         Top             =   1485
         Width           =   75
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   2530
         Y2              =   2530
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
         TabIndex        =   25
         Top             =   1845
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
         TabIndex        =   24
         Top             =   1860
         Width           =   1155
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
         TabIndex        =   23
         Top             =   2205
         Width           =   75
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
         TabIndex        =   22
         Top             =   2220
         Width           =   1245
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
         TabIndex        =   21
         Top             =   140
         Width           =   1245
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
         TabIndex        =   20
         Top             =   130
         Width           =   75
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   480
      TabIndex        =   64
      Top             =   3120
      Width           =   11655
      Begin VB.TextBox txtKmNextSvc 
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
         Left            =   2520
         MaxLength       =   25
         TabIndex        =   68
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txttglNextSvc 
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
         ForeColor       =   &H80000005&
         Height          =   330
         Left            =   2520
         MaxLength       =   40
         TabIndex        =   67
         Top             =   120
         Width           =   1575
      End
      Begin VB.ComboBox cboMenit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8055
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   525
         Width           =   735
      End
      Begin VB.ComboBox cboJam 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6855
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   525
         Width           =   735
      End
      Begin TDBNumber6Ctl.TDBNumber txtHargaJasa 
         Height          =   330
         Left            =   7200
         TabIndex        =   69
         Top             =   120
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   582
         Calculator      =   "RemainderCallRetention.frx":14E13
         Caption         =   "RemainderCallRetention.frx":14E33
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "RemainderCallRetention.frx":14E9F
         Keys            =   "RemainderCallRetention.frx":14EBD
         Spin            =   "RemainderCallRetention.frx":14F07
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
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
      Begin TDBNumber6Ctl.TDBNumber txtHargaPart 
         Height          =   330
         Left            =   9480
         TabIndex        =   103
         Top             =   120
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2822
         _ExtentY        =   582
         Calculator      =   "RemainderCallRetention.frx":14F2F
         Caption         =   "RemainderCallRetention.frx":14F4F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "RemainderCallRetention.frx":14FBB
         Keys            =   "RemainderCallRetention.frx":14FD9
         Spin            =   "RemainderCallRetention.frx":15023
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
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
         MinValue        =   0
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
      Begin VB.Label lblFOP 
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
         Height          =   375
         Left            =   9480
         TabIndex        =   132
         Top             =   480
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label14 
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
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label13 
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
         TabIndex        =   76
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km Estimasi Service"
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
         TabIndex        =   75
         Top             =   480
         Width           =   1965
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Estimasi Service"
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
         TabIndex        =   74
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Waktu                               Jam                  Menit"
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
         Left            =   4920
         TabIndex        =   73
         Top             =   540
         Width           =   4380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Harga                Rp.                                      + Rp."
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
         Left            =   4920
         TabIndex        =   72
         Top             =   120
         Width           =   4410
      End
      Begin VB.Label Label22 
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
         Left            =   6720
         TabIndex        =   71
         Top             =   510
         Width           =   75
      End
      Begin VB.Label Label23 
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
         Left            =   6720
         TabIndex        =   70
         Top             =   120
         Width           =   75
      End
      Begin VB.Shape Shape5 
         Height          =   345
         Left            =   6840
         Top             =   510
         Width           =   765
      End
      Begin VB.Shape Shape6 
         Height          =   345
         Left            =   8040
         Top             =   510
         Width           =   765
      End
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2655
      X2              =   4080
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1680
      X2              =   4080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   615
      X2              =   4080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM REMINDER SERVICE - CALL"
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
      Left            =   1440
      TabIndex        =   63
      Top             =   120
      Width           =   10215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11685
   End
End
Attribute VB_Name = "RemainderCallRetention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TglNext As Date
'Private Remainder As String
Private noWO As String
Private Editing As Boolean
Private objR As New ClsRemainder
Private Rec As Boolean
Private rsRem As ADODB.Recordset
Private rsAll As ADODB.Recordset
Private rsCall As ADODB.Recordset

Private is1000 As Boolean
Private isNoPol As String
Private isHansaSA   As Boolean

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
Dim mdlKend As String
Private isMRABoard As Boolean

Private tagihanFOP  As String
Private kodeMDL     As String

Public Property Let xIsMRABoard(value As Boolean)
    isMRABoard = value
End Property

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

Public Property Let IsHansaBySA(value As Boolean)
   isHansaSA = value
End Property

Public Property Let xNoPol(value As String)
   isNoPol = value
End Property

Private Sub chJual_Click()
    
    If chJual.value = Checked Then
        If MsgBox("Berhasil mendapatkan Data pelanggan yang baru ?", vbYesNo + vbDefaultButton1, "Customer Baru") = vbYes Then
            cmdMaster.Enabled = True
            Call cmdMaster_Click
            flagTdkDiundang = False
        Else
            flagTdkDiundang = True
            cmdMaster.Enabled = False
        End If
    Else
        cmdMaster.Enabled = False
        flagTdkDiundang = False
    End If
End Sub

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

Private Sub cmdBatal_Click(Index As Integer)
    Select Case Index
        Case 0
            frmeHitungUlang.Visible = False
        Case 1
            frmeTidakDiterima.Visible = False
    End Select
End Sub

Private Sub cmdBooking_Click()
      Dim rsSearch As ADODB.Recordset

1     On Error Resume Next
2         isBookByMrs = True
3         If rsRem.Fields(0).Name = "NoSPK" Then
4             rsAll.Find "SPK_No ='" & rsRem.Fields(0).value & "'", , adSearchForward, 1
5         Else
6             rsAll.Find "NoInv ='" & rsRem.Fields(0).value & "'", , adSearchForward, 1
7         End If
          
8         With frmBooking
              Dim nDpn1 As String
              Dim nTgh1 As String
              Dim nBlk1 As String
9             Call pecahNoPol(isNoPol, nDpn1, nTgh1, nBlk1)
              
10            .Text(0) = Trim(nDpn1)
11            .Text(56) = Trim(nTgh1)
12            .Text(57) = Trim(nBlk1)
              
              ''20171010 RETENTION BOARD
13            If isMRABoard = True Then
14                .isRetentionBoard = True
15            End If
              
16            If rsRem.Fields(0).Name = "NoSPK" Then
17                .Text(3).Text = rsAll.Fields(6).value
18                .Text(4).Text = rsAll.Fields(7).value
19                .Text(5).Text = rsAll.Fields(8).value
20                .Text(8).Text = rsAll.Fields(13).value
21                .Text(6).Text = rsAll.Fields(14).value
22                .Text(7).Text = rsAll.Fields(15).value
23                .Text(10).Text = rsAll.Fields(16).value
24                .Text(13).Text = rsAll.Fields(34).value
25                .Text(41).Text = rsAll.Fields(31).value
26                .Text(9).Text = rsAll.Fields(33).value
27                .Text(12).Text = rsAll.Fields(17).value
28                Call .panggilHistoryBok(rsAll.Fields(12).value, Trim(isNoPol))
                  
29                        Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH(NOLOCK) WHERE NAMALgn = '" & rsAll.Fields(6).value & "'")
30                        If rsSearch.EOF Then
31                            .Text(2) = ""
32                            .Text(44) = ""
33                        Else
34                            .Text(2) = rsSearch("KODELGN")
35                            .Text(44) = rsSearch("KODEWIL")
36                        End If
                      
37                    If Log.RPLHistori = 0 Then
38                        Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & rsAll.Fields(16).value & "'")
39                    Else
40                        Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & rsAll.Fields(16).value & "'")
41                    End If
42                    If Not rsSearch.BOF Then
43                        .Text(10) = rsSearch("MDL")
44                        .Text(12) = rsSearch("Desc")
45                        .Text(11) = rsSearch("MODEL")
46                    End If
47            Else
48                .Text(2).Text = rsAll.Fields(17).value
49                .Text(44).Text = rsAll.Fields(57).value
50                .Text(3).Text = rsAll.Fields(6).value
51                .Text(4).Text = rsAll.Fields(7).value
52                .Text(5).Text = rsAll.Fields(8).value
53                .Text(8).Text = rsAll.Fields(13).value
54                .Text(6).Text = rsAll.Fields(14).value
55                .Text(7).Text = rsAll.Fields(15).value
56                .Text(9).Text = rsAll.Fields(24).value
57                .Text(10).Text = rsAll.Fields(22).value
58                .Text(11).Text = rsAll.Fields(23).value
59                .Text(12).Text = rsAll.Fields(11).value
60                .Text(13).Text = rsAll.Fields(25).value
61                .Text(14).Text = rsAll.Fields(26).value
62                .Text(15).Text = rsAll.Fields(27).value
63                .Text(41).Text = rsAll.Fields(12).value
64                .Text(9).Text = rsAll.Fields(54).value
65                .Text(11).Text = rsAll.Fields(28).value
66                .Text(12).Text = rsAll.Fields(29).value
67                .Text(14).Text = rsAll.Fields(31).value
68                .Text(15).Text = rsAll.Fields(32).value
69                .Text(16).Text = Format(IIf(IsNull(rsAll.Fields(58).value) = True, "01-01-1900", rsAll.Fields(58).value), "DD-MM-YYYY")
70                .Text(17).Text = Format(IIf(IsNull(rsAll.Fields(33).value) = True, "01-01-1900", rsAll.Fields(33).value), "DD-MM-YYYY")
71                Call .panggilHistoryBok(rsAll.Fields(12).value, Trim(isNoPol))
72            End If
73            .show vbModal
74        End With
End Sub

Private Sub cmdCalculate_Click()
Dim rsSearch As ADODB.Recordset
    
    txtKMNow = 0
    tglNow = Format(isnow, "DD/MM/YYYY")
    txtKmEst = txtKmNextSvc
    txtTglEst = txttglNextSvc
    
    Set rsSearch = objO.FetchData("SELECT TOP 1 * FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(txtNoChassis) & "' ORDER BY INVDATE DESC")
    If Not rsSearch.EOF Then
        kmLast = CDbl(rsSearch("KM"))
        TglLast = Format(rsSearch("DATE"), "DD/MM/YYYY")
    End If
    
    If is1000 Then
        Set rsSearch = objO.FetchDataSPK("SELECT DATEADD(day,30,TGLCR1) AS JT_BLN FROM T_MRS WITH(NOLOCK) WHERE NORANGKA = '" & Trim(txtNoChassis) & "'")
    Else
        Set rsSearch = objO.FetchData("SELECT * FROM M_ESTIMASI WITH(NOLOCK) WHERE NORANGKA = '" & Trim(txtNoChassis) & "' AND SBENEXT = '" & Trim(rsRem("KMNEXT")) & "' ")
    End If
    If Not rsSearch.EOF Then
        tglJTBln = Format(rsSearch("JT_BLN"), "DD/MM/YYYY")
    End If
    
    frmeHitungUlang.Visible = True
End Sub

Private Sub cmdCancel_Click()
   Rec = False
   Editing = False

      Call BindToControl
      Dim i As Integer
      With Me
         .txtKetTelp.Enabled = False
         For i = 0 To 2
            .OptPel(i).Enabled = True
            .OptKen(i).Enabled = True
            .OptDM(i).Enabled = True
            .OptSMS(i).Enabled = True
         Next i
         .cmdEdit.Enabled = True
         .cmdCancel.Enabled = False
         .cmdSave.Enabled = False
         .OptTelp(0).Enabled = True
         .OptTelp(1).Enabled = True
         .OptApp(0).Enabled = True
         .OptApp(1).Enabled = True
         .cmdExit.Enabled = True
         .cmdRecord.Enabled = True
         .txtNama.Enabled = False
         .txtAlamat.Enabled = False
         .txtKota.Enabled = False
         .txtPos.Enabled = False
         .txtTelp.Enabled = False
         .txtCDMA.Enabled = False
         .txtGSM.Enabled = False
         .txtNoPol.Enabled = False
         .cmdCancel.Enabled = False
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

Private Sub cmdMaster_Click()
    With frmtCustomer
        If Trim(txtNoChassis) <> "" Then
            .Text(50) = Left(txtNoChassis, 3)
            If Len(txtNoChassis) > 3 Then
                .Text(1) = Mid(txtNoChassis, 4, 6)
            End If
            If Len(txtNoChassis) > 9 Then
                .Text(0) = Mid(txtNoChassis, 10, Len(txtNoChassis) - 9)
            End If
        Else
            MsgBox "No Rangka tidak lengkap. "
        End If
        .show vbModal
    End With
End Sub

Private Sub cmdSaveEst_Click()
    Call countH_Click
    If CDbl(txtKMNow) > 0 And CDbl(txtKMNow) < CDbl(txtKmEst) Then
         ''2016-10-11 pengecekan service 1000 atau tidak jika 1000 update ke spk
        If is1000 = True Then
            Call objO.UpdateData("UPDATE [" & Log.LogSPK & "]..T_MRS SET TGL1 = '" & Format(txtTglEst, "MM/DD/YYYY") & "' WHERE NORANGKA = '" & Trim(txtNoChassis) & "' ")
        Else
            Call objO.UpdateData("UPDATE M_ESTIMASI SET JT_TGL = '" & Format(txtTglEst, "MM/DD/YYYY") & "' WHERE NORANGKA = '" & Trim(txtNoChassis) & "' AND SBENEXT = '" & Trim(rsRem("KMNEXT")) & "'")
        End If
'        Call objO.UpdateData("UPDATE M_ESTIMASI SET JT_TGL = '" & Format(txtTglEst, "MM/DD/YYYY") & "' WHERE NORANGKA = '" & Trim(txtNoChassis) & "' AND SBENEXT = '" & Trim(rsRem("KMNEXT")) & "'")
        Call objO.InsertData("Insert into HitungUlang (NoRangka,LastNoWo,Tgl,KmSkrg,KmEst,TglEst) values ('" & Trim(txtNoChassis) & "','" & noWO & "',getdate()," & Val(txtKMNow.Text) & "," & Val(txtKmEst.Text) & ",'" & TglNul2(txtTglEst.Text) & "')")
        MsgBox "Perhitungan Ulang berhasil disimpan !"
        frmeHitungUlang.Visible = False
    Else
        If CDbl(txtKMNow) = 0 Then
            MsgBox "KM sekarang harus lebih besar dari " & Format(CDbl(kmLast), "##,#"), vbCritical + vbOKOnly, "Salah input"
        ElseIf CDbl(txtKMNow) > CDbl(txtKmEst) Then
            MsgBox "KM sekarang lebih besar dari KM estimasi : " & Format(CDbl(txtKmEst), "##,#") & vbCr & "Unit harus segera service, tidak perlu hitung ulang !!", vbCritical + vbOKOnly, "Salah input"
            frmeHitungUlang.Visible = False
        End If
    End If
End Sub

Private Sub countH_Click()
    If Trim(txtKMNow) = "" Then txtKMNow = 0
    If CDbl(txtKMNow) > CDbl(kmLast) Then
        txtTglEst = tampilLebihKecil(tglJTBln, DateAdd("d", Round((CDbl(txtKmEst) - CDbl(txtKMNow)) / ((CDbl(txtKMNow) - CDbl(kmLast)) / Abs(DateDiff("d", TglLast, isnow)))), Format(isnow, "DD/MM/YYYY")))
    End If
End Sub

Private Sub Form_Load()
   Rec = False
   Editing = False
   Call ListComboWaktu
   Call BindToControl
   OptTelp(0).value = True
   Call ItemRemainderCall
   Call Eta
   gBantuan = ""
   gBantuan2 = ""
   
    If Log.NHS = 1 Then
        If NulSID(txtNoChassis.Text) <> "" Then
            Call panggilWarranty(NulSID(txtNoChassis.Text), "")
        End If
    End If
End Sub

Private Sub ListComboWaktu()
   With Me
      .cboJam.Clear
      Dim i As Integer
      For i = 0 To 24
         .cboJam.AddItem i
      Next i
      .cboMenit.Clear
      For i = 0 To 59
         .cboMenit.AddItem i
      Next i
      .cboJam.Text = 0
      .cboMenit.Text = 0
   End With
End Sub

Private Sub BindToControl()
Dim KmNext As String

   
With Me
   If is1000 = False Then
      noWO = Trim(NullText(rsRem.Fields(1).value))
      KmNext = Trim(NullText(rsRem.Fields(20).value))
      mdlKend = objR.myModel(Trim(NullText(rsRem.Fields(12).value)))
      .txttglNextSvc.Text = Trim(NullText(rsRem.Fields(4).value))
      .txtKmNextSvc.Text = Trim(NullText(rsRem.Fields(5).value))
      .txtNoChassis.Text = Trim(NullText(rsRem.Fields(12).value))
      .txtNoEngine.Text = Trim(NullText(rsRem.Fields(11).value))
      .txtNama.Text = Trim(NullText(rsRem.Fields(6).value))
      .txtAlamat.Text = Trim(NullText(rsRem.Fields(7).value))
      .txtKota.Text = Trim(NullText(rsRem.Fields(8).value))
      .txtPos.Text = Trim(NullText(rsRem.Fields(9).value))
      .txtNoPol.Text = Trim(NullText(rsRem.Fields(10).value))
      .txtTelp.Text = Trim(NullText(rsRem.Fields(13).value))
      .txtGSM.Text = Trim(NullText(rsRem.Fields(14).value))
      .txtCDMA.Text = Trim(NullText(rsRem.Fields(15).value))
      .txtStatusData.Text = "" 'Trim(NullText(rsRem.Fields(35).value))
   Else
      noWO = Trim(NullText(rsRem.Fields(0).value))
      KmNext = "1000"
      mdlKend = Trim(NullText(rsRem.Fields(16).value))
      .txttglNextSvc.Text = Trim(NullText(rsRem.Fields(30).value))
      .txtKmNextSvc.Text = "1000"
      .txtNoChassis.Text = Trim(NullText(rsRem.Fields(12).value))
      .txtNoEngine.Text = Trim(NullText(rsRem.Fields(17).value))
      .txtNama.Text = Trim(NullText(rsRem.Fields(6).value))
      .txtAlamat.Text = Trim(NullText(rsRem.Fields(7).value))
      .txtKota.Text = Trim(NullText(rsRem.Fields(8).value))
      .txtPos.Text = Trim(NullText(rsRem.Fields(9).value))
      .txtNoPol.Text = Trim(NullText(rsRem.Fields(10).value))
      .txtTelp.Text = Trim(NullText(rsRem.Fields(13).value))
      .txtGSM.Text = Trim(NullText(rsRem.Fields(14).value))
      .txtCDMA.Text = Trim(NullText(rsRem.Fields(15).value))
      .txtStatusData.Text = Trim(NullText(rsRem.Fields(35).value))
   End If
   ' HITUNGAN ESTIMASI HARGA
   
   .txtHargaJasa = estimasiJasa(KmNext, mdlKend, .txtNoChassis)
   .txtHargaPart = estimasiPart(KmNext, mdlKend)
   
   ' HITUNGAN ESTIMASI WAKTU
   Dim wktEstimasi As String
   .cboJam.Enabled = True
   .cboMenit.Enabled = True
   wktEstimasi = estimasiWaktu(KmNext, mdlKend)
   .cboJam = CDbl(Left(wktEstimasi, InStr(wktEstimasi, ":") - 1))
   .cboMenit = CDbl(Mid(wktEstimasi, InStr(wktEstimasi, ":") + 1, Len(wktEstimasi) - InStr(wktEstimasi, ":")))
   
    Nama = .txtNama.Text
    Alamat = .txtAlamat.Text
    KodePos = .txtPos.Text
    Kota = .txtKota.Text
    Propinsi = ""
    Rumah = .txtTelp.Text
    GSM = .txtGSM.Text
    CDMA = .txtCDMA.Text
    NoPolisi = .txtNoPol.Text
    
    If Log.freepart = 1 And NulSID(.txtNoChassis.Text) <> "" And is1000 = False Then
        tagihanFOP = ""
        kodeMDL = objO.FetchData_String("select Mdl from SingleIDOri..UnitCustomer where norangka = '" & NulSID(.txtNoChassis.Text) & "'")
        If CekUnitTerlibatFOP(.txtNoChassis.Text, kodeMDL, tagihanFOP) Then
            lblFOP.Caption = "Free Part + Oli : " & tagihanFOP
        Else
            lblFOP.Caption = ""
        End If
        ''harga labor dan part default 0
        .txtHargaJasa = 0
        .txtHargaPart = 0
   End If
End With
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
      .cmdEdit.Enabled = False
      .cmdCancel.Enabled = False
      .cmdSave.Enabled = False
      .txtKetTelp.Enabled = True
      .txtPesan.Enabled = False
      .txtKetTelp.SetFocus
      .chJual.Enabled = False
      If cmdCancel.Enabled = False Then
         .OptPel(0).value = True
         .OptKen(0).value = True
         .OptDM(0).value = True
         .OptSMS(0).value = True
      End If
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
      .cmdCalculate.Enabled = False
   ElseIf OptTelp(1).value = True Then
        ''diterima
        sSQL = "select keterangantipe as Tipe, pesan as keterangan,isRecovery FROM M_VoiceOfCustomer with (nolock) where tipe = 1"
        
      .OptTelp(1).ForeColor = vbBlue
      .cmdEdit.Enabled = True
      .cmdCancel.Enabled = False
      .cmdSave.Enabled = False
      .txtKetTelp.Enabled = False
      .txtKetTelp.Text = ""
      .chJual.Enabled = True
      If cmdCancel.Enabled = False Then
         .OptPel(2).value = True
         .OptKen(2).value = True
         .OptDM(2).value = True
         .OptSMS(2).value = True
      End If
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
      .cmdCalculate.Enabled = True
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
   Dim rsEta As ADODB.Recordset
   Set rsEta = objR.FetchEta(noWO, "SM")
   If rsEta.RecordCount = 0 Then Exit Sub
   With Me
      .txtHargaJasa = rsEta.Fields(0).value
      .txtHargaPart = rsEta.Fields(1).value
      .cboJam = Val(rsEta.Fields(2).value)
      .cboMenit.Text = rsEta.Fields(3).value
   End With
End Sub

Private Sub ItemRemainderCall()
   Set rsCall = objR.FetchItemRemainderCall(noWO, txtNoChassis, NulA(txtKmNextSvc))
   If rsCall.RecordCount = 0 Then
      lCall1.Caption = "Call : 1"
      lCall2.Caption = "Call : 1"
      Exit Sub
   End If
   rsCall.MoveLast
   lCall1.Caption = "Call : " & rsCall.RecordCount + 1
   lCall2.Caption = "Call : " & rsCall.RecordCount + 1
   Call BindingLastRemainder(rsCall)
End Sub

Private Sub BindingLastRemainder(rs As ADODB.Recordset)
With Me
   .OptTelp(Val(IIf(rs.Fields(6).value = True, 1, 0))).value = True
   .txtKetTelp.Text = Trim(NullText(rs.Fields(7).value))
   .OptPel(Val(rs.Fields(11).value)).value = True
   .OptKen(Val(rs.Fields(12).value)).value = True
   .OptDM(Val(rs.Fields(13).value)).value = True
   .OptSMS(Val(rs.Fields(14).value)).value = True
   .txtHargaJasa = Val(rs.Fields(15).value)
   .txtHargaPart = Val(rs.Fields(16).value)
   .cboJam.Text = Val(rs.Fields(17).value)
   .cboMenit.Text = Val(rs.Fields(18).value)
   .OptApp(Val(IIf(rs.Fields(20).value = True, 1, 0))).value = True
   .txtApp.Text = Trim(NullText(rs.Fields(21).value))
   .txtPesan.Text = Trim(NullText(rs.Fields(22).value))
   
End With
End Sub

Private Sub cmdEdit_Click()
Dim i As Integer
With Me
   .txtKetTelp.Enabled = False
   .txtPesan.Enabled = False
   For i = 0 To 2
      .OptPel(i).Enabled = False
      .OptKen(i).Enabled = False
      .OptDM(i).Enabled = False
      .OptSMS(i).Enabled = False
   Next i
   .cmdEdit.Enabled = False
   .cmdCancel.Enabled = True
   .cmdSave.Enabled = True
   .OptTelp(0).Enabled = False
   .OptTelp(1).Enabled = False
   .OptApp(0).Enabled = False
   .OptApp(1).Enabled = False
   .txtApp.Enabled = False
   .cmdBooking.Enabled = False
   .cmdExit.Enabled = False
   .cmdRecord.Enabled = False
   .txtNama.Enabled = True
   .txtAlamat.Enabled = True
   .txtKota.Enabled = True
   .txtPos.Enabled = True
   .txtTelp.Enabled = True
   .txtCDMA.Enabled = True
   .txtGSM.Enabled = True
   .txtNoPol.Enabled = True
End With
End Sub

Private Sub cmdSave_Click()
   If MsgBox("Simpan Perubahan Data Pelanggan ?", vbQuestion + vbYesNo) = vbYes Then
      Editing = True
      Call objR.UpdateItemCustdat_Use(txtNama.Text, txtAlamat.Text, txtKota.Text, txtPos.Text, txtTelp.Text, txtGSM.Text, txtCDMA.Text, txtNoPol.Text, Trim(txtNoChassis.Text))
      
    Dim objD As New ClsDataPelanggan
    With objD
       .Nama = Trim(txtNama.Text)
       .Alamat = Trim(txtAlamat.Text)
       .KodePos = Trim(txtPos.Text)
       .Kota = Trim(txtKota.Text)
       .NoPolisi = Trim(txtNoPol.Text)
       .Rumah = Trim(txtTelp.Text)
       .GSM = Trim(txtGSM.Text)
       .CDMA = Trim(txtCDMA.Text)
       If ch.value = 1 Then Call .AddNew_SCust(Trim(txtNoChassis.Text), Nama, Alamat, KodePos, Kota, Propinsi, Rumah, GSM, CDMA, NoPolisi)
       Call .UpdateItemCustomer(Trim(txtNoChassis.Text), 1)
    End With
            
      Dim i As Integer
      With Me
         .txtKetTelp.Enabled = False
         .txtPesan.Enabled = True
         For i = 0 To 2
            .OptPel(i).Enabled = True
            .OptKen(i).Enabled = True
            .OptDM(i).Enabled = True
            .OptSMS(i).Enabled = True
         Next i
         .cmdEdit.Enabled = True
         .cmdCancel.Enabled = False
         .cmdSave.Enabled = False
         .OptTelp(0).Enabled = True
         .OptTelp(1).Enabled = True
         .OptApp(0).Enabled = True
         .OptApp(1).Enabled = True
         .cmdExit.Enabled = True
         .cmdRecord.Enabled = True
         .txtNama.Enabled = False
         .txtAlamat.Enabled = False
         .txtKota.Enabled = False
         .txtPos.Enabled = False
         .txtTelp.Enabled = False
         .txtCDMA.Enabled = False
         .txtGSM.Enabled = False
         .txtNoPol.Enabled = False
         .cmdCancel.Enabled = False
      End With
   End If
End Sub

Private Sub cmdRecord_Click()
1        If MsgBox("Record Data & " & lCall1.Caption & " ?", vbQuestion + vbYesNo) = vbYes Then
2             If cbPilihKeterangan.value = "" Then
3                 MsgBox "Silahkan pilih keterangan konfirmasi Telp", vbInformation
4                 Exit Sub
5             End If
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

              
6             If isMRABoard = True Then
7                 sSQL = "INSERT INTO [listReminderBoard]([NoRangka],[NoWo],[NoPolisi],[NoBooking]) VALUES ("
8                 sSQL = sSQL & " '" & Trim(txtNoChassis.Text) & "','" & Trim(noWO) & "','" & IIf(Trim(txtNoPol.Text) <> "", Trim(txtNoPol.Text), gBantuan2) & "','" & Trim(gBantuan) & "')"
                  
9                 Call objO.InsertData(sSQL)
                   
10            End If
              
11            If Log.hansaRenkei = 1 And NulSID(cbPilihKeterangan.Columns(1).value) = NulSID("P. Sudah Service di dealer lain") Then
12                If is1000 Then
13                    sSQL = "INSERT INTO tbSudahServiceHansaRenkei (norangka,SbeNext,tglEntry,userEntry) VALUES ('" & txtNoChassis.Text & "','1000',GETDATE(),'" & gSA & "')"
14                Else
15                    sSQL = "INSERT INTO tbSudahServiceHansaRenkei (norangka,SbeNext,tglEntry,userEntry) VALUES ('" & txtNoChassis.Text & "','" & NulSID(rsRem!KmNext) & "',GETDATE(),'" & gSA & "')"
16                End If
                  ''sSQL = "INSERT INTO tbSudahServiceHansaRenkei (norangka,SbeNext,tglEntry,userEntry) VALUES ('" & txtNoChassis.Text & "','" & NulSID(rsRem!KmNext) & "',GETDATE(),'" & gSA & "')"
17                Call objO.InsertData(sSQL)
18            End If
              
19            If cbPilihKeterangan.Columns(2).value = "1" Then
                  ''jika masuk dalam range recovery
20                sSQL = "INSERT INTO [recoveryReminder]"
21                sSQL = sSQL & " ([no_wo],[noRangka],[pesanLastReminder],[tglReminder],[kmNext],[mdlKendaraan],[tglNextService] ,[noEngine] ,[nama] ,[alamat] ,[kota] ,[kodePos] ,[noPol] ,[telp] ,[gsm] ,[cdma])"
22                sSQL = sSQL & " VALUES "
23                sSQL = sSQL & " ('" & noWO & "','" & txtNoChassis.Text & "','" & Nul(cbPilihKeterangan.value) & "',CONVERT(varchar(20),GETDATE(),101),'" & txtKmNextSvc.Text & "','" & mdlKend & "','" & Format(txttglNextSvc, "mm/dd/yyyy") & "',"
24                sSQL = sSQL & "'" & Nul(txtNoEngine.Text) & "','" & Nul(txtNama.Text) & "','" & txtAlamat.Text & "','" & txtKota.Text & "','" & txtPos.Text & "','" & txtNoPol.Text & "','" & txtTelp.Text & "','" & txtGSM.Text & "','" & txtCDMA.Text & "')"
                    
25                Call objO.InsertData(sSQL)
                  
26            End If
              ''2017-08-15
27            Call simpanHasil
28       End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : simpanHasil
' DateTime  : 12/11/2025 13:39
' Author    : HO-VM-RUDI
' Purpose   :
'---------------------------------------------------------------------------------------
'
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
         
   On Error GoTo simpanHasil_Error

1        Call objR.AddNew_LogRemainder_Call(noWO, Format(isnow, "dd/mm/yyyy"), DateDiff("d", Format(isnow, "dd/mm/yyyy"), _
                                 Format(CDate(txttglNextSvc.Text), "dd/mm/yyyy")), txtNoChassis.Text, _
                                 Format(CDate(txttglNextSvc.Text), "dd/mm/yyyy"), Val(txtKmNextSvc.Text), _
                                 IIf(OptTelp(1).value = True, 1, 0), cbPilihKeterangan.value, IIf(ChUlang.value = 1, 1, 0), _
                                 Format(TglNewKonfirm.value, "dd/mm/yyyy"), Format(Me.JamNewConfirm.value, "hh:nn:ss"), _
                                 IIf(OptPel(1).value = True, 1, (IIf(OptPel(0).value = True, 0, 2))), _
                                 IIf(OptKen(1).value = True, 1, (IIf(OptKen(0).value = True, 0, 2))), _
                                 IIf(OptDM(1).value = True, 1, (IIf(OptDM(0).value = True, 0, 2))), _
                                 IIf(OptSMS(1).value = True, 1, (IIf(OptSMS(0).value = True, 0, 2))), _
                                 Val(Replace(txtHargaJasa, ".", "")), Val(Replace(txtHargaPart, ".", "")), Val(cboJam.Text), Val(cboMenit.Text), IIf(Editing = True, 1, 0), _
                                 IIf(OptApp(1).value = True, 1, 0), Trim(txtApp.Text), Trim(txtPesan.Text), NulSID(rsRem.Fields("reminder_id").value))
         
2           Rec = True
3           If ChUlang.value = Unchecked And OptTelp(1).value = True And Not objR.pasifCustomer(Trim(txtNoChassis)) And flagTdkDiundang = False Then
              Dim heavy As Boolean
              Dim FSRC As String
              Dim NomorWO As String
              Dim noInv As String
4             heavy = IIf((CDbl(rsRem.Fields(5)) - CDbl(rsRem.Fields(3))) = 5000, 1, 0)
5             NomorWO = ""
6             noInv = ""
              Dim Xes As Integer
7             Xes = objO.FetchData_Count("Select NOWO from M_Estimasi WITH(NOLOCK) Where NOWO = '" & noWO & "'")
8             If Xes < 3 Then
9                 If is1000 Then
10                    FSRC = 1000
11                    NomorWO = Trim(rsRem.Fields(0))
12                    noInv = ""
13                Else
14                    If Trim(rsRem.Fields(19)) = "R001" Then
15                        FSRC = IIf(heavy, 5, 10) * 1000
16                    Else
                          'FSRC = (CDbl(Right(rsRem.Fields(19), 3)) + IIf(heavy, 5, 10)) * 1000
17                        FSRC = Val(rsRem.Fields(20).value) '+ IIf(heavy, 5, 10) * 1000
18                    End If
19                    NomorWO = Trim(rsRem.Fields(1))
20                    noInv = Trim(rsRem.Fields(0))
21                End If
                  'Call objO.UpdateData("EXEC NEXTESTIMASI '" & Trim(txtNoChassis) & "','" & Format(DateAdd("m", IIf(is1000, 1, IIf(heavy, 3, 6)), rsRem.Fields(2)), "MM/DD/YYYY") & "','" & CDbl(rsRem.Fields(3)) + IIf(is1000, 0, IIf(heavy, 5000, 10000)) & "','" & CDbl(FSRC) & "','" & IIf(heavy, 5000, 10000) & "','" & Trim(noinv) & "','" & Trim(NomorWO) & "','" & Trim(rsRem.Fields(6)) & "','" & Trim(rsRem.Fields(7)) & "','" & Trim(rsRem.Fields(10)) & "','" & Trim(txtTelp) & "','" & Trim(txtGSM) & "','" & Trim(txtCDMA) & "'")
                  'Call objO.UpdateData("EXEC NEXTESTIMASI '" & Trim(txtNoChassis) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & CDbl(rsRem.Fields(3)) + IIf(is1000, 0, IIf(heavy, 5000, 10000)) & "','" & CDbl(FSRC) & "','" & IIf(heavy, 5000, 10000) & "','" & Trim(noinv) & "','" & Trim(NomorWO) & "','" & Trim(rsRem.Fields(6)) & "','" & Trim(rsRem.Fields(7)) & "','" & Trim(rsRem.Fields(10)) & "','" & Trim(txtTelp) & "','" & Trim(txtGSM) & "','" & Trim(txtCDMA) & "'")
22                If isHansaSA = False Then
23                    Call objO.UpdateData("EXEC " & Trim(hTipePenghitunganDealer) & " '" & Trim(txtNoChassis) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & CDbl(rsRem.Fields(3)) + IIf(is1000, 0, IIf(heavy, 5000, 10000)) & "','" & CDbl(FSRC) & "','" & IIf(heavy, 5000, 10000) & "','" & Trim(noInv) & "','" & Trim(NomorWO) & "','" & Trim(rsRem.Fields(6)) & "','" & Trim(rsRem.Fields(7)) & "','" & Trim(rsRem.Fields(10)) & "','" & Trim(txtTelp) & "','" & Trim(txtGSM) & "','" & Trim(txtCDMA) & "'")
24                End If
25            End If
26          End If
            
            'Update data Remainder SMS dan DM
            Dim rsAllRem As ADODB.Recordset
27          Set rsAllRem = objR.Fetch_SMSnDM_Remainder(noWO)
28          If rsAllRem.RecordCount = 0 Then
29            If MsgBox("Keluar dari Remainder Call ?", vbQuestion + vbYesNo) = vbYes Then
30               Unload Me
31            End If
32            Exit Sub
33          End If
34          rsAllRem.MoveFirst
35          Do Until rsAllRem.EOF
36             If Trim(rsAllRem.Fields(0).value) = "SM" Then
37                Call objR.UpdateReminder(noWO, "SM", IIf(OptSMS(1).value = True, 1, IIf(OptSMS(0).value = True, 0, 2)), DateDiff("d", Format(rsAllRem.Fields(1).value, "dd/mm/yyyy"), Format(isnow, "dd/mm/yyyy")), Val(txtHargaJasa.Text), Val(txtHargaPart.Text), Val(cboJam.Text), Val(cboMenit.Text))
38             ElseIf Trim(rsAllRem.Fields(0).value) = "IK" Then
39                Call objR.UpdateReminder(noWO, "1K", IIf(OptSMS(1).value = True, 1, IIf(OptSMS(0).value = True, 0, 2)), DateDiff("d", Format(rsAllRem.Fields(1).value, "dd/mm/yyyy"), Format(isnow, "dd/mm/yyyy")), Val(txtHargaJasa.Text), Val(txtHargaPart.Text), Val(cboJam.Text), Val(cboMenit.Text))
40             ElseIf Trim(rsAllRem.Fields(0).value) = "DM" Then
41                Call objR.UpdateReminder(noWO, "DM", IIf(OptSMS(1).value = True, 1, IIf(OptSMS(0).value = True, 0, 2)), DateDiff("d", Format(rsAllRem.Fields(1).value, "dd/mm/yyyy"), Format(isnow, "dd/mm/yyyy")), Val(txtHargaJasa.Text), Val(txtHargaPart.Text), Val(cboJam.Text), Val(cboMenit.Text))
42             End If
43             rsAllRem.MoveNext
44          Loop
45          If MsgBox("Keluar dari Remainder Call ?", vbQuestion + vbYesNo) = vbYes Then
46             Unload Me
47          End If

   On Error GoTo 0
   Exit Sub

simpanHasil_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure simpanHasil of Form RemainderCallRetention"
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

Private Function estimasiJasa(nKm As String, nTipeKend As String, nRangka As String) As Double
      Dim rsHarga As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim nRate As Double
          
1         estimasiJasa = 0
2         nRate = 0
3         Set rsSearch = objO.FetchData("SELECT * FROM HO_FLATE WITH(NOLOCK) WHERE MDL = '" & Trim(nTipeKend) & "' AND FS_ITEM LIKE '%" & Format(nKm, "##,0") & "%'")
4         If Not rsSearch.EOF Then
5             nRate = CDbl(rsSearch("RATE"))
6         End If

7         Select Case nKm
              Case "1000"
8                 estimasiJasa = 0
9             Case Else
                  
10                If Log.RPLHistori = 0 Then
11                    Set rsHarga = objO.FetchData("SELECT * FROM MODEL A WITH(NOLOCK) LEFT JOIN CARCLASS B WITH(NOLOCK) ON A.MODEL = B.CLASSCODE WHERE MDL = '" & Trim(nTipeKend) & "'")
12                Else
13                    Set rsHarga = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..CARCLASS B WITH(NOLOCK) ON A.MODEL = B.CLASSCODE WHERE MDL = '" & Trim(nTipeKend) & "'")
14                End If
15                If Not rsHarga.EOF Then
16                    Select Case dealerJual(Trim(nRangka))
                          Case 0
                              ' EXTERN
17                            estimasiJasa = CDbl(nRate) * CDbl(rsHarga("PRICETAM"))
18                        Case 1
                              ' INTERN
19                            estimasiJasa = CDbl(nRate) * CDbl(rsHarga("PRICE"))
20                    End Select
21                End If
22        End Select
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

Private Sub simpanAlasan_Click()
Dim boleh As Boolean
Dim pesan As String

    boleh = True
    If OptAlasan(3).value Then
        If Trim(txtAlasanLain) = "" Then
            boleh = False
            pesan = "Isi alasannya !!!"
        End If
    End If
    
    If boleh Then
        If OptAlasan(0).value Then
            txtKetTelp = OptAlasan(0).Caption & " : " & Trim(txtAlasanLain)
        Else
            If OptAlasan(1).value Then
                txtKetTelp = OptAlasan(1).Caption & " : " & Trim(txtAlasanLain)
            Else
                If OptAlasan(2).value Then
                    txtKetTelp = OptAlasan(2).Caption & " : " & Trim(txtAlasanLain)
                Else
                    txtKetTelp = "4. " & Trim(txtAlasanLain)
                End If
            End If
        End If
        Call simpanHasil
    Else
        MsgBox pesan
    End If
End Sub

Private Function tampilLebihKecil(date1 As Date, date2 As Date) As Date

    tampilLebihKecil = IIf(DateDiff("d", date1, date2) > 0, date1, date2)
End Function

Private Sub panggilWarranty(noRangka As String, modL As String)
      Dim rsRec As ADODB.Recordset
      Dim terlibat As Boolean
      Dim rsTemp    As ADODB.Recordset
      '' 20180606DEV01'NOTIFSCC
      Dim warranty As String
      Dim No As Integer
1        On Error GoTo panggilWarranty_error
 
          ''20180704 AGAR WARRANTY TIDAK DOUBLE2
          
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT DISTINCT A.ID_WARRANTY,a.warranty FROM ho_warranty A WITH (NOLOCK) INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
                           & "WHERE no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
                           & "ORDER BY A.id_warranty")
4         Else
5             Set rsTemp = objO.FetchData("SELECT DISTINCT A.ID_WARRANTY,a.warranty FROM [" & dbReplikasiHO & "]..ho_warranty A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
                           & "WHERE no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
                           & "ORDER BY A.id_warranty")
6         End If
          
7         terlibat = False
8         No = 1
9         If Not rsTemp.EOF Then
10            Do While Not rsTemp.EOF
              
                  ''Set rsRec = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(noRangka) & "' AND KODETIPE = '" & NulSID(rsTemp("ID_WARRANTY")) & "' AND VALIDTWC <> 'NOK'")
                  ''20200316 || pergantian req mb thania
11                Set rsRec = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(noRangka) & "' AND KODETIPE = '" & NulSID(rsTemp("ID_WARRANTY")) & "' AND VALIDTWC <> ''")
12                If Not rsRec.EOF Then
13                    terlibat = False
14                    Exit Do
15                Else
16                    terlibat = True
17                End If
                  
18                Set rsRec = objO.FetchData("SELECT NO_RANGKA FROM TEMPWAR WITH (NOLOCK) WHERE NO_RANGKA = '" & noRangka & "' AND ID_WARRANTY = '" & rsTemp("ID_WARRANTY") & "'")
19                If Not rsRec.EOF Then
20                    terlibat = False
21                    Exit Do
22                Else
23                    terlibat = True
24                End If
                  
25                Set rsRec = objO.FetchData("SELECT no_chassis FROM History WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
26                If Not rsRec.EOF Then
27                    terlibat = False
28                    Exit Do
29                Else
30                    terlibat = True
31                End If
                  
32                If Log.RPLHistori = 1 Then
33                    Set rsRec = objO.FetchData("SELECT no_chassis FROM [" & dbReplikasiHO & "]..Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
34                    If Not rsRec.EOF Then
35                        terlibat = False
36                        Exit Do
37                    Else
38                        terlibat = True
39                    End If
40                Else
41                    If Log.LogHOConn = 1 Then
42                        If ConHO = True Then
                              
43                            Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
44                            If Not rsRec.EOF Then
45                                terlibat = False
46                                Exit Do
47                            Else
48                                terlibat = True
49                            End If
50                        End If
51                    End If
52                End If
                  
53                warranty = warranty & No & ". " & NulSID(rsTemp!warranty) & vbCrLf
                        
54                No = No + 1
55                rsTemp.MoveNext
56            Loop
              '' 20180606DEV01'NOTIFSCC
57            If terlibat Then
58                MsgBox "UNIT TERLIBAT WARRANTY !!!" & vbCrLf & warranty, vbCritical, "UNIT SSC"
59                ';cmdWarranty.BackColor = &HFF&
                  '
60                ';cmdWarranty.Enabled = True
61                'opTpSvc(2).Enabled = True
62            End If
63        Else
64            'cmdWarranty.BackColor = &HC0C0C0
65            'cmdWarranty.Enabled = False
66        End If

         'On Error GoTo 0
67       Exit Sub

panggilWarranty_error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
68    varErlError = Erl
69    varErrNumber = err.Number
      '' varErlError,varErrNumber
70        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure panggilWarranty of Form RemainderCall")
End Sub




