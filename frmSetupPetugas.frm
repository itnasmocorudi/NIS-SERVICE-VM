VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Begin VB.Form frmSetupPetugas 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Setup Data Petugas"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   435
      TabIndex        =   54
      Top             =   345
      Width           =   14580
      Begin VB.OptionButton opDiv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BP"
         Height          =   255
         Index           =   1
         Left            =   9360
         TabIndex        =   6
         Top             =   848
         Width           =   735
      End
      Begin VB.OptionButton opDiv 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GR"
         Height          =   255
         Index           =   0
         Left            =   8640
         TabIndex        =   5
         Top             =   848
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkHO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HO"
         Height          =   285
         Left            =   8640
         TabIndex        =   99
         Top             =   533
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox cboPetugas 
         Height          =   315
         ItemData        =   "frmSetupPetugas.frx":0000
         Left            =   9600
         List            =   "frmSetupPetugas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   540
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   2
         Left            =   7305
         TabIndex        =   53
         Top             =   210
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         ForeColor       =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   3
         Left            =   3720
         TabIndex        =   45
         Top             =   840
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   6
         Left            =   6645
         TabIndex        =   3
         Top             =   210
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   210
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   4
         Left            =   6645
         TabIndex        =   4
         Top             =   540
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   5
         Left            =   6645
         TabIndex        =   227
         Top             =   840
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NIK"
         Height          =   195
         Left            =   6165
         TabIndex        =   228
         Top             =   840
         Width           =   270
      End
      Begin VB.Shape Shape1 
         Height          =   270
         Left            =   8520
         Top             =   540
         Width           =   45
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
         Height          =   195
         Left            =   6150
         TabIndex        =   68
         Top             =   540
         Width           =   315
      End
      Begin VB.Shape Shape3 
         Height          =   270
         Left            =   7215
         Top             =   210
         Width           =   45
      End
      Begin VB.Shape Shape2 
         Height          =   270
         Left            =   2505
         Top             =   210
         Width           =   45
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   1140
         TabIndex        =   58
         Top             =   885
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Akses Setara Dengan"
         Height          =   195
         Left            =   4965
         TabIndex        =   57
         Top             =   210
         Width           =   1560
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Petugas/ Group"
         Height          =   195
         Left            =   240
         TabIndex        =   56
         Top             =   615
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Petugas/ Group "
         Height          =   195
         Left            =   270
         TabIndex        =   55
         Top             =   285
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   7050
      Left            =   435
      TabIndex        =   59
      Top             =   1440
      Width           =   14580
      Begin VB.Frame FrameHO 
         BackColor       =   &H00FFFFFF&
         Height          =   2055
         Index           =   4
         Left            =   10600
         TabIndex        =   118
         Top             =   4900
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CheckBox mnu1_25_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Perbaikan No Rangka"
            Height          =   285
            Left            =   480
            TabIndex        =   124
            Top             =   360
            Width           =   2445
         End
         Begin VB.CheckBox mnu1_25_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Perbaikan KM Service"
            Height          =   285
            Left            =   480
            TabIndex        =   123
            Top             =   600
            Width           =   2325
         End
         Begin VB.CheckBox mnu1_25_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Perbaikan Nama Mekanik"
            Height          =   195
            Left            =   480
            TabIndex        =   122
            Top             =   870
            Width           =   2565
         End
         Begin VB.CheckBox mnu1_25_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Perbaikan Job Order SBE"
            Height          =   285
            Left            =   480
            TabIndex        =   121
            Top             =   1560
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.CheckBox mnu1_25_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Perbaikan Panggilan selesai Service"
            Height          =   285
            Left            =   480
            TabIndex        =   120
            Top             =   1080
            Width           =   3165
         End
         Begin VB.CheckBox mnu1_221_25 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Perbaikan Data"
            Height          =   285
            Left            =   120
            TabIndex        =   119
            Top             =   120
            Width           =   1845
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Laporan"
         Height          =   6855
         Index           =   9
         Left            =   7440
         TabIndex        =   69
         Top             =   120
         Width           =   3100
         Begin VB.CheckBox mnu1_10_8_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "History Service"
            Height          =   285
            Left            =   1560
            TabIndex        =   229
            Top             =   4800
            Width           =   1485
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   120
            TabIndex        =   125
            Top             =   2640
            Width           =   2895
            Begin VB.CheckBox mnu1_10_6_8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Konfirmasi Cuci"
               Height          =   285
               Left            =   300
               TabIndex        =   133
               Top             =   1700
               Width           =   2325
            End
            Begin VB.CheckBox mnu1_10_6_7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Cuci"
               Height          =   285
               Left            =   300
               TabIndex        =   132
               Top             =   1500
               Width           =   2325
            End
            Begin VB.CheckBox mnu1_10_6_2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Bulanan"
               Height          =   285
               Left            =   300
               TabIndex        =   131
               Top             =   600
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_6_1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Harian"
               Height          =   285
               Left            =   300
               TabIndex        =   130
               Top             =   360
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_6 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Rekap"
               Height          =   195
               Left            =   0
               TabIndex        =   129
               Top             =   120
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_6_3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan FIR Follow Up"
               Height          =   285
               Left            =   300
               TabIndex        =   128
               Top             =   825
               Width           =   2085
            End
            Begin VB.CheckBox mnu1_10_6_4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Konfirmasi Sublet"
               Height          =   285
               Left            =   300
               TabIndex        =   127
               Top             =   1035
               Width           =   2325
            End
            Begin VB.CheckBox mnu1_10_6_5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Chasing Form"
               Height          =   285
               Left            =   300
               TabIndex        =   126
               Top             =   1280
               Width           =   2325
            End
         End
         Begin VB.CheckBox mnu1_10_8_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Supply Slip"
            Height          =   285
            Left            =   360
            TabIndex        =   116
            Top             =   4800
            Width           =   1125
         End
         Begin VB.CheckBox mnu1_10_8_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Invoice Retur"
            Height          =   285
            Left            =   360
            TabIndex        =   115
            Top             =   5040
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_10_8_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MOU"
            Height          =   285
            Left            =   360
            TabIndex        =   114
            Top             =   5280
            Width           =   1485
         End
         Begin VB.CheckBox mnu1_10_8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Laporan Monitoring"
            Height          =   285
            Left            =   120
            TabIndex        =   113
            Top             =   4560
            Width           =   1845
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   120
            TabIndex        =   108
            Top             =   5280
            Width           =   2900
            Begin VB.CheckBox mnu1_10_7_4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Mapping"
               Height          =   195
               Left            =   240
               TabIndex        =   117
               Top             =   1200
               Width           =   2445
            End
            Begin VB.CheckBox mnu1_10_7_3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "List Unit Belum SBI"
               Height          =   285
               Left            =   240
               TabIndex        =   112
               Top             =   960
               Width           =   2445
            End
            Begin VB.CheckBox mnu1_10_7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan MRS"
               Height          =   285
               Left            =   0
               TabIndex        =   111
               Top             =   240
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_7_1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Early"
               Height          =   285
               Left            =   240
               TabIndex        =   110
               Top             =   480
               Width           =   1365
            End
            Begin VB.CheckBox mnu1_10_7_2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Telp tidak terhubung"
               Height          =   285
               Left            =   240
               TabIndex        =   109
               Top             =   720
               Width           =   2445
            End
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   101
            Top             =   1920
            Width           =   2900
            Begin VB.CheckBox mnu1_10_5_2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Harian Mekanik"
               Height          =   285
               Left            =   300
               TabIndex        =   107
               Top             =   240
               Width           =   2205
            End
            Begin VB.CheckBox mnu1_10_5_1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Produktifitas Mekanik"
               Height          =   285
               Left            =   300
               TabIndex        =   24
               Top             =   480
               Visible         =   0   'False
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Mekanik"
               Height          =   285
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Visible         =   0   'False
               Width           =   2685
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   120
            TabIndex        =   97
            Top             =   1200
            Width           =   2535
            Begin VB.CheckBox mnu1_10_4_2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Total Revenue"
               Height          =   285
               Left            =   300
               TabIndex        =   23
               Top             =   240
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_4 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Revenue Unit"
               Height          =   285
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_4_1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Rincian Revenue"
               Height          =   285
               Left            =   300
               TabIndex        =   22
               Top             =   480
               Width           =   2685
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   2655
            Begin VB.CheckBox mnu1_10_3_3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Jumlah Unit Out"
               Height          =   285
               Left            =   300
               TabIndex        =   134
               Top             =   720
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_3_2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan SO dalam Proses"
               Height          =   285
               Left            =   300
               TabIndex        =   21
               Top             =   480
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_3_1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Unit Masuk"
               Height          =   285
               Left            =   300
               TabIndex        =   20
               Top             =   240
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_3 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Laporan Unit Entry"
               Height          =   285
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   2685
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Laporan"
            Height          =   1860
            Left            =   4440
            TabIndex        =   89
            Top             =   4320
            Visible         =   0   'False
            Width           =   3660
            Begin VB.CheckBox mnu1_10_1_1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Semua SO Masuk"
               Height          =   285
               Left            =   660
               TabIndex        =   95
               Top             =   450
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_1_2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Unit Masuk/SA"
               Height          =   285
               Left            =   660
               TabIndex        =   94
               Top             =   690
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_1_3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Semua SO dalam Proses"
               Height          =   285
               Left            =   660
               TabIndex        =   93
               Top             =   930
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_1_5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Unit"
               Height          =   285
               Left            =   660
               TabIndex        =   92
               Top             =   1440
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_1_4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Audit Bengkel"
               Height          =   285
               Left            =   660
               TabIndex        =   91
               Top             =   1170
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Order"
               Height          =   285
               Left            =   360
               TabIndex        =   90
               Top             =   210
               Width           =   2685
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0E0E0&
            Height          =   4710
            Left            =   10200
            TabIndex        =   70
            Top             =   600
            Visible         =   0   'False
            Width           =   3540
            Begin VB.CheckBox mnu1_10_2_4 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Penjualan"
               Height          =   285
               Left            =   645
               TabIndex        =   88
               Top             =   1200
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Penjualan/SA"
               Height          =   285
               Left            =   645
               TabIndex        =   87
               Top             =   960
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Penjualan &Finance"
               Height          =   285
               Left            =   645
               TabIndex        =   86
               Top             =   480
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Penjualan &Service"
               Height          =   285
               Left            =   645
               TabIndex        =   85
               Top             =   720
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_8 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cust. Pasif IRC 5000 /Hari"
               Height          =   285
               Left            =   645
               TabIndex        =   84
               Top             =   2160
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_7 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cust. Pasif IRC 1000 /Hari"
               Height          =   285
               Left            =   645
               TabIndex        =   83
               Top             =   1920
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_5 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Harian Detail Penjualan"
               Height          =   285
               Left            =   645
               TabIndex        =   82
               Top             =   1440
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_6 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cust. Pasif/Hari"
               Height          =   285
               Left            =   645
               TabIndex        =   81
               Top             =   1680
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_11 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bulan Regular Check"
               Height          =   285
               Left            =   645
               TabIndex        =   80
               Top             =   2880
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_9 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bulan Penjualan"
               Height          =   285
               Left            =   645
               TabIndex        =   79
               Top             =   2400
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_10 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bulan Detail Penjualan"
               Height          =   285
               Left            =   645
               TabIndex        =   78
               Top             =   2640
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Billing"
               Height          =   285
               Left            =   360
               TabIndex        =   77
               Top             =   225
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_16 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bulan Regular Check per No Polisi"
               Height          =   285
               Left            =   645
               TabIndex        =   76
               Top             =   3105
               Width           =   2970
            End
            Begin VB.CheckBox mnu1_10_2_12 
               BackColor       =   &H00E0E0E0&
               Caption         =   "F. Rate Group"
               Height          =   285
               Left            =   645
               TabIndex        =   75
               Top             =   3570
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_15 
               BackColor       =   &H00E0E0E0&
               Caption         =   "IRC 5000 pada Bulan"
               Height          =   255
               Left            =   645
               TabIndex        =   74
               Top             =   4290
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_13 
               BackColor       =   &H00E0E0E0&
               Caption         =   "F. Rate Mekanik"
               Height          =   285
               Left            =   645
               TabIndex        =   73
               Top             =   3810
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_14 
               BackColor       =   &H00E0E0E0&
               Caption         =   "IRC 1000 pada Bulan"
               Height          =   285
               Left            =   645
               TabIndex        =   72
               Top             =   4050
               Width           =   2685
            End
            Begin VB.CheckBox mnu1_10_2_17 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Bulan Penjualan/SA"
               Height          =   285
               Left            =   645
               TabIndex        =   71
               Top             =   3330
               Width           =   2685
            End
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Progress"
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   3960
         Width           =   3600
         Begin VB.CheckBox mnu1_3_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "List Job Mekanik"
            Height          =   285
            Left            =   120
            TabIndex        =   106
            Top             =   400
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_3_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Display Customer"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_3_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Progress"
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   200
            Width           =   3045
         End
         Begin VB.CheckBox mnu1_3_1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pengisian Progress"
            Height          =   285
            Left            =   375
            TabIndex        =   48
            Top             =   195
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.CheckBox mnu1_3_2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Display Aktivity Mekanik"
            Height          =   285
            Left            =   360
            TabIndex        =   49
            Top             =   1440
            Visible         =   0   'False
            Width           =   2685
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Antrian"
         Height          =   885
         Index           =   10
         Left            =   120
         TabIndex        =   105
         Top             =   1560
         Width           =   3600
         Begin VB.CheckBox mnu1_12_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cetak Antrian"
            Height          =   285
            Left            =   135
            TabIndex        =   7
            Top             =   195
            Width           =   1365
         End
         Begin VB.CheckBox mnu1_12_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Panggil Antrian"
            Height          =   285
            Left            =   135
            TabIndex        =   8
            Top             =   480
            Width           =   2685
         End
      End
      Begin VB.Frame FrameHO 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   3
         Left            =   10600
         TabIndex        =   104
         Top             =   3720
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CheckBox mnu1_24_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ganti Password"
            Height          =   285
            Left            =   480
            TabIndex        =   44
            Top             =   840
            Width           =   1485
         End
         Begin VB.CheckBox mnu1_24_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Profil"
            Height          =   285
            Left            =   480
            TabIndex        =   42
            Top             =   360
            Width           =   1365
         End
         Begin VB.CheckBox mnu1_24_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Petugas"
            Height          =   285
            Left            =   480
            TabIndex        =   43
            Top             =   600
            Width           =   1485
         End
         Begin VB.CheckBox mnu1_221_24 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Setting"
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   2565
         End
      End
      Begin VB.Frame FrameHO 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   2
         Left            =   10600
         TabIndex        =   103
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CheckBox mnu1_221_23 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Master Laporan"
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   2565
         End
         Begin VB.CheckBox mnu1_23_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Data Karyawan"
            Height          =   285
            Left            =   480
            TabIndex        =   41
            Top             =   600
            Width           =   1485
         End
         Begin VB.CheckBox mnu1_23_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Target Cabang"
            Height          =   285
            Left            =   480
            TabIndex        =   40
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame FrameHO 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Index           =   1
         Left            =   10600
         TabIndex        =   102
         Top             =   1780
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CheckBox mnu1_22_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bulanan"
            Height          =   285
            Left            =   480
            TabIndex        =   35
            Top             =   360
            Width           =   1125
         End
         Begin VB.CheckBox mnu1_22_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SMPS"
            Height          =   285
            Left            =   480
            TabIndex        =   36
            Top             =   600
            Width           =   1245
         End
         Begin VB.CheckBox mnu1_22_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "WMR"
            Height          =   195
            Left            =   480
            TabIndex        =   37
            Top             =   870
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_22_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MRS"
            Height          =   285
            Left            =   1920
            TabIndex        =   38
            Top             =   360
            Width           =   1005
         End
         Begin VB.CheckBox mnu1_22_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "WARRANTY"
            Height          =   285
            Left            =   1920
            TabIndex        =   39
            Top             =   600
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_221_22 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Laporan HO"
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   2565
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Billing"
         Height          =   1750
         Index           =   3
         Left            =   3720
         TabIndex        =   64
         Top             =   5160
         Width           =   3660
         Begin VB.VScrollBar VScrollBilling 
            Height          =   1455
            LargeChange     =   100
            Left            =   3360
            Max             =   3500
            TabIndex        =   216
            Top             =   240
            Width           =   300
         End
         Begin VB.PictureBox Picture6 
            Height          =   1455
            Left            =   0
            ScaleHeight     =   1395
            ScaleWidth      =   3315
            TabIndex        =   215
            Top             =   240
            Width           =   3375
            Begin VB.Frame FrameBilling 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   4815
               Left            =   0
               TabIndex        =   217
               Top             =   0
               Width           =   3375
               Begin VB.CheckBox mnu1_4_8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Rekap Tagihan"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   226
                  Top             =   525
                  Width           =   1920
               End
               Begin VB.CheckBox mnu1_4_2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pembatalan Invoice"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   225
                  Top             =   240
                  Width           =   1725
               End
               Begin VB.CheckBox mnu1_4_1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pembuatan Invoice"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   224
                  Top             =   0
                  Width           =   1725
               End
               Begin VB.CheckBox mnu1_4_11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "List Piutang (AR)"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   223
                  Top             =   1075
                  Width           =   1605
               End
               Begin VB.CheckBox mnu1_4_10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Faktur Pajak"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   222
                  Top             =   1375
                  Width           =   1245
               End
               Begin VB.CheckBox mnu1_4_9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Cetak Ulang"
                  Height          =   405
                  Left            =   120
                  TabIndex        =   221
                  Top             =   720
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_4_6 
                  BackColor       =   &H000000FF&
                  Caption         =   "Perubahan A.Rate, Group && Mekanik"
                  Height          =   300
                  Left            =   120
                  TabIndex        =   220
                  Top             =   4320
                  Visible         =   0   'False
                  Width           =   3120
               End
               Begin VB.CheckBox mnu1_4_4 
                  BackColor       =   &H000000FF&
                  Caption         =   "Batal Pencatatan Kredit"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   219
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   2685
               End
               Begin VB.CheckBox mnu1_4_3 
                  BackColor       =   &H000000FF&
                  Caption         =   "Pencatatan Kredit"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   218
                  Top             =   3840
                  Visible         =   0   'False
                  Width           =   2685
               End
            End
         End
      End
      Begin VB.Frame FrameHO 
         BackColor       =   &H00FFFFFF&
         Height          =   1695
         Index           =   0
         Left            =   10600
         TabIndex        =   100
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CheckBox mnu1_221_21 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Master HO"
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   1125
         End
         Begin VB.CheckBox mnu1_21_9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Model"
            Height          =   285
            Left            =   2160
            TabIndex        =   34
            Top             =   1320
            Width           =   1005
         End
         Begin VB.CheckBox mnu1_21_8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Flate Rate"
            Height          =   285
            Left            =   2160
            TabIndex        =   33
            Top             =   1080
            Width           =   1245
         End
         Begin VB.CheckBox mnu1_21_7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Package Part"
            Height          =   285
            Left            =   2160
            TabIndex        =   32
            Top             =   840
            Width           =   1485
         End
         Begin VB.CheckBox mnu1_21_6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dealer"
            Height          =   285
            Left            =   2160
            TabIndex        =   31
            Top             =   600
            Width           =   1005
         End
         Begin VB.CheckBox mnu1_21_5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Warranty"
            Height          =   285
            Left            =   2160
            TabIndex        =   30
            Top             =   360
            Width           =   1125
         End
         Begin VB.CheckBox mnu1_21_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Program Khusus"
            Height          =   285
            Left            =   480
            TabIndex        =   29
            Top             =   1320
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_21_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MOU"
            Height          =   285
            Left            =   480
            TabIndex        =   28
            Top             =   1080
            Width           =   1005
         End
         Begin VB.CheckBox mnu1_21_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Kelas Kendaraan"
            Height          =   195
            Left            =   480
            TabIndex        =   27
            Top             =   870
            Width           =   1605
         End
         Begin VB.CheckBox mnu1_21_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Job Order"
            Height          =   285
            Left            =   480
            TabIndex        =   26
            Top             =   600
            Width           =   1245
         End
         Begin VB.CheckBox mnu1_21_10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Group Job"
            Height          =   285
            Left            =   480
            TabIndex        =   25
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reminder System"
         Height          =   1050
         Index           =   5
         Left            =   3840
         TabIndex        =   98
         Top             =   960
         Width           =   3660
         Begin VB.VScrollBar VScroll1 
            Height          =   810
            LargeChange     =   100
            Left            =   3240
            Max             =   3500
            TabIndex        =   238
            Top             =   240
            Width           =   300
         End
         Begin VB.PictureBox Picture7 
            Height          =   735
            Left            =   50
            ScaleHeight     =   675
            ScaleWidth      =   3315
            TabIndex        =   230
            Top             =   240
            Width           =   3375
            Begin VB.Frame FrameReminder 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   1815
               Left            =   0
               TabIndex        =   231
               Top             =   0
               Width           =   3495
               Begin VB.CheckBox mnu1_13_6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Hasil Retensi Asuransi"
                  Height          =   345
                  Left            =   0
                  TabIndex        =   237
                  Top             =   1200
                  Width           =   2445
               End
               Begin VB.CheckBox mnu1_13_5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Retensi Asuransi"
                  Height          =   225
                  Left            =   0
                  TabIndex        =   236
                  Top             =   960
                  Width           =   1725
               End
               Begin VB.CheckBox mnu1_13_4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Inject Data To BP"
                  Height          =   225
                  Left            =   0
                  TabIndex        =   235
                  Top             =   720
                  Width           =   1725
               End
               Begin VB.CheckBox mnu1_13_1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Menu MRS"
                  Height          =   225
                  Left            =   0
                  TabIndex        =   234
                  Top             =   0
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_13_2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "SA Retention"
                  Height          =   225
                  Left            =   0
                  TabIndex        =   233
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_13_3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Inject Data SA to Sales"
                  Height          =   225
                  Left            =   0
                  TabIndex        =   232
                  Top             =   480
                  Width           =   1965
               End
            End
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Setting"
         Height          =   1695
         Index           =   8
         Left            =   3720
         TabIndex        =   66
         Top             =   3480
         Width           =   3660
         Begin VB.VScrollBar VScrollSetting 
            Height          =   1455
            LargeChange     =   100
            Left            =   3360
            Max             =   3500
            TabIndex        =   203
            Top             =   240
            Width           =   300
         End
         Begin VB.PictureBox Picture5 
            Height          =   1455
            Left            =   0
            ScaleHeight     =   1395
            ScaleWidth      =   3315
            TabIndex        =   202
            Top             =   240
            Width           =   3375
            Begin VB.Frame frameSetting 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   4095
               Left            =   0
               TabIndex        =   204
               Top             =   0
               Width           =   3375
               Begin VB.CheckBox mnu1_6_10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Setup"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   214
                  Top             =   1180
                  Width           =   1005
               End
               Begin VB.CheckBox mnu1_6_8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Profil AAS"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   213
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.CheckBox mnu1_6_5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ganti No. Polisi"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   212
                  Top             =   970
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_6_3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Backup"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   211
                  Top             =   705
                  Width           =   1125
               End
               Begin VB.CheckBox mnu1_6_1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Profil"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   210
                  Top             =   0
                  Width           =   1005
               End
               Begin VB.CheckBox mnu1_6_2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Petugas"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   209
                  Top             =   465
                  Width           =   1005
               End
               Begin VB.CheckBox mnu1_6_11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Maintenance"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   208
                  Top             =   2160
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_6_4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ganti Password"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   207
                  Top             =   1920
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_6_7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Kode Jurnal"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   206
                  Top             =   1680
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_6_9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Kode Oli"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   205
                  Top             =   1440
                  Width           =   1485
               End
            End
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Booking"
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   3600
         Begin VB.PictureBox Picture1 
            Height          =   1215
            Left            =   0
            ScaleHeight     =   1155
            ScaleWidth      =   3195
            TabIndex        =   136
            Top             =   240
            Width           =   3255
            Begin VB.Frame frameBooking 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   2055
               Left            =   0
               TabIndex        =   137
               Top             =   0
               Width           =   3255
               Begin VB.CheckBox mnu1_1_4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Display Stall"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   143
                  Top             =   1080
                  Width           =   1845
               End
               Begin VB.CheckBox mnu1_1_5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Jam Tersedia"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   142
                  Top             =   795
                  Width           =   1845
               End
               Begin VB.CheckBox mnu1_1_6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Booking Online"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   141
                  Top             =   1320
                  Width           =   1965
               End
               Begin VB.CheckBox mnu1_1_1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pengisian Booking"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   140
                  Top             =   0
                  Width           =   1725
               End
               Begin VB.CheckBox mnu1_1_2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pembatalan Booking"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   139
                  Top             =   285
                  Width           =   1845
               End
               Begin VB.CheckBox mnu1_1_3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pembuatan Estimasi"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   138
                  Top             =   525
                  Width           =   1845
               End
            End
         End
         Begin VB.VScrollBar vscrollBooking 
            Height          =   1255
            LargeChange     =   100
            Left            =   3275
            Max             =   500
            TabIndex        =   135
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Order"
         Height          =   1500
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   2450
         Width           =   3600
         Begin VB.VScrollBar VScrollOrder 
            Height          =   1215
            LargeChange     =   100
            Left            =   3240
            Max             =   3500
            TabIndex        =   145
            Top             =   240
            Width           =   300
         End
         Begin VB.PictureBox Picture2 
            Height          =   1215
            Left            =   0
            ScaleHeight     =   1155
            ScaleWidth      =   3195
            TabIndex        =   144
            Top             =   240
            Width           =   3255
            Begin VB.Frame frameOrder 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   5175
               Left            =   0
               TabIndex        =   146
               Top             =   0
               Width           =   3235
               Begin VB.CheckBox mnu1_2_8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Validasi TWC"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   163
                  Top             =   1680
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Retur Sublet"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   162
                  Top             =   1440
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pemakaian Sublet"
                  Height          =   405
                  Left            =   135
                  TabIndex        =   161
                  Top             =   1080
                  Width           =   2685
               End
               Begin VB.CheckBox mnu1_2_5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Validasi Order"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   160
                  Top             =   885
                  Width           =   2685
               End
               Begin VB.CheckBox mnu1_2_4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "3 Day's Follow Up"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   159
                  Top             =   625
                  Width           =   2685
               End
               Begin VB.CheckBox mnu1_2_1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pembuatan Order"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   158
                  Top             =   120
                  Width           =   2685
               End
               Begin VB.CheckBox mnu1_2_2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pembatalan Order"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   157
                  Top             =   345
                  Width           =   2685
               End
               Begin VB.CheckBox mnu1_2_15 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Complain Order"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   156
                  Top             =   3360
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_14 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Support TL"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   155
                  Top             =   3120
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_13 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "List Cuci"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   154
                  Top             =   2880
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_12 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Menu Pengering"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   153
                  Top             =   2640
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Menu Cuci"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   152
                  Top             =   2400
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "List Unit Come"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   151
                  Top             =   2160
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Penerimaan Unit"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   150
                  Top             =   1920
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_16 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "List Sublet"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   149
                  Top             =   3600
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_17 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Validasi Customer"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   148
                  Top             =   3840
                  Width           =   2565
               End
               Begin VB.CheckBox mnu1_2_18 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ganti Unit Customer"
                  Height          =   285
                  Left            =   135
                  TabIndex        =   147
                  Top             =   4080
                  Width           =   2565
               End
            End
         End
         Begin VB.CheckBox mnu1_2_3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Display All Repair"
            Height          =   285
            Left            =   3840
            TabIndex        =   47
            Top             =   600
            Visible         =   0   'False
            Width           =   1965
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informasi"
         Height          =   1750
         Index           =   6
         Left            =   120
         TabIndex        =   63
         Top             =   5160
         Width           =   3600
         Begin VB.VScrollBar VScrollInformasi 
            Height          =   1575
            LargeChange     =   100
            Left            =   3240
            Max             =   8000
            TabIndex        =   165
            Top             =   240
            Width           =   300
         End
         Begin VB.PictureBox Picture3 
            Height          =   1575
            Left            =   0
            ScaleHeight     =   1515
            ScaleWidth      =   3195
            TabIndex        =   164
            Top             =   240
            Width           =   3255
            Begin VB.Frame frameInformasi 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   5055
               Left            =   0
               TabIndex        =   166
               Top             =   0
               Width           =   3255
               Begin VB.CheckBox mnu1_9_19 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Mapping Customer"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   184
                  Top             =   2160
                  Width           =   1620
               End
               Begin VB.CheckBox mnu1_9_15 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Konfirmasi Sublet"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   183
                  Top             =   1425
                  Width           =   1605
               End
               Begin VB.CheckBox mnu1_9_14 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "TPSS Customer"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   182
                  Top             =   1185
                  Width           =   2205
               End
               Begin VB.CheckBox mnu1_9_7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Booking"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   181
                  Top             =   945
                  Width           =   1245
               End
               Begin VB.CheckBox mnu1_9_6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Retur Invoice"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   180
                  Top             =   705
                  Width           =   1365
               End
               Begin VB.CheckBox mnu1_9_5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Batal Order"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   179
                  Top             =   480
                  Width           =   1365
               End
               Begin VB.CheckBox mnu1_9_1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Repair History"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   178
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.CheckBox mnu1_9_2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Service Per Tanggal"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   177
                  Top             =   240
                  Width           =   1845
               End
               Begin VB.CheckBox mnu1_9_18 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Tagih Cuci"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   176
                  Top             =   2445
                  Width           =   2085
               End
               Begin VB.CheckBox mnu1_9_17 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Audit Bengkel"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   175
                  Top             =   1920
                  Width           =   1485
               End
               Begin VB.CheckBox mnu1_9_16 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Persetujuan Diskon"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   174
                  Top             =   1680
                  Width           =   2205
               End
               Begin VB.CheckBox mnu1_9_20 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Validasi Request Disc"
                  Height          =   330
                  Left            =   120
                  TabIndex        =   173
                  Top             =   2640
                  Width           =   2220
               End
               Begin VB.CheckBox mnu1_9_4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Display All Repair"
                  Height          =   285
                  Left            =   720
                  TabIndex        =   172
                  Top             =   4755
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.CheckBox mnu1_9_3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "To Day on Job Progress"
                  Height          =   285
                  Left            =   720
                  TabIndex        =   171
                  Top             =   4470
                  Visible         =   0   'False
                  Width           =   2085
               End
               Begin VB.CheckBox mnu1_9_21 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Mapping Retention"
                  Height          =   330
                  Left            =   120
                  TabIndex        =   170
                  Top             =   2925
                  Width           =   2220
               End
               Begin VB.CheckBox mnu1_9_22 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Recovery Reminder Mapping"
                  Height          =   330
                  Left            =   120
                  TabIndex        =   169
                  Top             =   3195
                  Width           =   2940
               End
               Begin VB.CheckBox mnu1_9_23 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Laporan LS"
                  Height          =   330
                  Left            =   120
                  TabIndex        =   168
                  Top             =   3480
                  Width           =   2220
               End
               Begin VB.CheckBox mnu1_9_24 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Inject SA"
                  Height          =   330
                  Left            =   120
                  TabIndex        =   167
                  Top             =   3720
                  Width           =   2220
               End
            End
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tabel Master"
         Height          =   1500
         Index           =   7
         Left            =   3720
         TabIndex        =   67
         Top             =   1965
         Width           =   3660
         Begin VB.VScrollBar VScrollMaster 
            Height          =   1215
            LargeChange     =   100
            Left            =   3360
            Max             =   3500
            TabIndex        =   186
            Top             =   240
            Width           =   300
         End
         Begin VB.PictureBox Picture4 
            Height          =   1215
            Left            =   0
            ScaleHeight     =   1155
            ScaleWidth      =   3315
            TabIndex        =   185
            Top             =   240
            Width           =   3375
            Begin VB.Frame frameMaster 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   5175
               Left            =   0
               TabIndex        =   187
               Top             =   0
               Width           =   3375
               Begin VB.CheckBox mnu1_5_9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pelanggan"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   201
                  Top             =   990
                  Width           =   1155
               End
               Begin VB.CheckBox mnu1_5_8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Hobi"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   200
                  Top             =   750
                  Width           =   2835
               End
               Begin VB.CheckBox mnu1_5_3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Mekanik"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   199
                  Top             =   0
                  Width           =   2805
               End
               Begin VB.CheckBox mnu1_5_1 
                  BackColor       =   &H000000FF&
                  Caption         =   "Jenis Pekerjaan"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   198
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.CheckBox mnu1_5_5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Meterai"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   197
                  Top             =   270
                  Width           =   2235
               End
               Begin VB.CheckBox mnu1_5_7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Customer"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   196
                  Top             =   510
                  Width           =   1035
               End
               Begin VB.CheckBox mnu1_5_14 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Kode Dealer"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   195
                  Top             =   2280
                  Width           =   1635
               End
               Begin VB.CheckBox mnu1_5_13 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Actual Rate"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   194
                  Top             =   2040
                  Width           =   1635
               End
               Begin VB.CheckBox mnu1_5_12 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Asuransi"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   193
                  Top             =   1800
                  Width           =   1875
               End
               Begin VB.CheckBox mnu1_5_11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Agama"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   192
                  Top             =   1560
                  Width           =   1755
               End
               Begin VB.CheckBox mnu1_5_10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Wilayah"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   191
                  Top             =   1320
                  Width           =   1155
               End
               Begin VB.CheckBox mnu1_5_2 
                  BackColor       =   &H000000FF&
                  Caption         =   "Model Kendaraan"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   190
                  Top             =   3840
                  Visible         =   0   'False
                  Width           =   1605
               End
               Begin VB.CheckBox mnu1_5_6 
                  BackColor       =   &H000000FF&
                  Caption         =   "Dealer"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   189
                  Top             =   4080
                  Visible         =   0   'False
                  Width           =   2115
               End
               Begin VB.CheckBox mnu1_5_4 
                  BackColor       =   &H000000FF&
                  Caption         =   "Flate Rate"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   188
                  Top             =   4320
                  Visible         =   0   'False
                  Width           =   1125
               End
            End
         End
      End
      Begin VB.Frame fCabang 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Registrasi"
         Height          =   765
         Index           =   4
         Left            =   3720
         TabIndex        =   65
         Top             =   120
         Width           =   3660
         Begin VB.CheckBox mnu1_8_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pembatalan Registrasi"
            Height          =   285
            Left            =   150
            TabIndex        =   12
            Top             =   435
            Width           =   2685
         End
         Begin VB.CheckBox mnu1_8_1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pengisian Registrasi"
            Height          =   285
            Left            =   150
            TabIndex        =   11
            Top             =   195
            Width           =   2685
         End
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   10080
      TabIndex        =   50
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   11400
      TabIndex        =   51
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   12660
      TabIndex        =   52
      Top             =   8640
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   30
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9255
      Left            =   0
      Picture         =   "frmSetupPetugas.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmSetupPetugas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSQL As String
Dim rsTemp As ADODB.Recordset

Public Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(Left(gsSql1(i), 2))
            Select Case aa
            Case "FN", "FB", "FM"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(Left(gsSql1(i), 2))
            Select Case aa
            Case "FN", "FB", "FM"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkHO_Click()
    If chkHO = Checked Then
        FrameHO(0).Visible = True
        FrameHO(1).Visible = True
        FrameHO(2).Visible = True
        FrameHO(3).Visible = True
    Else
        FrameHO(0).Visible = False
        FrameHO(1).Visible = False
        FrameHO(2).Visible = False
        FrameHO(3).Visible = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub ClearCek()
    For i = 1 To Me.Controls.Count - 1
        If Left(Me.Controls(i).Name, 3) = "mnu" Then
           Me.Controls(i).value = 0
        End If
    Next
End Sub

Private Sub ClearVariabel()
    txtEntry(1) = ""
    txtEntry(2) = ""
    txtPass = ""
    txtEntry(6) = ""
End Sub

Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdHapus_Click()
Dim rsSearch As ADODB.Recordset

    On Error GoTo errornih
    If Trim(txtEntry(0)) <> "NON" Then
        Set rsSearch = objO.FetchData("Select * from t_petugas WITH(NOLOCK) where fs_kd_petugas='" + txtEntry(0) + "'")
        If Not rsSearch.BOF Then
            yn = MsgBox("Yakin untuk menghapus record ini ?", vbYesNo, "Konfirmasi ?")
            If yn = vbYes Then
                Call objO.DeleteData("DELETE t_petugas WHERE fs_kd_petugas='" & txtEntry(0).Text & "'")
                Call objO.DeleteData("Delete t_akses where left(akses,5)='" & (AddSpace(txtEntry(0), 5, 1, " ")) & "'")
                cmdHapus.Enabled = False
                cmdSimpan.Enabled = False
                Call ClearVariabel
                Call ClearCek
            End If
        Else
            MsgBox "Data Belum Ada"
        End If
         
    End If
    txtEntry(0) = ""
    txtEntry(0).SetFocus
    On Error GoTo 0
    Exit Sub
    
errornih:
    oRs.Close
    MsgBox "Data Masih Terpakai Pada Transaksi, Proses Dibatalkan !!!"
    On Error GoTo 0
End Sub

Private Sub cmdHapus_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdSimpan_Click()
      Dim rsSearch As ADODB.Recordset
          
1         If Trim(txtEntry(0)) <> "" And Trim(txtEntry(1)) <> "" Then
2             Set rsSearch = objO.FetchData("Select * from t_petugas WITH(NOLOCK) where fs_kd_petugas='" + txtEntry(0) + "'")
3             Call ClearsSql
4             gsSql1(0) = "FS_KD_PETUGAS"
5             gsSql2(0) = txtEntry(0)
6             gsSql1(1) = "FS_nm_PETUGAS"
7             gsSql2(1) = txtEntry(1)
8             gsSql1(2) = "FS_KD_group"
9             gsSql2(2) = txtEntry(6)
10            gsSql1(3) = "FS_KD_pass"
11            gsSql2(3) = AddSpace(txtEntry(2), 2, 3, "0") + UCase(txtPass)        'txtEntry(3)
12            gsSql1(4) = "TIPE"
13            gsSql2(4) = Trim(txtEntry(4))
14            gsSql1(5) = "DIVISI"
15            gsSql2(5) = IIf(opDiv(0), "SVC", "BP")
16            gsSql1(6) = "nik"
17            gsSql2(6) = NulSID(txtEntry(5))
18            If Log.NHS = 1 Then
                  
19                gsSql1(7) = "dealer"
20                gsSql2(7) = NulSID(gKdCbg)
21            End If
              
22            If Not rsSearch.BOF Then
23                sSQL = GensSql("t_petugas", True) + " where fs_kd_petugas='" & txtEntry(0) & "'"
24            Else
25                sSQL = GensSql("t_petugas", False)
26            End If
27            Call objO.UpdateData(sSQL)
              
28            Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..t_petugas WITH(NOLOCK) where fs_kd_petugas='" & txtEntry(0) & "' and dealer = '" & gDb & "'")
              
29            If Not rsSearch.BOF Then
30                gsSql1(7) = "dealer"
31                gsSql2(7) = NulSID(gDb)
32                sSQL = GensSql("[" & dbReplikasiHO & "]..t_petugas", True) + " where fs_kd_petugas='" & txtEntry(0) & "' and dealer = '" & gDb & "'"
33            Else
                  
34                gsSql1(7) = "dealer"
35                gsSql2(7) = NulSID(gDb)
36                sSQL = GensSql("[" & dbReplikasiHO & "]..t_petugas", False)
37            End If
38            Call objO.UpdateData(sSQL)

              'Simpan data ke t_akses--------
39            Call objO.DeleteData("delete t_akses where left(akses,5)='" & AddSpace(txtEntry(0), 5, 1, " ") & "'")

40            Call ClearsSql
41            gsSql1(0) = "akses"
42            gsSql1(1) = "fn_induk"
43            For i = 1 To Me.Controls.Count - 1
44                If Left(Me.Controls(i).Name, 3) = "mnu" And Right(Me.Controls(i).Name, 1) <> "s" Then
45                    If Me.Controls(i).value = 1 Then
46                        gsSql2(0) = AddSpace(txtEntry(0), 5, 1, " ") + Me.Controls(i).Name
47                        Set rsSearch = objO.FetchData("SELECT * FROM t_menu WITH(NOLOCK) where fs_nama ='" & Me.Controls(i).Name & "'")
48                        If Not rsSearch.BOF Then
49                            gsSql2(1) = rsSearch("fn_urut")
50                        End If
                          
51                        sSQL = GensSql("t_akses", False)
52                        Call objO.UpdateData(sSQL)
53                    End If
54                End If
55            Next

56            Call ClearCek
57            txtEntry(0).SetFocus
58            Sendkeys ("{home}+{end}")
              
59            cmdHapus.Enabled = False
60            cmdSimpan.Enabled = False
61            Call ClearVariabel
62            txtEntry(0) = ""
63            txtEntry(0).SetFocus
64        Else
65            MsgBox "Data Belum Lengkap, Ulangi Pengisian Data !!", vbInformation + vbOKOnly, "Perhatian"
66        End If
End Sub

Private Sub cmdSimpan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Form_Load()
    Call CenterForm(frmSetupPetugas, True)
    
    If gUstat = 1 Then      ' USER HO
        For i = 0 To 3
            FrameHO(i).Visible = True
        Next
        FrameHO(0).Left = 105
        FrameHO(0).Top = 120
        FrameHO(1).Left = FrameHO(0).Left
        FrameHO(1).Top = FrameHO(0).Top + FrameHO(0).Height + 50
        FrameHO(2).Left = FrameHO(1).Left
        FrameHO(2).Top = FrameHO(1).Top + FrameHO(1).Height + 50
        FrameHO(3).Left = FrameHO(2).Left
        FrameHO(3).Top = FrameHO(2).Top + FrameHO(2).Height + 50
        
        For i = 0 To 10
            fCabang(i).Visible = False
        Next
    Else                    ' USER CABANG
        For i = 0 To 3
            FrameHO(i).Visible = False
        Next
        For i = 0 To 10
            fCabang(i).Visible = True
        Next
    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub mnu1_10_1_Click()
    Call TekanLapOrder(mnu1_10_1.value)
End Sub

Private Sub TekanLapOrder(nilai As Byte)
    mnu1_10_1_1.value = nilai
    mnu1_10_1_2.value = nilai
    mnu1_10_1_3.value = nilai
    mnu1_10_1_4.value = nilai
    mnu1_10_1_5.value = nilai
End Sub

Private Sub TekanLapBilling(nilai As Byte)
    mnu1_10_2_1.value = nilai
    mnu1_10_2_2.value = nilai
    mnu1_10_2_3.value = nilai
    mnu1_10_2_4.value = nilai
    mnu1_10_2_5.value = nilai
    mnu1_10_2_6.value = nilai
    mnu1_10_2_7.value = nilai
    mnu1_10_2_8.value = nilai
    mnu1_10_2_9.value = nilai
    mnu1_10_2_10.value = nilai
    mnu1_10_2_11.value = nilai
    mnu1_10_2_12.value = nilai
    mnu1_10_2_13.value = nilai
    mnu1_10_2_14.value = nilai
    mnu1_10_2_15.value = nilai
End Sub

Private Sub TekanLapQty(nilai As Byte)
    mnu1_10_3_1.value = nilai
    mnu1_10_3_2.value = nilai
    mnu1_10_3_3.value = nilai
End Sub

Private Sub TekanLapHO(nilai As Byte)
    mnu1_22_1.value = nilai
    mnu1_22_2.value = nilai
    mnu1_22_3.value = nilai
    mnu1_22_4.value = nilai
    mnu1_22_5.value = nilai
End Sub

Private Sub TekanMasterLapHO(nilai As Byte)
    mnu1_23_1.value = nilai
    mnu1_23_2.value = nilai
End Sub

Private Sub TekanPerbaikanDataHO(nilai As Byte)
    mnu1_25_1.value = nilai
    mnu1_25_2.value = nilai
    mnu1_25_3.value = nilai
    mnu1_25_4.value = nilai
    mnu1_25_5.value = nilai
End Sub

Private Sub TekanMasterSettingHO(nilai As Byte)
    mnu1_24_1.value = nilai
    mnu1_24_2.value = nilai
    mnu1_24_3.value = nilai
End Sub

Private Sub TekanLapMRS(nilai As Byte)
    mnu1_10_7_1.value = nilai
    mnu1_10_7_2.value = nilai
    mnu1_10_7_3.value = nilai
    mnu1_10_7_4.value = nilai
End Sub

Private Sub TekanLapRekap(nilai As Byte)
    mnu1_10_6_1.value = nilai
    mnu1_10_6_2.value = nilai
    mnu1_10_6_3.value = nilai
    mnu1_10_6_4.value = nilai
    mnu1_10_6_5.value = nilai
    mnu1_10_6_7.value = nilai
    mnu1_10_6_8.value = nilai
End Sub

Private Sub tekanLapMonitoring(nilai As Byte)
    mnu1_10_8_1.value = nilai
    mnu1_10_8_2.value = nilai
    mnu1_10_8_3.value = nilai
    mnu1_10_8_5.value = nilai
End Sub

Private Sub TekanMenuHO(nilai As Byte)
    mnu1_21_1.value = nilai
    mnu1_21_2.value = nilai
    mnu1_21_3.value = nilai
    mnu1_21_4.value = nilai
    mnu1_21_5.value = nilai
    mnu1_21_6.value = nilai
    mnu1_21_7.value = nilai
    mnu1_21_8.value = nilai
    mnu1_21_9.value = nilai
    mnu1_21_10.value = nilai
End Sub

Private Sub TekanLapRev(nilai As Byte)
    mnu1_10_4_1.value = nilai
    mnu1_10_4_2.value = nilai
End Sub

Private Sub TekanLapMech(nilai As Byte)
    'mnu1_10_5_1.value = Nilai
    mnu1_10_5_2.value = nilai
End Sub

Private Sub mnu1_10_1_KeyPress(KeyAscii As Integer)
    Call TekanLapOrder(mnu1_10_1.value)
End Sub

Private Sub mnu1_10_2_Click()
    Call TekanLapBilling(mnu1_10_2.value)
End Sub

Private Sub mnu1_10_2_KeyPress(KeyAscii As Integer)
    Call TekanLapBilling(mnu1_10_2.value)
End Sub

Private Sub mnu1_10_3_1_Click()
    mnu1_10_3.value = mnu1_10_3_1.value
End Sub

Private Sub mnu1_10_3_Click()
    Call TekanLapQty(mnu1_10_3.value)
End Sub

Private Sub mnu1_10_4_2_Click()
    mnu1_10_4.value = mnu1_10_4_2.value
End Sub

Private Sub mnu1_10_4_Click()
    Call TekanLapRev(mnu1_10_4.value)
End Sub

Private Sub mnu1_10_5_2_Click()
    mnu1_10_5.value = mnu1_10_5_2.value
End Sub

Private Sub mnu1_10_5_Click()
    Call TekanLapMech(mnu1_10_5.value)
End Sub

Private Sub mnu1_10_6_Click()
    Call TekanLapRekap(mnu1_10_6.value)
End Sub

Private Sub mnu1_10_7_Click()
    Call TekanLapMRS(mnu1_10_7.value)
End Sub

Private Sub mnu1_10_8_Click()
    Call tekanLapMonitoring(mnu1_10_8.value)
End Sub

Private Sub mnu1_221_21_Click()
    Call TekanMenuHO(mnu1_221_21.value)
End Sub

Private Sub mnu1_221_22_Click()
    Call TekanLapHO(mnu1_221_22.value)
End Sub

Private Sub mnu1_221_23_Click()
    Call TekanMasterLapHO(mnu1_221_23.value)
End Sub

Private Sub mnu1_221_24_Click()
    Call TekanMasterSettingHO(mnu1_221_24.value)
End Sub

Private Sub mnu1_221_25_Click()
    Call TekanPerbaikanDataHO(mnu1_221_25.value)
End Sub

Private Sub opDiv_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If txtEntry(Index).BackColor = &H808000 Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            txtEntry(Index) = gMainKeyCari
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

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
        
    If KeyCode = vbKeyF12 Then
        Select Case Index
        Case 0, 6
            gMainKeyCari = txtEntry(Index)
            gSearchOption = "PETUGAS"
            frmCari.show
        Case 4
            gMainKeyCari = txtEntry(Index)
            gSearchOption = "TIPEPETUGAS"
            frmCari.show
        End Select
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset
Dim rsHelp As ADODB.Recordset
Dim vNama As String
    
    txtEntry(Index) = LTrim(txtEntry(Index))
    Select Case Index
        Case 0
            If Trim(txtEntry(0)) <> "" Then
                Set rsSearch = objO.FetchData("Select * from t_petugas WITH(NOLOCK) where fs_kd_petugas='" + txtEntry(0) + "'")
                If Not rsSearch.BOF Then
                    txtEntry(1) = rsSearch("fs_nm_petugas")
                    txtEntry(6) = rsSearch("fs_kd_group")
                    txtEntry(4) = rsSearch("TIPE")
                    txtEntry(5) = NulSID(rsSearch("nik"))
                    txtPass = Mid(rsSearch("fs_kd_pass"), 3, 10)
                    If rsSearch("TIPE") = "ADMIN" Then
                        Set rsTemp = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE FS_KD_PETUGAS = '" & gSA & "'")
                        If Not rsTemp.EOF Then
                            If rsTemp("TIPE") = "ADMIN" Then
                                txtEntry(4) = rsSearch("TIPE")
                            End If
                        Else
                            MsgBox "Tidak bisa Mengakses user tersebut !!", vbCritical, "LIMIT AKSES"
                            For i = 0 To 3
                                txtEntry(i) = ""
                            Next
                            txtEntry(6) = ""
                            Exit Sub
                        End If
                        Call setClose(rsTemp)
                    End If
                    txtEntry(2) = ""
                    If Trim(txtEntry(6)) <> "" Then
                        Set rsHelp = objO.FetchData("Select fs_nm_petugas from t_petugas WITH(NOLOCK) where fs_kd_petugas = '" & txtEntry(6) & "'")
                        If Not rsHelp.BOF Then
                            txtEntry(2) = rsHelp("fs_nm_petugas")
                        End If
                        If rsHelp.State <> adStateClosed Then rsHelp.Close
                    End If
                    cmdHapus.Enabled = True
                Else
                    Call ClearVariabel
                    cmdHapus.Enabled = False
                End If
                cmdSimpan.Enabled = True
                 
                
                Call ClearCek
                Set rsSearch = objO.FetchData("select * from t_akses WITH(NOLOCK) where Left(akses,5)='" & AddSpace(txtEntry(0), 5, 1, " ") & "'")
                If Not rsSearch.BOF Then
                    rsSearch.MoveFirst
                    Do While Not rsSearch.EOF
                        vNama = Mid(rsSearch("akses"), 6, Len(rsSearch("akses")) - 5)
                        If Right(vNama, 1) = "s" Then
                            Me.Controls(vNama).value = 1
                            vNama = Left(vNama, Len(vNama) - 1)
                        End If
                        Me.Controls(vNama).value = 1
                        rsSearch.MoveNext
                    Loop
                End If
                 
            Else
                txtEntry(1) = ""
                cmdSimpan.Enabled = False
                cmdHapus.Enabled = False
            End If
        Case 4
            If Trim(txtEntry(Index)) <> "" Then
                Set rsSearch = objO.FetchData("SELECT TIPE FROM TIPEPETUGAS WITH(NOLOCK) WHERE TIPE = '" & Trim(txtEntry(Index)) & "'")
                If Not rsSearch.BOF Then
                    If rsSearch("TIPE") = "HO" Or rsSearch("TIPE") = "ADMIN" Then
                        FrameHO(0).Visible = True
                        FrameHO(1).Visible = True
                        FrameHO(2).Visible = True
                        FrameHO(3).Visible = True
                        FrameHO(4).Visible = True
                    Else
                        FrameHO(0).Visible = False
                        FrameHO(1).Visible = False
                        FrameHO(2).Visible = False
                        FrameHO(3).Visible = False
                        FrameHO(4).Visible = False
                    End If
                Else
                    Cancel = True
                    gMainKeyCari = txtEntry(Index)
                    gSearchOption = "TIPEPETUGAS"
                    frmCari.show
                End If
            End If
        Case 6
            If Trim(txtEntry(Index)) <> "" Then
                Set rsSearch = objO.FetchData("Select * from t_petugas WITH(NOLOCK) where fs_kd_petugas = '" + txtEntry(6) + "'")
                If Not rsSearch.BOF Then
                    txtEntry(2) = rsSearch("fs_nm_petugas")
                Else
                    Cancel = True
                    gMainKeyCari = txtEntry(Index)
                    gSearchOption = "PETUGAS"
                    frmCari.show
                End If
                cmdSimpan.Enabled = True
            Else
                txtEntry(2) = ""
            End If
    End Select
     
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        FrameReminder.Top = (-VScroll1.value * 10) + 20
    Else
        FrameReminder.Top = 0
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        FrameReminder.Top = -VScroll1.value + 20
    Else
        FrameReminder.Top = 0
    End If
End Sub

Private Sub VScrollBilling_Change()
    If VScrollBilling.value > 0 Then
        FrameBilling.Top = -VScrollBilling.value + 20
    Else
        FrameBilling.Top = 0
    End If
End Sub

Private Sub VScrollBilling_Scroll()
    If VScrollBilling.value > 0 Then
        FrameBilling.Top = -VScrollBilling.value + 20
    Else
        FrameBilling.Top = 0
    End If
End Sub

Private Sub vscrollBooking_Change()
    If vscrollBooking.value > 0 Then
        frameBooking.Top = -vscrollBooking.value + 20
    Else
        frameBooking.Top = 0
    End If
End Sub

Private Sub vscrollBooking_Scroll()
    If vscrollBooking.value > 0 Then
        frameBooking.Top = -vscrollBooking.value + 20
    Else
        frameBooking.Top = 0
    End If
End Sub

Private Sub VScrollInformasi_Change()
    If VScrollInformasi.value > 0 Then
        frameInformasi.Top = -VScrollInformasi.value + 20
    Else
        frameInformasi.Top = 0
    End If
End Sub

Private Sub VScrollInformasi_Scroll()
    If VScrollInformasi.value > 0 Then
        frameInformasi.Top = -VScrollInformasi.value + 20
    Else
        frameInformasi.Top = 0
    End If
End Sub

Private Sub VScrollMaster_Change()
    If VScrollMaster.value > 0 Then
        frameMaster.Top = -VScrollMaster.value + 20
    Else
        frameMaster.Top = 0
    End If
End Sub

Private Sub VScrollMaster_Scroll()
    If VScrollMaster.value > 0 Then
        frameMaster.Top = -VScrollMaster.value + 20
    Else
        frameMaster.Top = 0
    End If
End Sub

Private Sub VScrollOrder_Change()
    If VScrollOrder.value > 0 Then
        frameOrder.Top = -VScrollOrder.value + 20
    Else
        frameOrder.Top = 0
    End If
End Sub

Private Sub VScrollOrder_Scroll()
    If VScrollOrder.value > 0 Then
        frameOrder.Top = -VScrollOrder.value + 20
    Else
        frameOrder.Top = 0
    End If
End Sub

Private Sub VScrollSetting_Change()
    If VScrollSetting.value > 0 Then
        frameSetting.Top = -VScrollSetting.value + 20
    Else
        frameSetting.Top = 0
    End If
End Sub

Private Sub VScrollSetting_Scroll()
    If VScrollSetting.value > 0 Then
        frameSetting.Top = -VScrollSetting.value + 20
    Else
        frameSetting.Top = 0
    End If
End Sub
