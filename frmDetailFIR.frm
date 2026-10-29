VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmDetailFIR 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detail Informasi Gejala & Diagnosa"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detail Informasi Gejala & Diagnosa"
      Height          =   4575
      Left            =   120
      TabIndex        =   53
      Top             =   5280
      Width           =   12495
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Skip"
         Height          =   255
         Left            =   9720
         TabIndex        =   115
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Skip"
         Height          =   255
         Left            =   9720
         TabIndex        =   114
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Skip"
         Height          =   255
         Left            =   4800
         TabIndex        =   113
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Skip"
         Height          =   255
         Left            =   4800
         TabIndex        =   112
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   3120
         TabIndex        =   70
         Top             =   2880
         Width           =   9255
         Begin VB.OptionButton optCuaca 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
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
            Index           =   5
            Left            =   3840
            TabIndex        =   111
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optCuaca 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lembab"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   75
            Top             =   260
            Width           =   975
         End
         Begin VB.OptionButton optCuaca 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hujan"
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
            Index           =   2
            Left            =   1800
            TabIndex        =   74
            Top             =   320
            Width           =   1215
         End
         Begin VB.OptionButton optCuaca 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berawan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   73
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optCuaca 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cerah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton optCuaca 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Panas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   71
            Top             =   120
            Width           =   855
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   7320
            TabIndex        =   76
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16761024
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   78
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Temperatur Udara Luar :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   5040
            TabIndex        =   77
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3000
         TabIndex        =   65
         Top             =   2520
         Width           =   9255
         Begin VB.OptionButton optLaluLintas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
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
            Index           =   3
            Left            =   7680
            TabIndex        =   110
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optLaluLintas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lainnya"
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
            Left            =   3960
            TabIndex        =   68
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optLaluLintas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lancar"
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
            Left            =   1920
            TabIndex        =   67
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optLaluLintas 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Macet Stop & GO"
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
            Left            =   120
            TabIndex        =   66
            Top             =   0
            Width           =   1575
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   4920
            TabIndex        =   69
            Top             =   0
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            PromptChar      =   ">"
         End
      End
      Begin VB.CheckBox Check46 
         BackColor       =   &H00FFFFFF&
         Caption         =   "General Road"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   64
         Top             =   1000
         Width           =   1575
      End
      Begin VB.CheckBox Check45 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bergelombang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   63
         Top             =   1000
         Width           =   1575
      End
      Begin VB.CheckBox Check44 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tol Belok Kiri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   62
         Top             =   1000
         Width           =   1695
      End
      Begin VB.CheckBox Check43 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jalan Lurus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   61
         Top             =   1000
         Width           =   1455
      End
      Begin VB.CheckBox Check42 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   60
         Top             =   1410
         Width           =   1335
      End
      Begin VB.CheckBox Check41 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tanjakan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10560
         TabIndex        =   59
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox Check40 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Akselerasi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   58
         Top             =   1410
         Width           =   1455
      End
      Begin VB.CheckBox Check39 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pengereman"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   57
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox Check38 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lainnya"
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
         Left            =   3120
         TabIndex        =   56
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check47 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Belok Kanan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   55
         Top             =   1000
         Width           =   1485
      End
      Begin VB.CheckBox Check48 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turunan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         TabIndex        =   54
         Top             =   1410
         Width           =   1485
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   79
         Top             =   240
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   8040
         TabIndex        =   80
         Top             =   240
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
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
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   3120
         TabIndex        =   81
         Top             =   600
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
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
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   82
         Top             =   600
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   525
         Index           =   7
         Left            =   4320
         TabIndex        =   83
         Top             =   1920
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   926
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   4560
         TabIndex        =   84
         Top             =   3480
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   8160
         TabIndex        =   85
         Top             =   3480
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   645
         Index           =   13
         Left            =   3000
         TabIndex        =   103
         Top             =   3840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1138
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Catatan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   330
         TabIndex        =   104
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kecepatan Kendaraan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "KM / J"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   99
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "RPM / Putaran Mesin :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   98
         Top             =   645
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kondisi Jalan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   97
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Beban Kendaraan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   5040
         TabIndex        =   96
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   95
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RPM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   94
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jumlah Penumpang :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5040
         TabIndex        =   93
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Orang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   92
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lalu Lintas :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   91
         Top             =   2540
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kondisi Cuaca :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   90
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Blower Speed :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3120
         TabIndex        =   89
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Untuk Gejala AC :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   88
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Temperatur Setting :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   6120
         TabIndex        =   87
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   86
         Top             =   3480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   12495
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3120
         TabIndex        =   43
         Top             =   1800
         Width           =   8775
         Begin VB.OptionButton optKondisiMesin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6000
            TabIndex        =   109
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton optKondisiMesin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lainnya : Semua"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   46
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton optKondisiMesin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dingin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   45
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optKondisiMesin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Panas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3120
         TabIndex        =   39
         Top             =   1440
         Width           =   8655
         Begin VB.OptionButton optKondisi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6000
            TabIndex        =   108
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton optKondisi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Kecepatan Konstan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   42
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton optKondisi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Starting"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   41
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optKondisi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Idling"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3000
         TabIndex        =   35
         Top             =   1050
         Width           =   7575
         Begin VB.OptionButton optIndikator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6120
            TabIndex        =   107
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optIndikator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4080
            TabIndex        =   38
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optIndikator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Selalu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2040
            TabIndex        =   37
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optIndikator 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ON"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3120
         TabIndex        =   31
         Top             =   600
         Width           =   9135
         Begin VB.OptionButton optFrekuensi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6000
            TabIndex        =   106
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optFrekuensi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Selalu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   34
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optFrekuensi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Kadang"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1920
            TabIndex        =   33
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optFrekuensi 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sekali"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   3120
         TabIndex        =   25
         Top             =   120
         Width           =   9255
         Begin VB.OptionButton optGejala 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   7320
            TabIndex        =   105
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optGejala 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lainnya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   28
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optGejala 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Minggu Lalu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   27
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton optGejala 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hari Ini"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   2
            Left            =   5040
            TabIndex        =   29
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   ">"
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hari Yang Lalu"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   30
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   24
         Top             =   2220
         Width           =   615
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   23
         Top             =   2220
         Width           =   495
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   22
         Top             =   2220
         Width           =   495
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   21
         Top             =   2220
         Width           =   615
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   20
         Top             =   2220
         Width           =   615
      End
      Begin VB.CheckBox Check21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox Check22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   18
         Top             =   2640
         Width           =   495
      End
      Begin VB.CheckBox Check23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox Check24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox Check25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox Check26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lainnya : Semua Kondisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   14
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gejala Terjadi Pertama Kali :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frekuensi Gejala Terjadi :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   645
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Indikator (MIL) saat gejala terjadi :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kondisi Kendaraan saat MIL terjadi :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1455
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kondisi Mesin :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   48
         Top             =   1845
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Posisi Shift Level / Gigi Kecepatan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   47
         Top             =   2250
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Kendaraan"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   116
         Top             =   890
         Width           =   11175
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   1200
         TabIndex        =   3
         Top             =   550
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   5400
         TabIndex        =   7
         Top             =   550
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12000
         TabIndex        =   12
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keluhan : "
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
         Left            =   450
         TabIndex        =   9
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model :"
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
         Left            =   4845
         TabIndex        =   8
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Model : "
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
         Left            =   4560
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi : "
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
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal : "
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
         Left            =   90
         TabIndex        =   2
         Top             =   285
         Width           =   1080
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "KONDISI GEJALA SAAT TERJADI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   12255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   12495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "DIAGNOSA AWAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1500
      Width           =   12255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   12495
   End
End
Attribute VB_Name = "frmDetailFIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isBooking As Boolean
Public noOrder As String
Public NoBooking As String

Public noChassis As String
Public sSQL As String
Public rsTemp As ADODB.Recordset
Dim gejala As String
Dim frekuensi As String
Dim indikator As String
Dim kondisiMIL As String
Dim kondisiMesin As String
Dim gigiKec As String
Dim kondisiJalan As String
Dim laluLintas As String
Dim cuaca As String
Dim isBoleh As Boolean
Dim pesanError As String


Private Sub Check38_Click()
   If Check38.value = vbChecked Then
      Text(7).BackColor = &HFFFFFF
      
   Else
      Text(7).BackColor = &HC0E0FF
      Text(7).Enabled = True
   End If
End Sub

Private Sub Command1_Click()
1        isBoleh = True
2        pesanError = ""
         
3        If Check1.value = vbUnchecked Then
4           If Trim(Text(3)) = "" Then
5              isBoleh = False
6              pesanError = pesanError & "Kecepatan Kendaraan, "
7           End If
8        End If
         
9        If Check2.value = vbUnchecked Then
10          If Trim(Text(5)) = "" Then
11             isBoleh = False
12             pesanError = pesanError & "RPM, "
13          End If
14       End If
         
15       If Check3.value = vbUnchecked Then
16          If Trim(Text(4)) = "" Then
17             isBoleh = False
18             pesanError = pesanError & "Beban, "
19          End If
20       End If
         
21       If Check4.value = vbUnchecked Then
22          If Trim(Text(6)) = "" Then
23             isBoleh = False
24             pesanError = pesanError & "Jumlah Penumpang, "
25          End If
26       End If
         
         'If Check5.value = vbUnchecked Then
      '      If Trim(Text(11)) = "" Then
      '         isBoleh = False
      '         pesanError = pesanError & "Blower AC, "
      '      End If
         'End If
         
        ' If Check6.value = vbUnchecked Then
      '      If Trim(Text(12)) = "" Then
      '         isBoleh = False
      '         pesanError = pesanError & "Temperatur AC, "
      '      End If
        ' End If
         
27       If optGejala(0).value = True Then
28          gejala = "Hari Ini"
29       ElseIf optGejala(1).value = True Then
30          gejala = "Minggu Lalu"
31       ElseIf optGejala(2).value = True Then
32          gejala = Trim(Text(2)) & " Hari Yang Lalu"
33       ElseIf optGejala(3).value = True Then
         
34       Else
35          isBoleh = False
36          pesanError = pesanError & "Gejala, "
37       End If
         
38       If optFrekuensi(0).value = True Then
39          frekuensi = "Sekali"
40       ElseIf optFrekuensi(1).value = True Then
41          frekuensi = "Kadang"
42       ElseIf optFrekuensi(2).value = True Then
43          frekuensi = "Selalu"
44       ElseIf optFrekuensi(3).value = True Then
         
45       Else
46          isBoleh = False
47          pesanError = pesanError & "Frekuensi Gejala, "
48       End If
         
49       If optIndikator(0).value = True Then
50          indikator = "ON"
51       ElseIf optIndikator(1).value = True Then
52          indikator = "Selalu"
53       ElseIf optIndikator(2).value = True Then
54          indikator = "OFF"
55       ElseIf optIndikator(3).value = True Then
         
56       Else
57          isBoleh = False
58          pesanError = pesanError & "Indikator MIL, "
59       End If
         
60       If optKondisi(0).value = True Then
61          kondisiMIL = "Idling"
62       ElseIf optKondisi(1).value = True Then
63          kondisiMIL = "Starting"
64       ElseIf optKondisi(2).value = True Then
65          kondisiMIL = "Kecepatan Konstan"
66       ElseIf optKondisi(3).value = True Then
         
67       Else
68          isBoleh = False
69          pesanError = pesanError & "Kondisi Kendaraan Saat MIL, "
70       End If
         
71       If optKondisiMesin(0).value = True Then
72          kondisiMesin = "Panas"
73       ElseIf optKondisiMesin(1).value = True Then
74          kondisiMesin = "Dingin"
75       ElseIf optKondisiMesin(2).value = True Then
76          kondisiMesin = "Semua"
77       ElseIf optKondisiMesin(3).value = True Then
         
78       Else
79          isBoleh = False
80          pesanError = pesanError & "Kondisi Mesin, "
81       End If
         
82       gigiKec = ""
83       If Check16.value = vbChecked Then gigiKec = gigiKec & "1,"
84       If Check19.value = vbChecked Then gigiKec = gigiKec & "2,"
85       If Check17.value = vbChecked Then gigiKec = gigiKec & "3,"
86       If Check20.value = vbChecked Then gigiKec = gigiKec & "4,"
87       If Check18.value = vbChecked Then gigiKec = gigiKec & "5,"
88       If Check23.value = vbChecked Then gigiKec = gigiKec & "P,"
89       If Check24.value = vbChecked Then gigiKec = gigiKec & "R,"
90       If Check22.value = vbChecked Then gigiKec = gigiKec & "N,"
91       If Check25.value = vbChecked Then gigiKec = gigiKec & "D,"
92       If Check21.value = vbChecked Then gigiKec = gigiKec & "S,"
93       If Check26.value = vbChecked Then gigiKec = gigiKec & "Semua Kondisi,"
         
94       If gigiKec <> "" Then
95          gigiKec = Mid(gigiKec, 1, Len(gigiKec) - 1)
96       Else
97          isBoleh = False
98          pesanError = pesanError & "Posisi Shift, "
99       End If
         
100      kondisiJalan = ""
101      If Check46.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check46.Caption) & ","
102      If Check45.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check45.Caption) & ","
103      If Check44.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check44.Caption) & ","
104      If Check47.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check47.Caption) & ","
105      If Check43.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check43.Caption) & ","
106      If Check42.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check42.Caption) & ","
107      If Check41.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check41.Caption) & ","
108      If Check48.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check48.Caption) & ","
109      If Check40.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check40.Caption) & ","
110      If Check39.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Check39.Caption) & ","
111      If Check38.value = vbChecked Then kondisiJalan = kondisiJalan & Trim(Text(7)) & ","
         
112      If kondisiJalan <> "" Then kondisiJalan = Mid(kondisiJalan, 1, Len(kondisiJalan) - 1)
         
113      If optLaluLintas(0).value = True Then
114         laluLintas = "Macet Stop & Go"
115      ElseIf optLaluLintas(1).value = True Then
116         laluLintas = "Lancar"
117      ElseIf optLaluLintas(2).value = True Then
118         laluLintas = Trim(Text(8))
119      ElseIf optLaluLintas(3).value = True Then
         
120      Else
121         isBoleh = False
122         pesanError = pesanError & "Kondisi Lalu Lintas, "
123      End If
         
124      If optCuaca(0).value = True Then
125         cuaca = Trim(optCuaca(0).Caption)
126      ElseIf optCuaca(1).value = True Then
127         cuaca = Trim(optCuaca(1).Caption)
128      ElseIf optCuaca(2).value = True Then
129         cuaca = Trim(optCuaca(2).Caption)
130      ElseIf optCuaca(3).value = True Then
131         cuaca = Trim(optCuaca(3).Caption)
132      ElseIf optCuaca(4).value = True Then
133         cuaca = Trim(optCuaca(4).Caption)
134      ElseIf optCuaca(5).value = True Then
         
135      Else
136         isBoleh = False
137         pesanError = pesanError & "Cuaca, "
138      End If
         
139      If isBooking = True Then
            ''jika data dari booking
140         If NoBooking = "" Then
               ''jika booking tapi blm ada nobooking
141            sSQL = "SELECT * FROM DETAILFIR WITH (NOLOCK) WHERE NO_POLISI = '" & Trim(Text(14)) & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "' and TANGGAL = '" & Format(Trim(Text(16)), "mm/dd/yyyy") & "' "
142         Else
               ''jika no booking sudah ada dicari berdasarkan nobooking dan nopolisi
143            sSQL = "SELECT * FROM DETAILFIR WITH (NOLOCK) WHERE NOFAK = '" & Trim(NoBooking) & "' AND NO_POLISI = '" & Trim(Text(14)) & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "' "
144         End If
145      Else
146         sSQL = "SELECT * FROM DETAILFIR WITH (NOLOCK) WHERE nowo = '" & noOrder & "' and no_chassis = '" & noChassis & "' and nofak = '" & NoBooking & "' and NO_POLISI = '" & Trim(Text(14)) & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "' and TANGGAL = '" & Format(Trim(Text(16)), "mm/dd/yyyy") & "' "
147      End If
148      Set rsTemp = objO.FetchData(sSQL)
149      If Not rsTemp.EOF Then
            ''jika sudah ada data update data
150         Call ClearsSql
151         gsSql1(0) = "NOWO"
152         gsSql2(0) = Trim(noOrder)
153         gsSql1(1) = "NOFAK"
154         gsSql2(1) = Trim(NoBooking)
155         gsSql1(2) = "NO_POLISI"
156         gsSql2(2) = Trim(Text(14))
157         gsSql1(3) = "NO_CHASSIS"
158         gsSql2(3) = Trim(noChassis)
159         gsSql1(4) = "IDKELUHAN"
160         gsSql2(4) = Trim(Label15.Caption)
161         gsSql1(5) = "KELUHAN"
162         gsSql2(5) = Trim(Text1.Text)
163         gsSql1(6) = "GEJALAPERTAMA"
164         gsSql2(6) = gejala
165         gsSql1(7) = "FREKUENSI"
166         gsSql2(7) = frekuensi
167         gsSql1(8) = "INDIKATOR"
168         gsSql2(8) = indikator
169         gsSql1(9) = "SAATMIL"
170         gsSql2(9) = kondisiMIL
            
171         gsSql1(10) = "KONDISIMESIN"
172         gsSql2(10) = kondisiMesin
173         gsSql1(11) = "SHIFTLEVEL"
174         gsSql2(11) = gigiKec
175         gsSql1(12) = "SPEED"
176         gsSql2(12) = Trim(Text(3))
177         gsSql1(13) = "RPM"
178         gsSql2(13) = Trim(Text(5))
179         gsSql1(14) = "BEBAN"
180         gsSql2(14) = Trim(Trim(Text(4)))
181         gsSql1(15) = "PENUMPANG"
182         gsSql2(15) = Trim(Text(6))
183         gsSql1(16) = "KONDISIJALAN"
184         gsSql2(16) = kondisiJalan
185         gsSql1(17) = "LALULINTAS"
186         gsSql2(17) = laluLintas
187         gsSql1(18) = "CUACA"
188         gsSql2(18) = cuaca
189         gsSql1(19) = "BLOWERSPEEDAC"
190         gsSql2(19) = Trim(Text(11))
191         gsSql1(20) = "TEMPSETTING"
192         gsSql2(20) = Trim(Text(12))
193         gsSql1(21) = "TANGGAL"
194         gsSql2(21) = Format(Trim(Text(16)), "mm/dd/yyyy")
            
195         If isBooking = True Then
196            sSQL = GensSql2("DETAILFIR", True) & " WHERE NO_POLISI = '" & Trim(Text(14)) & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "' and TANGGAL = '" & Format(Trim(Text(16)), "mm/dd/yyyy") & "' "
197         Else
198            sSQL = GensSql2("DETAILFIR", True) & " WHERE nowo = '" & noOrder & "' and no_chassis = '" & noChassis & "' and nofak = '" & NoBooking & "' and NO_POLISI = '" & Trim(Text(14)) & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "' and TANGGAL = '" & Format(Trim(Text(16)), "mm/dd/yyyy") & "' "
199         End If
200         If isBoleh Then
201            Call objO.InsertData(sSQL)
202            MsgBox "Data berhasil disimpan", vbInformation
203         Else
204            MsgBox "Lengkapi data sebagai berikut" & vbCrLf & pesanError, vbInformation
205            Exit Sub
206         End If
207         Unload Me
            
208      Else
            '' data belum ada insert data
209         Call ClearsSql
210         gsSql1(0) = "NOWO"
211         gsSql2(0) = Trim(noOrder)
212         gsSql1(1) = "NOFAK"
213         gsSql2(1) = Trim(NoBooking)
214         gsSql1(2) = "NO_POLISI"
215         gsSql2(2) = Trim(Text(14))
216         gsSql1(3) = "NO_CHASSIS"
217         gsSql2(3) = Trim(noChassis)
218         gsSql1(4) = "IDKELUHAN"
219         gsSql2(4) = Trim(Label15.Caption)
220         gsSql1(5) = "KELUHAN"
221         gsSql2(5) = Trim(Text1.Text)
222         gsSql1(6) = "GEJALAPERTAMA"
223         gsSql2(6) = gejala
224         gsSql1(7) = "FREKUENSI"
225         gsSql2(7) = frekuensi
226         gsSql1(8) = "INDIKATOR"
227         gsSql2(8) = indikator
228         gsSql1(9) = "SAATMIL"
229         gsSql2(9) = kondisiMIL
            
230         gsSql1(10) = "KONDISIMESIN"
231         gsSql2(10) = kondisiMesin
232         gsSql1(11) = "SHIFTLEVEL"
233         gsSql2(11) = gigiKec
234         gsSql1(12) = "SPEED"
235         gsSql2(12) = Trim(Text(3))
236         gsSql1(13) = "RPM"
237         gsSql2(13) = Trim(Text(5))
238         gsSql1(14) = "BEBAN"
239         gsSql2(14) = IIf(Trim(Trim(Text(4))) = "", 0, Trim(Trim(Text(4))))
240         gsSql1(15) = "PENUMPANG"
241         gsSql2(15) = IIf(Trim(Text(6)) = "", 0, Trim(Text(6)))
242         gsSql1(16) = "KONDISIJALAN"
243         gsSql2(16) = kondisiJalan
244         gsSql1(17) = "LALULINTAS"
245         gsSql2(17) = laluLintas
246         gsSql1(18) = "CUACA"
247         gsSql2(18) = cuaca
248         gsSql1(19) = "BLOWERSPEEDAC"
249         gsSql2(19) = Trim(Text(11))
250         gsSql1(20) = "TEMPSETTING"
251         gsSql2(20) = IIf(Trim(Text(12)) = "", 0, Trim(Text(12)))
252         gsSql1(21) = "TANGGAL"
253         gsSql2(21) = Format(Trim(Text(16)), "mm/dd/yyyy")
254         gsSql1(22) = "CATATAN"
255         gsSql2(22) = Trim(Text(13))
            
256         sSQL = GensSql2("DETAILFIR", False)
257         If isBoleh Then
258            Call objO.InsertData(sSQL)
259            MsgBox "Data berhasil disimpan", vbInformation
260         Else
261            MsgBox "Lengkapi data sebagai berikut" & vbCrLf & pesanError, vbInformation
262            Exit Sub
263         End If
264         Unload Me
265      End If
End Sub

Private Sub Command2_Click()
   Unload Me
   
End Sub

Private Sub Form_Activate()
1        If isBooking = True Then
            ''jika dari booking
2           sSQL = "SELECT * FROM DETAILFIR WITH (NOLOCK) WHERE NOFAK = '" & Trim(NoBooking) & "' AND NO_POLISI = '" & Trim(Text(14)) & "' AND NO_CHASSIS = '" & noChassis & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "'"
3        Else
            '' dari wo
4           If NoBooking <> "" Then
5              sSQL = "SELECT * FROM DETAILFIR WITH (NOLOCK) WHERE nowo = '" & noOrder & "' AND NO_POLISI = '" & Trim(Text(14)) & "' AND NO_CHASSIS = '" & noChassis & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "'"
6           Else
7              sSQL = "SELECT * FROM DETAILFIR WITH (NOLOCK) WHERE nowo = '" & noOrder & "' and NOFAK = '" & Trim(NoBooking) & "' AND NO_POLISI = '" & Trim(Text(14)) & "' AND NO_CHASSIS = '" & noChassis & "' AND IDKELUHAN = '" & Trim(Label15.Caption) & "'"
8           End If
            
               
9        End If
10       Set rsTemp = objO.FetchData(sSQL)
11       If Not rsTemp.EOF Then
12          rsTemp.MoveFirst
13          If rsTemp!gejalapertama <> "" Then
14             Select Case rsTemp!gejalapertama
                  Case "Hari Ini"
15                   optGejala(0).value = True
16                Case "Minggu Lalu"
17                   optGejala(1).value = True
18                Case Else
19                   optGejala(2).value = True
20                   Text(2).Enabled = True
21                   Text(2).BackColor = &HFFFFFF
22                   Text(2) = Mid(Trim(rsTemp!gejalapertama), 1, Len(rsTemp!gejalapertama) - 15)
23             End Select
24          Else
25             optGejala(3).value = True
26          End If
            
27          If Trim(rsTemp!frekuensi) <> "" Then
28             Select Case Trim(rsTemp!frekuensi)
                  Case "Sekali"
29                   optFrekuensi(0).value = True
30                Case "Kadang"
31                   optFrekuensi(1).value = True
32                Case Else
33                   optFrekuensi(2).value = True
34             End Select
35          Else
36             optFrekuensi(3).value = True
37          End If
            
38          If Trim(rsTemp!indikator) <> "" Then
39             Select Case Trim(rsTemp!indikator)
                  Case "ON"
40                   optIndikator(0).value = True
41                Case "Selalu"
42                   optIndikator(1).value = True
43                Case Else
44                   optIndikator(2).value = True
45             End Select
46          Else
47             optIndikator(3).value = True
48          End If
            
49           If Trim(rsTemp!saatmil) <> "" Then
50             Select Case Trim(rsTemp!saatmil)
                  Case "Idling"
51                   optKondisi(0).value = True
52                Case "Starting"
53                   optKondisi(1).value = True
54                Case Else
55                   optKondisi(2).value = True
56             End Select
57          Else
58             optKondisi(3).value = True
59          End If
            
60          If Trim(rsTemp!kondisiMesin) <> "" Then
61             Select Case Trim(rsTemp!kondisiMesin)
                  Case "Panas"
62                   optKondisiMesin(0).value = True
63                Case "Dingin"
64                   optKondisiMesin(1).value = True
65                Case Else
66                   optKondisiMesin(2).value = True
67             End Select
68          End If
            Dim hasil() As String
69          If Trim(rsTemp!shiftlevel) <> "" Then
                  
70                hasil() = Split(Trim(rsTemp!shiftlevel), ",")
71                For i = 0 To UBound(hasil)
72                   Select Case hasil(i)
                        Case "1"
73                         Check16.value = vbChecked
74                      Case "2"
75                         Check19.value = vbChecked
76                      Case "3"
77                         Check17.value = vbChecked
78                      Case "4"
79                         Check20.value = vbChecked
80                      Case "5"
81                         Check18.value = vbChecked
82                      Case "P"
83                         Check23.value = vbChecked
84                      Case "R"
85                         Check24.value = vbChecked
86                      Case "N"
87                         Check22.value = vbChecked
88                      Case "D"
89                         Check25.value = vbChecked
90                      Case "S"
91                         Check21.value = vbChecked
92                      Case Else
93                         Check26.value = vbChecked
94                   End Select
95                Next
96          End If
            
97          If Trim(rsTemp!Speed) <> "" Then
98             Text(3) = Trim(rsTemp!Speed)
99          End If
            
100         If Trim(rsTemp!rpm) <> "" Then
101            Text(5) = Trim(rsTemp!rpm)
102         End If
            
103         If Trim(rsTemp!beban) <> "" Then
104            Text(4) = Trim(rsTemp!beban)
105         End If
            
106         If Trim(rsTemp!penumpang) <> "" Then
107            Text(6) = Round(Trim(rsTemp!penumpang), 0)
108         End If
            
109         If Trim(rsTemp!kondisiJalan) <> "" Then
                  
110               hasil() = Split(Trim(rsTemp!kondisiJalan), ",")
111               For i = 0 To UBound(hasil)
112                  Select Case hasil(i)
                        Case Trim(Check46.Caption)
113                        Check46.value = vbChecked
114                     Case Trim(Check45.Caption)
115                        Check45.value = vbChecked
116                     Case Trim(Check44.Caption)
117                        Check44.value = vbChecked
118                     Case Trim(Check47.Caption)
119                        Check47.value = vbChecked
120                     Case Trim(Check43.Caption)
121                        Check43.value = vbChecked
122                     Case Trim(Check42.Caption)
123                        Check42.value = vbChecked
124                     Case Trim(Check41.Caption)
125                        Check41.value = vbChecked
126                     Case Trim(Check48.Caption)
127                        Check48.value = vbChecked
128                     Case Trim(Check40.Caption)
129                        Check40.value = vbChecked
130                     Case Trim(Check39.Caption)
131                        Check39.value = vbChecked
132                     Case Else
133                        Check38.value = vbChecked
134                        Text(7) = Trim(hasil(i))
135                  End Select
136               Next
137         End If
            
138         If Trim(rsTemp!laluLintas) <> "" Then
139            Select Case Trim(rsTemp!laluLintas)
                  Case "Macet Stop & Go"
140                  optLaluLintas(0).value = True
141               Case "Lancar"
142                  optLaluLintas(1).value = True
143               Case Else
144                  optLaluLintas(2).value = True
145            End Select
146         End If
             
147         If Trim(rsTemp!cuaca) <> "" Then
148            Select Case Trim(rsTemp!cuaca)
                  Case Trim(optCuaca(0).Caption)
149                  optCuaca(0).value = True
150               Case Trim(optCuaca(1).Caption)
151                  optCuaca(1).value = True
152               Case Trim(optCuaca(2).Caption)
153                  optCuaca(2).value = True
154               Case Trim(optCuaca(3).Caption)
155                  optCuaca(3).value = True
156               Case Else
157                  optCuaca(4).value = True
158            End Select
159         End If
            
160         If Trim(rsTemp!blowerspeedac) <> "" Then
161            Text(11) = Trim(rsTemp!blowerspeedac)
162         End If
            
163         If Trim(rsTemp!tempsetting) <> "" Then
164            Text(12) = Round(Trim(rsTemp!tempsetting), 0)
165         End If
            
166         If Trim(rsTemp!catatan) <> "" Then
167            Text(13) = Trim(rsTemp!catatan)
168         End If
169      End If
          
End Sub

Private Sub optGejala_Click(Index As Integer)
   If optGejala(2).value = True Then
      Text(2).Enabled = True
      Text(2).BackColor = &HFFFFFF
   Else
      Text(2).Enabled = False
      Text(2).BackColor = &HC0E0FF
   End If
End Sub

Private Sub optLaluLintas_Click(Index As Integer)
   If optLaluLintas(2).value = True Then
      Text(8).BackColor = &HFFFFFF
   Else
      Text(8).BackColor = &HC0E0FF
   End If
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 20
    Else
        Frame2.Top = 0
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 20
    Else
        Frame2.Top = 0
    End If
End Sub

Private Sub VScroll2_Change()
    If VScroll2.value > 0 Then
        Frame3.Top = -VScroll2.value + 20
    Else
        Frame3.Top = 0
    End If
End Sub

Private Sub VScroll2_Scroll()
    If VScroll2.value > 0 Then
        Frame3.Top = -VScroll2.value + 20
    Else
        Frame3.Top = 0
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
   If Text(Index).BackColor = &HC0E0FF Then KeyAscii = 0
End Sub
