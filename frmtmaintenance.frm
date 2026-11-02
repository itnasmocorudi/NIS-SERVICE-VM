VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmtmaintenance 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Maintenance"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame nPerbaikanFI 
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Perbaikan data Final Inspection "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   56
      Top             =   8400
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdPerbaikanFI 
         Caption         =   "Proses !"
         Height          =   375
         Left            =   3000
         TabIndex        =   57
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   795
         TabIndex        =   58
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2055
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   8
         Left            =   4320
         Picture         =   "frmtmaintenance.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diperbaiki"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   61
         Top             =   960
         Width           =   2445
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   60
         Top             =   510
         Width           =   510
      End
   End
   Begin VB.Frame nFinishWash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Menutup proses yang belum di-Finish Washing     "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      TabIndex        =   48
      Top             =   4680
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdWash 
         Caption         =   "Proses !"
         Height          =   375
         Left            =   3000
         TabIndex        =   49
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   795
         TabIndex        =   50
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   2055
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   53
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diselesaikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   960
         Width           =   2625
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   7
         Left            =   4320
         Picture         =   "frmtmaintenance.frx":391B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame nFinishQc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Menutup proses yang belum di-Finish Qc     "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      TabIndex        =   42
      Top             =   3240
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdQc 
         Caption         =   "Proses !"
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   795
         TabIndex        =   44
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox7 
         Height          =   270
         Left            =   2055
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   6
         Left            =   4320
         Picture         =   "frmtmaintenance.frx":7236
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diselesaikan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   46
         Top             =   510
         Width           =   510
      End
   End
   Begin VB.Frame nfNamaSA 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                       Penggantian Nama SA                     "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      TabIndex        =   33
      Top             =   6720
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdSa 
         Caption         =   "Proses !"
         Height          =   375
         Left            =   2520
         TabIndex        =   34
         Top             =   675
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   915
         TabIndex        =   35
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   2175
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   915
         TabIndex        =   39
         Top             =   960
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   2175
         TabIndex        =   40
         Top             =   960
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
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
         Left            =   585
         TabIndex        =   41
         Top             =   990
         Width           =   195
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   5
         Left            =   3840
         Picture         =   "frmtmaintenance.frx":AB51
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   270
         TabIndex        =   38
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diganti Nama SA && Inisial SA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   3780
      End
   End
   Begin VB.Frame nfRtj 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                status WO -> Return Job                "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdRtj 
         Caption         =   "Proses !"
         Height          =   375
         Left            =   2400
         TabIndex        =   26
         Top             =   435
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   795
         TabIndex        =   27
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   2055
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diganti status                                    "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   960
         Width           =   4365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   29
         Top             =   510
         Width           =   510
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   4
         Left            =   3480
         Picture         =   "frmtmaintenance.frx":E46C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame nfPauseJob 
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Menutup proses yang belum di-Finish     "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   21
      Top             =   4560
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdProses 
         Caption         =   "Proses !"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   795
         TabIndex        =   8
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   2055
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   24
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan di-Pause"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   960
         Width           =   2385
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   3
         Left            =   4320
         Picture         =   "frmtmaintenance.frx":11D87
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Frame nfAktifMech 
      BackColor       =   &H00FFFFFF&
      Caption         =   "           Mengaktifkan Jam Mekanik            "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdMech 
         Caption         =   "Aktifkan !"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   795
         TabIndex        =   6
         Top             =   480
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Left            =   1590
         TabIndex        =   18
         Top             =   480
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   2
         Left            =   3240
         Picture         =   "frmtmaintenance.frx":156A2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis Mekanik yang akan diaktifkan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   2460
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mech"
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
         Left            =   285
         TabIndex        =   19
         Top             =   510
         Width           =   375
      End
   End
   Begin VB.Frame nfGanti 
      BackColor       =   &H00FFFFFF&
      Caption         =   "                      Ganti Pekerjaan -> Diagnosis                        "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdGanti 
         Caption         =   "Ganti !"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   0
         Left            =   795
         TabIndex        =   4
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   2055
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   1
         Left            =   4440
         Picture         =   "frmtmaintenance.frx":18FBD
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   16
         Top             =   510
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diganti pekerjaan dengan diagnosis"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   4305
      End
   End
   Begin VB.Frame nfAktifJob 
      BackColor       =   &H00FFFFFF&
      Caption         =   "     Mengaktifkan kembali proses yang sudah Finish     "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdAktif 
         Caption         =   "Aktifkan !"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   795
         TabIndex        =   2
         Top             =   480
         Width           =   1260
         _ExtentX        =   2223
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   2055
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Index           =   0
         Left            =   4320
         Picture         =   "frmtmaintenance.frx":1C8D8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tulis No. WO yang akan diaktifkan lagi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2745
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   150
         TabIndex        =   11
         Top             =   510
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdSublet 
         Caption         =   "Master Sublet"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   8640
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Master Slip Code"
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Top             =   9000
         Width           =   3135
      End
      Begin VB.CommandButton cmdSupplier 
         Caption         =   "Master Supplier"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   8280
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Validasi NEA / Profit Sharing"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   7920
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "KONVERSI"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   9360
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Image btnPerbaikanFI 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":201F3
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   3180
      End
      Begin VB.Image btnFinishWs 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":280E3
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   3180
      End
      Begin VB.Image btnFinishQc 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":2E9EE
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   3180
      End
      Begin VB.Image btnGantiSA 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":37A18
         Stretch         =   -1  'True
         Top             =   6075
         Width           =   3180
      End
      Begin VB.Image btnRtj 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":396D6
         Stretch         =   -1  'True
         Top             =   5205
         Width           =   3180
      End
      Begin VB.Image btnPauseP 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":3B271
         Stretch         =   -1  'True
         Top             =   960
         Width           =   3180
      End
      Begin VB.Image btnAlokasiMech 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":3D023
         Stretch         =   -1  'True
         Top             =   4350
         Width           =   3180
      End
      Begin VB.Image btnDiag 
         Height          =   795
         Left            =   120
         Picture         =   "frmtmaintenance.frx":3EC2B
         Stretch         =   -1  'True
         Top             =   3495
         Width           =   3180
      End
      Begin VB.Image btnAktifPause 
         Height          =   795
         Left            =   105
         Picture         =   "frmtmaintenance.frx":40650
         Stretch         =   -1  'True
         Top             =   105
         Width           =   3180
      End
   End
   Begin VB.Image btnClose 
      Height          =   375
      Left            =   3480
      Picture         =   "frmtmaintenance.frx":42410
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   375
   End
End
Attribute VB_Name = "frmtmaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmtmaintenance
' DateTime  : 07/09/2022 14:27
' Author    : Rudi-VM
' Purpose   :
'---------------------------------------------------------------------------------------

Dim objBilling As New BClsBilling

Private Sub btnAktifPause_Click()
    nfAktifJob.Top = Frame1.Top
    nfAktifJob.Left = Frame1.Left + Frame1.Width + 100
    nfAktifJob.Visible = True
End Sub

Private Sub btnAlokasiMech_Click()
    nfAktifMech.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + btnFinishQc.Height + btnFinishWs.Height + btnDiag.Height + 140
    nfAktifMech.Left = Frame1.Left + Frame1.Width + 100
    nfAktifMech.Visible = True
End Sub

Private Sub btnBatal_Click(Index As Integer)
    Select Case Index
        Case 0
            nfAktifJob.Visible = False
        Case 1
            nfGanti.Visible = False
        Case 2
            nfAktifMech.Visible = False
        Case 3
            nfPauseJob.Visible = False
        Case 4
            nfRtj.Visible = False
        Case 5
            nfNamaSA.Visible = False
        Case 6
            nFinishQc.Visible = False
        Case 7
            nFinishWash.Visible = False
        Case 8
            nPerbaikanFI.Visible = False
    End Select
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDiag_Click()
    nfGanti.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + btnFinishQc.Height + btnFinishWs.Height + 100
    nfGanti.Left = Frame1.Left + Frame1.Width + 100
    nfGanti.Visible = True
End Sub

Private Sub btnFinishQc_Click()
    nFinishQc.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + 50
    nFinishQc.Left = Frame1.Left + Frame1.Width + 100
    nFinishQc.Visible = True
End Sub

Private Sub btnFinishWs_Click()
    nFinishWash.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + btnFinishQc.Height + 50
    nFinishWash.Left = Frame1.Left + Frame1.Width + 100
    nFinishWash.Visible = True
End Sub

Private Sub btnGantiSA_Click()
    nfNamaSA.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + btnFinishQc.Height + btnFinishWs.Height + btnDiag.Height + btnAlokasiMech.Height + btnRtj.Height + 200
    nfNamaSA.Left = Frame1.Left + Frame1.Width + 100
    nfNamaSA.Visible = True
End Sub

Private Sub btnPauseP_Click()
    nfPauseJob.Top = Frame1.Top + btnAktifPause.Height + 50
    nfPauseJob.Left = Frame1.Left + Frame1.Width + 100
    nfPauseJob.Visible = True
End Sub

Private Sub btnPerbaikanFI_Click()
    nPerbaikanFI.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + btnFinishQc.Height + btnFinishWs.Height + btnDiag.Height + btnAlokasiMech.Height + btnRtj.Height + 200
    nPerbaikanFI.Left = Frame1.Left + Frame1.Width + 100
    nPerbaikanFI.Visible = True
End Sub

Private Sub btnRtj_Click()
    nfRtj.Top = Frame1.Top + btnAktifPause.Height + btnPauseP.Height + btnFinishQc.Height + btnFinishWs.Height + btnDiag.Height + btnAlokasiMech.Height + 200
    nfRtj.Left = Frame1.Left + Frame1.Width + 100
    nfRtj.Visible = True
End Sub

Private Sub cmdAktif_Click()
      Dim rsHlp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim rsTemp As ADODB.Recordset
      Dim updateNoStall As String

1         If Trim(Text(1)) <> "" Then
2             Me.Tag = ""
                 
3             frmtmAlasan.show vbModal
              
4             If Me.Tag = "valid" Then
5                 Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(1)) & "' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')")
6                 If Not rsHlp.EOF Then
7                     If Trim(rsHlp("DELIV")) <> "BTL" Then
8                         If Trim(rsHlp("NOINV")) <> "" Then
9                             If Trim(rsHlp("RPR_COND")) = "K" Then
10                                MsgBox ("WO SUDAH CLOSING")
11                            Else
12                                MsgBox ("WO SUDAH DIINVOICE")
13                            End If
14                        Else
                              
15                            Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH(NOLOCK) WHERE FLAG = 'F' AND KETERANGAN LIKE '%|" & Trim(Text(1)) & "|%'")
16                            If Not rsSearch.EOF Then
17                                rsSearch.MoveFirst
18                                updateNoStall = ""
                                  ''2016-06-02 || stall final tidak dapat diubah, harus lewat jpcb
                                  ''--------------------------------------------------------------
                                  ''2016-06-22 || warning itu hanya muncul klo unit yang dimaintenance unit stall hari itu.Tapi klo hari kemaren, dicari ke tempprogress stall terakhir di stall brp, kemudian diupdate nostall nya juga, Selain update flag.
19                                If Trim(rsSearch!noStall) = "STALL FINL" And DateDiff("d", Format(Trim(rsSearch!Tanggal), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) = 0 Then
20                                  MsgBox ("WO dalam stall finish inspection, aktifkan unit lewat JPCB !!")
21                                  Exit Sub
22                                ElseIf Trim(rsSearch!noStall) = "STALL FINL" And DateDiff("d", Format(Trim(rsSearch!Tanggal), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) > 0 Then
23                                  Set rsTemp = objO.FetchData("select TOP 1 NOSTALL from tempprogress where noorder = '" & Trim(Text(1)) & "' AND NOSTALL NOT LIKE '%FINL%' ORDER BY [PAUSE] DESC")
24                                  If Not rsTemp.EOF Then
25                                     updateNoStall = ", NOSTALL = '" & Trim(rsTemp!noStall) & "'"
26                                  End If
27                                End If
28                                erTranAll = False
29                                onTR = True
30                                Call objHT.OpenTransactionBengkel
                                  
31                                Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'P' " & updateNoStall & " WHERE FLAG = 'F' AND KETERANGAN LIKE '%|" & Trim(Text(1)) & "|%'")
32                                Call objHT.UpdateData_BengkelOnly("UPDATE PROGRESS SET FLAG = 'P' WHERE FLAG = 'F' AND NOORDER = '" & Trim(Text(1)) & "'")
33                                Call objHT.UpdateData_BengkelOnly("UPDATE TEMPPROGRESS SET FLAGP = 'P' WHERE FLAGP = 'F' AND NOORDER = '" & Trim(Text(1)) & "'")
34                                Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = 'P',deliv='' WHERE NO_WO = '" & Trim(Text(1)) & "'")
35                                Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','AKTIF FINISH','" & Trim(Text(1)) & "','" & Trim(gBantuan4) & "')")
                                  
36                                If erTranAll = False Then
37                                    Call objHT.CloseTransactionBengkel
38                                    onTR = False
39                                Else
40                                    Call objHT.CloseTransactionBengkel_err
41                                    onTR = False
42                                End If
                                  
43                            End If
44                        End If
45                    End If
46                Else
47                    MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
48                End If
49            End If
              
50        End If
51        nfAktifJob.Visible = False
52        Text(1) = ""
End Sub

Private Sub cmdGanti_Click()
      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset

1         If Trim(Text(0)) <> "" Then
2             Me.Tag = ""
3             frmtmAlasan.show vbModal
              
4             If Me.Tag = "valid" Then
5                 Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')")
6                 If Not rsHlp.EOF Then
7                     If Trim(rsHlp("DELIV")) <> "BTL" Then
8                         If Trim(rsHlp("NOINV")) <> "" Then
9                             If Trim(rsHlp("RPR_COND")) = "K" Then
10                                MsgBox ("WO SUDAH CLOSING")
11                            Else
12                                MsgBox ("WO SUDAH DIINVOICE")
13                            End If
14                        Else
15                            erTranAll = False
16                            onTR = True
17                            Call objHT.OpenTransactionBengkel
                              
18                            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(TglDDMMYYYY(rsHlp("DATE")), "DD/MM/YYYY")) >= 0 Then
19                                Call objHT.UpdateData_BengkelOnly("UPDATE jobdetail SET klpkerja = 'GRP', rep = 'DIAGNOSIS / INSPECTION',rte = 0, actrte = 0 ,OFFn = 0,disc = 0,id_program = '',Price = 0,isFP = '0' WHERE NO_WO = '" & Trim(Text(0)) & "' and rep = '" & NulSID(rsHlp!rep1) & "'")
20                                Call objHT.UpdateData_BengkelOnly("UPDATE [" & dbReplikasiHO & "]..jobdetail SET klpkerja = 'GRP', rep = 'DIAGNOSIS / INSPECTION',rte = 0, actrte = 0 ,OFFn = 0,disc = 0,id_program = '',Price = 0,isFP = '0' WHERE NO_WO = '" & Trim(Text(0)) & "' and rep = '" & NulSID(rsHlp!rep1) & "' and cabang = '" & gDb & "'")
21                            End If
22                            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET LABOR = LABOR - OFF1 WHERE NO_WO = '" & Trim(Text(0)) & "'")
                              'Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET REP1 = 'DIAGNOSIS / INSPECTION', RTE1 = '0', OFF1 = '0' , ACTRATE1 = '0' WHERE NO_WO = '" & Trim(text(0)) & "'")
                              'Informasi WO berubah -> GRP
23                            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET REP1 = 'DIAGNOSIS / INSPECTION', RTE1 = '0', OFF1 = '0' , ACTRATE1 = '0', KLPKERJA1= 'GRP' WHERE NO_WO = '" & Trim(Text(0)) & "'")
24                            Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','GANTI JOB','" & Trim(Text(0)) & "','" & Trim(gBantuan4) & "')")
                              
                              
                              ' UPDATE NILAI REGULAR CHECK DARI JOBDTA
25                            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET FS_RC = '' WHERE NO_WO = '" & Trim(Text(0)) & "'")
                              Dim rsTemp As ADODB.Recordset
                              
26                            If Log.RPLHistori = 0 Then
27                                Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH(NOLOCK) WHERE FS_RC <>''")
28                            Else
29                                Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_REGULAR WITH(NOLOCK) WHERE FS_RC <>''")
30                            End If
31                            If Not rsTemp.BOF Then
32                                rsTemp.MoveFirst
33                                Do While Not rsTemp.EOF
34                                    sSQL = "UPDATE JOBDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_WO ='" & Trim(Text(0)) & "' " _
                                           & " AND REP1+ REP2+ REP3+ REP4+ REP5+ REP6+ REP7 like '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%'"
35                                    Call objHT.UpdateData_BengkelOnly(sSQL)
36                                    rsTemp.MoveNext
37                                Loop
38                            End If
                              ' UPDATE NILAI REGULAR CHECK DARI JOBADD2
                              
39                            If Log.RPLHistori = 0 Then
40                                Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH(NOLOCK) WHERE FS_RC <>''")
41                            Else
42                                Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_REGULAR WITH(NOLOCK) WHERE FS_RC <>''")
43                            End If
44                            If Not rsTemp.BOF Then
45                                rsTemp.MoveFirst
46                                Do While Not rsTemp.EOF
47                                    sSQL = "UPDATE JOBDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_WO ='" & Trim(Text(0)) & "' " _
                                           & " AND NO_WO IN (SELECT NO_WO FROM JOBADD2 WITH(NOLOCK) WHERE REP like '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%')"
48                                    Call objHT.UpdateData_BengkelOnly(sSQL)
49                                    rsTemp.MoveNext
50                                Loop
51                            End If
52                            If erTranAll = False Then
53                                Call objHT.CloseTransactionBengkel
54                                onTR = False
55                            Else
56                                Call objHT.CloseTransactionBengkel_err
57                                onTR = False
58                            End If
59                        End If
60                    End If
61                Else
62                    MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
63                End If
64            End If

65        End If
66        nfGanti.Visible = False
67        Text(0) = ""
End Sub

Private Sub cmdMech_Click()
    If Trim(Text(2)) <> "" Then
        Me.Tag = ""
        erTranAll = False
        onTR = True
        Call objHT.OpenTransactionBengkel
        frmtmAlasan.show vbModal
        
        If Me.Tag = "valid" Then
            Call objHT.UpdateData_BengkelOnly("UPDATE T_MEKANIK SET ALOKASI = 0, TERPAKAI = ABS(TERPAKAI) WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & Trim(Text(2)) & "'")
            Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','AKTIF MECH','" & Trim(Text(2)) & "','" & Trim(gBantuan4) & "')")
        End If
        
        If erTranAll = False Then
            Call objHT.CloseTransactionBengkel
            onTR = False
        Else
            Call objHT.CloseTransactionBengkel_err
            onTR = False
        End If
    End If
    nfAktifMech.Visible = False
    Text(2) = ""
End Sub

Private Sub cmdPerbaikanFI_Click()
      Dim rsHlp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset

1     If Trim(Text(9)) <> "" Then
2         Me.Tag = ""
3         frmtmAlasan.show vbModal
        
4         If Me.Tag = "valid" Then
5           Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(9)) & "' ")
6           If Not rsHlp.EOF Then
7               If Trim(rsHlp("DELIV")) <> "BTL" Then
                  '' =======================================================================================================================================
                  ''  20181008DEV1 ' DEV | MB THANIA    | MENAMBAHKAN PADA MENU MAINTENANCE PERBAIKAN FI
                  '' =======================================================================================================================================
8                 Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH(NOLOCK) WHERE FLAG = 'S' AND noStall like '%fin%' AND KETERANGAN LIKE '%|" & Trim(Text(9)) & "|%'")
9                 If Not rsSearch.EOF Then
10                    updateNoStall = ""
11                    Set rsTemp = objO.FetchData("select TOP 1 NOSTALL from tempprogress where noorder = '" & Trim(Text(9)) & "' AND NOSTALL NOT LIKE '%FINL%' ORDER BY [PAUSE] DESC")
12                    If Not rsTemp.EOF Then
13                       updateNoStall = ", NOSTALL = '" & Trim(rsTemp!noStall) & "'"
14                    End If
15                    erTranAll = False
16                    onTR = True
17                    Call objHT.OpenTransactionBengkel
                      
18                    Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'P' " & updateNoStall & " WHERE FLAG = 'S' AND KETERANGAN LIKE '%|" & Trim(Text(9)) & "|%'")
19                    Call objHT.InsertData_BengkelOnly("INSERT INTO PROGRESS SELECT NOORDER, '" & Trim(rsHlp("NO_CHASSIS")) & "','P',NOSTALL,TANGGAL,CONVERT(VARCHAR(5),DATEADD(n,3,START), 8),'" & Trim(gSA) & "','PAUSE SYSTEM',MECH FROM TEMPPROGRESS WITH(NOLOCK) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(9)) & "'")
20                    Call objHT.UpdateData_BengkelOnly("UPDATE TEMPPROGRESS SET FLAGP = 'P', PAUSE = CONVERT(VARCHAR(5),DATEADD(n,3,START), 8) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(9)) & "'")
21                    Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = 'P' WHERE NO_WO = '" & Trim(Text(9)) & "'")
22                    Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','PERBAIKAN FI','" & Trim(Text(9)) & "','" & Trim(gBantuan4) & "')")
                      
23                    If erTranAll = False Then
24                        Call objHT.CloseTransactionBengkel
25                        MsgBox "Data berhasil diubah", vbInformation
26                        onTR = False
27                    Else
28                        Call objHT.CloseTransactionBengkel_err
29                        MsgBox "Data gagal diubah", vbInformation
30                        onTR = False
31                    End If
                  
32                End If
33              Else
34                MsgBox ("WO tersebut telah batal !!")
35              End If
36          Else
37              MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
38          End If
39      End If
40    End If
41        nPerbaikanFI.Visible = False
42        Text(9) = ""
End Sub

Private Sub cmdProses_Click()
      Dim rsHlp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset

1     If Trim(Text(3)) <> "" Then
2         Me.Tag = ""
3         frmtmAlasan.show vbModal
        
4         If Me.Tag = "valid" Then
5           Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(3)) & "' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')")
6           If Not rsHlp.EOF Then
7               If Trim(rsHlp("DELIV")) <> "BTL" Then
8                   If Trim(rsHlp("NOINV")) <> "" Then
9                       If Trim(rsHlp("RPR_COND")) = "K" Then
10                          MsgBox ("WO SUDAH CLOSING")
11                      Else
12                          MsgBox ("WO SUDAH DIINVOICE")
13                      End If
14                  Else
15                      Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH(NOLOCK) WHERE FLAG = 'S' AND KETERANGAN LIKE '%|" & Trim(Text(3)) & "|%'")
16                      If Not rsSearch.EOF Then
                            
                          ''2017-02-03 || THANIA || stall final tidak dapat diubah, harus lewat jpcb
                          ''--------------------------------------------------------------
                          ''2017-02-03 || THANIA || warning itu hanya muncul klo unit yang dimaintenance unit stall hari itu.Tapi klo hari kemaren, dicari ke tempprogress stall terakhir di stall brp, kemudian diupdate nostall nya juga, Selain update flag.
17                          updateNoStall = ""
18                          If Trim(rsSearch!noStall) = "STALL FINL" And DateDiff("d", Format(Trim(rsSearch!Tanggal), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) = 0 Then
19                              MsgBox ("WO dalam stall finish inspection, aktifkan unit lewat JPCB !!")
20                              Exit Sub
21                           ElseIf Trim(rsSearch!noStall) = "STALL FINL" And DateDiff("d", Format(Trim(rsSearch!Tanggal), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) > 0 Then
22                              Set rsTemp = objO.FetchData("select TOP 1 NOSTALL from tempprogress where noorder = '" & Trim(Text(3)) & "' AND NOSTALL NOT LIKE '%FINL%' ORDER BY [PAUSE] DESC")
23                              If Not rsTemp.EOF Then
24                                 updateNoStall = ", NOSTALL = '" & Trim(rsTemp!noStall) & "'"
25                              End If
26                           End If

27                          erTranAll = False
28                          onTR = True
29                          Call objHT.OpenTransactionBengkel
        
30                          Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'P' " & updateNoStall & " WHERE FLAG = 'S' AND KETERANGAN LIKE '%|" & Trim(Text(3)) & "|%'")
31                          Call objHT.InsertData_BengkelOnly("INSERT INTO PROGRESS SELECT NOORDER, '" & Trim(rsHlp("NO_CHASSIS")) & "','P',NOSTALL,TANGGAL,CONVERT(VARCHAR(5),DATEADD(n,3,START), 8),'" & Trim(gSA) & "','PAUSE SYSTEM',MECH FROM TEMPPROGRESS WITH(NOLOCK) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(3)) & "'")
32                          Call objHT.UpdateData_BengkelOnly("UPDATE TEMPPROGRESS SET FLAGP = 'P', PAUSE = CONVERT(VARCHAR(5),DATEADD(n,3,START), 8) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(3)) & "'")
33                          Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = 'P' WHERE NO_WO = '" & Trim(Text(3)) & "'")
34                          Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','AKTIF START','" & Trim(Text(3)) & "','" & Trim(gBantuan4) & "')")
                            
35                          If erTranAll = False Then
36                              Call objHT.CloseTransactionBengkel
37                              onTR = False
38                          Else
39                              Call objHT.CloseTransactionBengkel_err
40                              onTR = False
41                          End If
                            
42                      End If
43                  End If
44              End If
45          Else
46              MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
47          End If
48      End If

49        End If
50        nfPauseJob.Visible = False
51        Text(3) = ""
End Sub

Private Sub cmdQc_Click()
      Dim rsHlp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim isWashing As Boolean
1         If Trim(Text(7)) <> "" Then
2             Me.Tag = ""
3             frmtmAlasan.show vbModal
              
4             If Me.Tag = "valid" Then
                  ''Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA A WITH(NOLOCK) INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(7)) & "' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')")
                  ''20220907 Menu maintenance selesai FI, tidak dibatasi hanya untuk unit yang sudah di validasi order. sehingga unit yg belum difinish setelah invoice, bisa di selesaikan prosesnya.Menu maintenance selesai FI, tidak dibatasi hanya untuk unit yang sudah di validasi order. sehingga unit yg belum difinish setelah invoice, bisa di selesaikan prosesnya.
                  ''-----------------------------------------------
5                 Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA A WITH(NOLOCK) INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(7)) & "'")
                  
6                 If Not rsHlp.EOF Then
7                 isWashing = IIf(rsHlp("WASHING") = 1, True, False)
8                 isWashing = False
9                     If Trim(rsHlp("DELIV")) <> "BTL" Then
10                        If Trim(rsHlp("NOINV")) <> "" Then
11                            If Trim(rsHlp("RPR_COND")) = "K" Then
12                                MsgBox ("WO SUDAH CLOSING")
13                            Else
                                 '' MsgBox ("WO SUDAH DIINVOICE")
                                 ''2023-01-10 revisi mb thania untuk unit invoice tetap bisa d ubah
14                               GoTo lanjut
15                            End If
16                        Else
lanjut:
17                            Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH(NOLOCK) WHERE FLAG NOT IN ('S','P','N','F','','I','K') AND KETERANGAN LIKE '%|" & Trim(Text(7)) & "|%'")
18                            If Not rsSearch.EOF Then
                                  
19                                erTranAll = False
20                                onTR = True
21                                Call objHT.OpenTransactionBengkel
22                                If Log.dmsNonDss = 1 Then
23                                    Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = '" & IIf(isWashing, "G", "F") & "' WHERE FLAG in ('H','Q','S') AND KETERANGAN LIKE '%|" & Trim(Text(7)) & "|%'")
24                                Else
25                                    Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'F' WHERE FLAG in ('H','Q','S') AND KETERANGAN LIKE '%|" & Trim(Text(7)) & "|%'")
26                                End If
                                 ' Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = '" & IIf(isWashing, "G", "F") & "' WHERE FLAG in ('H','Q') AND KETERANGAN LIKE '%|" & Trim(text(7)) & "|%'")
27                                Call objHT.InsertData_BengkelOnly("INSERT INTO PROGRESS SELECT NOORDER, '" & Trim(rsHlp("NO_CHASSIS")) & "','" & IIf(isWashing, "P", "F") & "',NOSTALL,TANGGAL,CONVERT(VARCHAR(5),DATEADD(n,3,START), 8),'" & Trim(gSA) & "','FINISH WASHING',MECH FROM TEMPPROGRESS WITH(NOLOCK) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(7)) & "'")
28                                If Trim(rsSearch("FLAG")) = "Q" Then
29                                    Call objHT.UpdateData_BengkelOnly("UPDATE TEMPPROGRESS SET FLAGP = '" & IIf(isWashing, "P", "F") & "', PAUSE = CONVERT(VARCHAR(5),DATEADD(n,3,START), 8) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(7)) & "'")
30                                ElseIf Trim(rsSearch("FLAG")) = "H" Then
31                                    Call objHT.InsertData_BengkelOnly("INSERT INTO TEMPPROGRESS VALUES ('" & Trim(rsHlp("NO_WO")) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & Trim(rsSearch("NOSTALL")) & "','" & Format(isnow, "HH:NN") & "','F','" & Format(DateAdd("n", 3, isnow), "HH:NN") & "','" & gSA & "','" & Trim(rsHlp("GROUP")) & "')")
32                                End If
33                                Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = '" & IIf(isWashing, "P", "S") & "', DELIV = '" & IIf(isWashing, "", "OKE") & "' WHERE NO_WO = '" & Trim(Text(7)) & "'")
34                                Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','FINISH WASHING','" & Trim(Text(7)) & "','" & Trim(gBantuan4) & "')")
                                  
35                                If erTranAll = False Then
36                                    Call objHT.CloseTransactionBengkel
37                                    onTR = False
38                                Else
39                                    Call objHT.CloseTransactionBengkel_err
40                                    onTR = False
41                                End If
42                            End If
43                        End If
44                    End If
45                Else
46                    MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
47                End If
48            End If
              
49        End If
50        nFinishQc.Visible = False
51        Text(7) = ""
End Sub

Private Sub cmdRtj_Click()
Dim rsHlp As ADODB.Recordset

    If Trim(Text(4)) <> "" Then
        Me.Tag = ""
        frmtmAlasan.show vbModal
        
        If Me.Tag = "valid" Then
            Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(4)) & "' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')")
            If Not rsHlp.EOF Then
                If Trim(rsHlp("DELIV")) <> "BTL" Then
                    If Trim(rsHlp("NOINV")) <> "" Then
                        If Trim(rsHlp("RPR_COND")) = "K" Then
                            MsgBox ("WO SUDAH CLOSING")
                        Else
                            MsgBox ("WO SUDAH DIINVOICE")
                        End If
                    Else
                    
                        erTranAll = False
                        onTR = True
                        Call objHT.OpenTransactionBengkel
                        
                        Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET FN_RETURN = '1' WHERE NO_WO = '" & Trim(Text(4)) & "'")
                        Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','AKTIF RTJ','" & Trim(Text(4)) & "','" & Trim(gBantuan4) & "')")
                        
                        If erTranAll = False Then
                            Call objHT.CloseTransactionBengkel
                            onTR = False
                        Else
                            Call objHT.CloseTransactionBengkel_err
                            onTR = False
                        End If
                    End If
                End If
            Else
                MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
            End If
        End If

    End If
    nfRtj.Visible = False
    Text(4) = ""
End Sub

Private Sub cmdSa_Click()
Dim rsHlp As ADODB.Recordset
    
    If Trim(Text(5)) <> "" And Trim(Text(6)) <> "" Then
        Me.Tag = ""
        frmtmAlasan.show vbModal
        
        If Me.Tag = "valid" Then
                    
            erTranAll = False
            onTR = True
            Call objHT.OpenTransactionBengkel
            
            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET FRONT = '" & Trim(Text(6)) & "' WHERE NO_WO = '" & Trim(Text(5)) & "'")
            If Log.RPLHistori = 1 Then
                Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET FRONT = '" & Trim(Text(6)) & "' WHERE NO_WO = '" & Trim(Text(5)) & "'")
            End If
            Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET FRONT = '" & Trim(Text(6)) & "' WHERE NO_WO = '" & Trim(Text(5)) & "'")
            Call objHT.UpdateData_BengkelOnly("UPDATE HISTORY SET FRONT = '" & Trim(Text(6)) & "' WHERE NO_WO = '" & Trim(Text(5)) & "'")
            Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','PERGANTIAN SA','" & Trim(Text(5)) & "','" & Trim(gBantuan4) & "')")
                        
            If erTranAll = False Then
                Call objHT.CloseTransactionBengkel
                onTR = False
            Else
                Call objHT.CloseTransactionBengkel_err
                onTR = False
            End If
        End If
    End If
    nfNamaSA.Visible = False
    Text(5) = ""
    Text(6) = ""
End Sub

Private Sub cmdSublet_Click()
    If Log.standarisasiSublet = 1 Then
        frmTSubletStandarisasi.show vbModal
        frmTSubletStandarisasi.WindowState = vbMaximized
    End If
End Sub

Private Sub cmdSupplier_Click()
    If Log.standarisasiSublet = 1 Then
        frmTSupplierStandarisasi.show vbModal
        frmTSupplierStandarisasi.WindowState = vbMaximized
    Else
        frmtSupplier.show
        frmtSupplier.WindowState = vbMaximized
    End If
    
End Sub

Private Sub cmdWash_Click()
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset

    If Trim(Text(8)) <> "" Then
        Me.Tag = ""
        frmtmAlasan.show vbModal
        
        If Me.Tag = "valid" Then
            Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(8)) & "' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '')")
            If Not rsHlp.EOF Then
                If Trim(rsHlp("DELIV")) <> "BTL" Then
                    If Trim(rsHlp("NOINV")) <> "" Then
                        If Trim(rsHlp("RPR_COND")) = "K" Then
                            MsgBox ("WO SUDAH CLOSING")
                        Else
                            MsgBox ("WO SUDAH DIINVOICE")
                        End If
                    Else
                        Set rsSearch = objO.FetchData("SELECT * FROM STALL WITH(NOLOCK) WHERE FLAG NOT IN ('S','P','N','F','','I','K') AND KETERANGAN LIKE '%|" & Trim(Text(8)) & "|%'")
                        If Not rsSearch.EOF Then
                        
                            erTranAll = False
                            onTR = True
                            Call objHT.OpenTransactionBengkel
                        
                            Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'F' WHERE FLAG IN ('A','G','Q') AND KETERANGAN LIKE '%|" & Trim(Text(8)) & "|%'")
                            Call objHT.InsertData_BengkelOnly("INSERT INTO PROGRESS VALUES ('" & Trim(rsHlp("NO_WO")) & "','" & Trim(rsHlp("NO_CHASSIS")) & "','F','" & Trim(rsSearch("NOSTALL")) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & Format(isnow, "HH:NN") & "','" & Trim(gSA) & "','SELESAI WASHING','" & Trim(gSA) & "')")
                            If Trim(rsSearch("FLAG")) = "A" Then
                                Call objHT.UpdateData_BengkelOnly("UPDATE TEMPPROGRESS SET FLAGP = 'F', PAUSE = CONVERT(VARCHAR(5),DATEADD(n,3,START), 8) WHERE FLAGP = '' AND NOORDER = '" & Trim(Text(8)) & "'")
                            ElseIf Trim(rsSearch("FLAG")) = "G" Then
                                Call objHT.InsertData_BengkelOnly("INSERT INTO TEMPPROGRESS VALUES ('" & Trim(rsHlp("NO_WO")) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & Trim(rsSearch("NOSTALL")) & "','" & Format(isnow, "HH:NN") & "','F','" & Format(DateAdd("n", 3, isnow), "HH:NN") & "','" & gSA & "','" & Trim(rsHlp("GROUP")) & "')")
                            End If
                            Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = 'S', DELIV = '" & IIf(rsHlp("FS_RC") <> "", "PRC", "OKE") & "' WHERE NO_WO = '" & Trim(Text(8)) & "'")
                            Call objHT.InsertData_BengkelOnly("INSERT INTO MAINTENANCE VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY") & "','FINISH WASHING','" & Trim(Text(8)) & "','" & Trim(gBantuan4) & "')")
                                     
                            If erTranAll = False Then
                                Call objHT.CloseTransactionBengkel
                                onTR = False
                            Else
                                Call objHT.CloseTransactionBengkel_err
                                onTR = False
                            End If
                            
                        End If
                    End If
                End If
            Else
                MsgBox ("WO tersebut belum terdaftar atau sudah divalidasi !!")
            End If
        End If
        
    End If
    nFinishWash.Visible = False
    Text(8) = ""
End Sub

Private Sub Command1_Click()
    frmKonversi.show
    frmKonversi.WindowState = vbMaximized
End Sub

Private Sub Command2_Click()
    frmtNEA_PS.show
    frmtNEA_PS.WindowState = vbMaximized
End Sub

Private Sub Command3_Click()
    frmtSlipCode.show
    frmtSlipCode.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    If gTipeUser = "ADMIN" Then
        Command1.Visible = True
    End If
    If Log.standarisasiSublet = 1 Then
        cmdSublet.Visible = True
        Command3.Top = 9000
        Command1.Top = 9360
    Else
        cmdSublet.Visible = False
        Command3.Top = 8640
        Command1.Top = 9000
    End If
    
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index).Mask = ""
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
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = ""
                frmCari.show (vbModal)
            Case 6
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 2
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT isnull(MENITKERJA,0) AS MENITKERJA, ISNULL(TERPAKAI,0) AS TERPAKAI, ISNULL(ALOKASI, 0) AS ALOKASI, LEVMECH FROM mekanik A WITH(NOLOCK) LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND B.HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND B.TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' WHERE A.mech = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MEKANIK"
                    gBantuan = "ALL"
                    gBantuan4 = ""
                    frmCari.show (vbModal)
                End If
            End If
        Case 6
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.show (vbModal)
                End If
            End If
    End Select
End Sub
