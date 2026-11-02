VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHONotifikasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notifikasi Pelanggan"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12660
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame2 
         Caption         =   "Input Booking"
         Height          =   7575
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11805
         Begin VB.Frame Frame3 
            Height          =   4080
            Left            =   5580
            TabIndex        =   27
            Top             =   180
            Width           =   6120
            Begin VB.Frame Frame9 
               Caption         =   "                                           Kondisi ketika terjadi Keluhan                                          "
               Height          =   2655
               Left            =   0
               TabIndex        =   28
               Top             =   1440
               Width           =   6135
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   43
                  Left            =   1005
                  TabIndex        =   29
                  Top             =   360
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   25
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   44
                  Left            =   1005
                  TabIndex        =   30
                  Top             =   720
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   25
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   45
                  Left            =   1005
                  TabIndex        =   31
                  Top             =   1080
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   25
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   46
                  Left            =   1005
                  TabIndex        =   32
                  Top             =   1440
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   25
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   47
                  Left            =   1005
                  TabIndex        =   33
                  Top             =   1800
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   25
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   48
                  Left            =   1005
                  TabIndex        =   34
                  Top             =   2160
                  Width           =   2850
                  _ExtentX        =   5027
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   12640511
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   25
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu"
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
                  TabIndex        =   39
                  Top             =   405
                  Width           =   465
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Frekuensi"
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
                  TabIndex        =   38
                  Top             =   765
                  Width           =   690
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lokasi"
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
                  Left            =   480
                  TabIndex        =   37
                  Top             =   1125
                  Width           =   435
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lampu Peringatan"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   105
                  TabIndex        =   36
                  Top             =   1485
                  Width           =   810
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
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
                  Height          =   195
                  Left            =   360
                  TabIndex        =   35
                  Top             =   2040
                  Width           =   555
               End
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   29
               Left            =   1185
               TabIndex        =   40
               Top             =   195
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   423
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
               Height          =   240
               Index           =   30
               Left            =   1185
               TabIndex        =   41
               Top             =   420
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   423
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
               Height          =   240
               Index           =   31
               Left            =   1185
               TabIndex        =   42
               Top             =   645
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   423
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
               Height          =   240
               Index           =   32
               Left            =   1185
               TabIndex        =   43
               Top             =   870
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   423
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
               Height          =   240
               Index           =   33
               Left            =   1185
               TabIndex        =   44
               Top             =   1095
               Width           =   4800
               _ExtentX        =   8467
               _ExtentY        =   423
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
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Keluhan && Permintaan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   90
               TabIndex        =   45
               Top             =   225
               Width           =   945
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1095
            Left            =   0
            TabIndex        =   15
            Top             =   180
            Width           =   5520
            Begin VB.CheckBox chkThs 
               Caption         =   "THS"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4080
               TabIndex        =   17
               Top             =   360
               Width           =   1335
            End
            Begin VB.CheckBox chkRtn 
               Caption         =   "Return Job"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4080
               TabIndex        =   16
               Top             =   120
               Width           =   1335
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   26
               Left            =   1260
               TabIndex        =   18
               Top             =   120
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">99-99-9999"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   27
               Left            =   1260
               TabIndex        =   19
               Top             =   435
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   28
               Left            =   1260
               TabIndex        =   20
               Top             =   750
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
               Mask            =   ">CC"
               PromptChar      =   ">"
            End
            Begin VB.Label Label20 
               Caption         =   "(HH:MM)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2025
               TabIndex        =   26
               Top             =   495
               Width           =   1140
            End
            Begin VB.Label Label30 
               Caption         =   "(DD-MM-YYYY)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2535
               TabIndex        =   25
               Top             =   142
               Width           =   1140
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "D [ Datang ] / T [Telepon] / S [Sms]"
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
               Left            =   1710
               TabIndex        =   24
               Top             =   795
               Width           =   2550
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Via"
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
               Left            =   1005
               TabIndex        =   23
               Top             =   795
               Width           =   210
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jam Booking"
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
               TabIndex        =   22
               Top             =   480
               Width           =   885
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Booking"
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
               Left            =   60
               TabIndex        =   21
               Top             =   165
               Width           =   1170
            End
         End
         Begin VB.Frame Frame5 
            Height          =   2670
            Left            =   90
            TabIndex        =   10
            Top             =   4300
            Width           =   11610
            Begin MSFlexGridLib.MSFlexGrid DGrid 
               Height          =   2265
               Left            =   120
               TabIndex        =   11
               Top             =   360
               Width           =   5625
               _ExtentX        =   9922
               _ExtentY        =   3995
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               ScrollBars      =   1
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid PartGrid 
               Height          =   2265
               Left            =   5760
               TabIndex        =   12
               Top             =   360
               Width           =   5745
               _ExtentX        =   10134
               _ExtentY        =   3995
               _Version        =   393216
               FixedCols       =   0
               HighLight       =   2
               ScrollBars      =   1
               SelectionMode   =   1
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " Keluhan && Order Booking"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   120
               Width           =   1845
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order Part"
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
               Left            =   5880
               TabIndex        =   13
               Top             =   120
               Width           =   765
            End
         End
         Begin VB.CommandButton TblKeluar 
            BackColor       =   &H00C0C0C0&
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
            Left            =   9840
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   7080
            Width           =   1890
         End
         Begin VB.CommandButton TblSimpan 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Batal"
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
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   7080
            Width           =   1890
         End
         Begin VB.CommandButton TblBatal 
            Caption         =   "&Batal"
            Height          =   390
            Left            =   8880
            TabIndex        =   7
            Top             =   10200
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Frame Frame7 
            Enabled         =   0   'False
            Height          =   855
            Left            =   120
            TabIndex        =   3
            Top             =   3240
            Width           =   2175
            Begin VB.OptionButton opSvc 
               Caption         =   "Ditunggu"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   480
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opSvc 
               Caption         =   "Ditinggal"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   4
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SERVICE"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   600
               TabIndex        =   6
               Top             =   120
               Width           =   1065
            End
         End
         Begin VB.ComboBox cboPelanggan 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            ItemData        =   "frmHONotifikasi.frx":0000
            Left            =   1200
            List            =   "frmHONotifikasi.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1920
            Width           =   1215
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   435
            Index           =   37
            Left            =   1200
            TabIndex        =   46
            Top             =   2325
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   767
            _Version        =   393216
            BackColor       =   12640511
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   405
            Index           =   39
            Left            =   3165
            TabIndex        =   47
            Top             =   2400
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   405
            Index           =   40
            Left            =   3165
            TabIndex        =   48
            Top             =   2805
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   405
            Index           =   41
            Left            =   3165
            TabIndex        =   49
            Top             =   3225
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   405
            Index           =   42
            Left            =   3165
            TabIndex        =   50
            Top             =   3840
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   38
            Left            =   1200
            TabIndex        =   51
            Top             =   2760
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   36
            Left            =   2400
            TabIndex        =   52
            Top             =   1920
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   476
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
            Index           =   34
            Left            =   1200
            TabIndex        =   53
            Top             =   1335
            Width           =   4260
            _ExtentX        =   7514
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
            Index           =   35
            Left            =   1200
            TabIndex        =   54
            Top             =   1605
            Width           =   4260
            _ExtentX        =   7514
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JASA"
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
            Left            =   2625
            TabIndex        =   64
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label28 
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
            Left            =   915
            TabIndex        =   63
            Top             =   2805
            Width           =   195
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "________________________  +"
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
            Left            =   3195
            TabIndex        =   62
            Top             =   3555
            Width           =   2370
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estimasi Selesai  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   390
            TabIndex        =   61
            Top             =   2325
            Width           =   720
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PART"
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
            Left            =   2610
            TabIndex        =   60
            Top             =   2880
            Width           =   390
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUBLET"
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
            Left            =   2460
            TabIndex        =   59
            Top             =   3330
            Width           =   540
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
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
            Left            =   2520
            TabIndex        =   58
            Top             =   3960
            Width           =   480
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Karakter"
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
            Left            =   495
            TabIndex        =   57
            Top             =   1980
            Width           =   615
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Saran"
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
            Left            =   690
            TabIndex        =   56
            Top             =   1650
            Width           =   420
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pesan Khusus"
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
            Left            =   120
            TabIndex        =   55
            Top             =   1380
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frmHONotifikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
