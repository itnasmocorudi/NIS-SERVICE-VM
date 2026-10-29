VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRetensiAsuransiCallHasil 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   Icon            =   "frmRetensiAsuransiCallHasil.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmRetensiAsuransiCallHasil.frx":29C12
   ScaleHeight     =   10500
   ScaleWidth      =   12060
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   10000
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   94
         Top             =   1440
         Width           =   2610
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   7380
         MaxLength       =   20
         TabIndex        =   90
         Top             =   1425
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   7380
         MaxLength       =   20
         TabIndex        =   75
         Top             =   1095
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   8
         Left            =   1545
         MaxLength       =   2
         TabIndex        =   74
         Top             =   1455
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   9
         Left            =   2055
         MaxLength       =   4
         TabIndex        =   73
         Top             =   1455
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   10
         Left            =   2805
         MaxLength       =   3
         TabIndex        =   72
         Top             =   1455
         Visible         =   0   'False
         Width           =   635
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   71
         Top             =   465
         Width           =   3885
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1545
         MaxLength       =   50
         TabIndex        =   70
         Top             =   1125
         Width           =   3885
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   69
         Top             =   795
         Width           =   2610
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3180
         MaxLength       =   20
         TabIndex        =   68
         Top             =   1845
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   7380
         MaxLength       =   210
         TabIndex        =   67
         Top             =   465
         Width           =   4005
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7380
         MaxLength       =   300
         TabIndex        =   66
         Top             =   795
         Width           =   4005
      End
      Begin VB.PictureBox PicMove 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   105
         ScaleHeight     =   7095
         ScaleWidth      =   11235
         TabIndex        =   1
         Top             =   2010
         Width           =   11235
         Begin VB.Frame frDetailPolish 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            TabIndex        =   100
            Top             =   6120
            Width           =   11175
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   13
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   101
               Top             =   180
               Width           =   3885
            End
            Begin TDBDate6Ctl.TDBDate tglpolish1 
               Height          =   330
               Left            =   1920
               TabIndex        =   102
               Top             =   540
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "frmRetensiAsuransiCallHasil.frx":2E8FD
               Caption         =   "frmRetensiAsuransiCallHasil.frx":2EA29
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRetensiAsuransiCallHasil.frx":2EA95
               Keys            =   "frmRetensiAsuransiCallHasil.frx":2EAB3
               Spin            =   "frmRetensiAsuransiCallHasil.frx":2EB11
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   -1
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
            Begin TDBDate6Ctl.TDBDate tglpolish2 
               Height          =   330
               Left            =   3825
               TabIndex        =   103
               Top             =   540
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "frmRetensiAsuransiCallHasil.frx":2EB39
               Caption         =   "frmRetensiAsuransiCallHasil.frx":2EC65
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmRetensiAsuransiCallHasil.frx":2ECD1
               Keys            =   "frmRetensiAsuransiCallHasil.frx":2ECEF
               Spin            =   "frmRetensiAsuransiCallHasil.frx":2ED4D
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   -1
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
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               Caption         =   "Detail Polish"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   109
               Top             =   0
               Width           =   1215
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H000080FF&
               Height          =   855
               Left            =   30
               Top             =   120
               Width           =   11085
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   1800
               TabIndex        =   108
               Top             =   240
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Polis"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   107
               Top             =   240
               Width           =   780
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Polis"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   106
               Top             =   615
               Width           =   1305
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3555
               TabIndex        =   105
               Top             =   615
               Width           =   90
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   1800
               TabIndex        =   104
               Top             =   615
               Width           =   60
            End
         End
         Begin VB.Frame FrThird 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   105
            TabIndex        =   55
            Top             =   4500
            Width           =   11250
            Begin VB.ComboBox cboPos 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   7440
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   1125
               Width           =   3495
            End
            Begin VB.ComboBox cboPos 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   3795
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   1125
               Width           =   3495
            End
            Begin VB.ComboBox cboPos 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   165
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   1125
               Width           =   3495
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Terorisme && sabotase"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   5835
               TabIndex        =   59
               Top             =   405
               Width           =   3570
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Gempa bumi, tsunami"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   5820
               TabIndex        =   58
               Top             =   105
               Width           =   3570
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Huru hara && kerusuhan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   57
               Top             =   405
               Width           =   3570
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Banjir termasuk angin topan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   56
               Top             =   105
               Width           =   3570
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kecelakaan diri penumpang"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   7440
               TabIndex        =   62
               Top             =   885
               Width           =   2715
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kecelakaan diri driver"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   3795
               TabIndex        =   61
               Top             =   900
               Width           =   2175
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pihak ke-3"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   165
               TabIndex        =   60
               Top             =   885
               Width           =   1050
            End
         End
         Begin VB.Frame FrJenis 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1545
            Left            =   105
            TabIndex        =   29
            Top             =   2985
            Width           =   11250
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   1080
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   765
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   450
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   150
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   1080
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   765
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   450
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   120
               Width           =   2325
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   7000
               TabIndex        =   54
               Top             =   1125
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 8"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   5955
               TabIndex        =   53
               Top             =   1125
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   7000
               TabIndex        =   51
               Top             =   810
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 7"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   5950
               TabIndex        =   50
               Top             =   810
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   7000
               TabIndex        =   48
               Top             =   495
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 6"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   5950
               TabIndex        =   47
               Top             =   495
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   7000
               TabIndex        =   45
               Top             =   195
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 5"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   5950
               TabIndex        =   44
               Top             =   195
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   1260
               TabIndex        =   42
               Top             =   1125
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 4"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   150
               TabIndex        =   41
               Top             =   1125
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   1260
               TabIndex        =   39
               Top             =   810
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 3"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   150
               TabIndex        =   38
               Top             =   810
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   1260
               TabIndex        =   36
               Top             =   495
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 2"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   150
               TabIndex        =   35
               Top             =   495
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   1260
               TabIndex        =   33
               Top             =   195
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 1"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   150
               TabIndex        =   32
               Top             =   195
               Width           =   780
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               Caption         =   "Jenis Asuransi "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   75
               TabIndex        =   30
               Top             =   -45
               Width           =   1470
            End
            Begin VB.Shape Shape6 
               BorderColor     =   &H000080FF&
               Height          =   1455
               Left            =   30
               Top             =   45
               Width           =   11085
            End
         End
         Begin VB.Frame FrTenor 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   90
            TabIndex        =   20
            Top             =   1995
            Width           =   11250
            Begin VB.ComboBox cboTenor 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   165
               Width           =   1140
            End
            Begin VB.ComboBox cboLea 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1425
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   165
               Width           =   2325
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DcAsuransi 
               Height          =   315
               Left            =   1425
               TabIndex        =   93
               Tag             =   "Kode Supplier TAM"
               Top             =   495
               Width           =   3090
               BevelType       =   0
               _Version        =   196616
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   12648447
               BackColorOdd    =   16777215
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   5450
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H000080FF&
               Height          =   900
               Left            =   30
               Top             =   30
               Width           =   11100
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   7000
               TabIndex        =   28
               Top             =   225
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   1275
               TabIndex        =   27
               Top             =   540
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   1275
               TabIndex        =   26
               Top             =   225
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tenor"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   5950
               TabIndex        =   23
               Top             =   210
               Width           =   570
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Asuransi"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   165
               TabIndex        =   22
               Top             =   540
               Width           =   855
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Leasing"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   165
               TabIndex        =   21
               Top             =   225
               Width           =   750
            End
         End
         Begin VB.Frame FrFollow 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1830
            Left            =   105
            TabIndex        =   6
            Top             =   135
            Width           =   11250
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   1440
               MaxLength       =   200
               TabIndex        =   19
               Top             =   945
               Width           =   9540
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   390
               Left            =   1440
               TabIndex        =   14
               Top             =   540
               Width           =   3450
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tidak Deal"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   3
                  Left            =   1455
                  TabIndex        =   16
                  Top             =   0
                  Width           =   2250
               End
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Deal"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   2
                  Left            =   0
                  TabIndex        =   15
                  Top             =   0
                  Width           =   1410
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   390
               Left            =   1440
               TabIndex        =   11
               Top             =   135
               Width           =   3450
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tidak Terhubung"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   1
                  Left            =   1455
                  TabIndex        =   13
                  Top             =   0
                  Width           =   2250
               End
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Terhubung"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   0
                  Left            =   0
                  TabIndex        =   12
                  Top             =   0
                  Width           =   1410
               End
            End
            Begin MSAdodcLib.Adodc AdoAs 
               Height          =   330
               Left            =   6075
               Top             =   120
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
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   1290
               TabIndex        =   18
               Top             =   1005
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   1290
               TabIndex        =   17
               Top             =   638
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   1290
               TabIndex        =   10
               Top             =   233
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alasan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   165
               TabIndex        =   9
               Top             =   1005
               Width           =   660
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hasil"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   165
               TabIndex        =   8
               Top             =   638
               Width           =   480
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Follow Up"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   165
               TabIndex        =   7
               Top             =   233
               Width           =   945
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H000080FF&
               Height          =   1710
               Left            =   30
               Top             =   75
               Width           =   11085
            End
         End
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   8955
         TabIndex        =   2
         Top             =   9360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&s"
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
         MICON           =   "frmRetensiAsuransiCallHasil.frx":2ED75
         PICN            =   "frmRetensiAsuransiCallHasil.frx":2ED91
         PICH            =   "frmRetensiAsuransiCallHasil.frx":2F2A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   10200
         TabIndex        =   3
         Top             =   9360
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
         MICON           =   "frmRetensiAsuransiCallHasil.frx":2F846
         PICN            =   "frmRetensiAsuransiCallHasil.frx":2F862
         PICH            =   "frmRetensiAsuransiCallHasil.frx":2FD36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdValidasi 
         Height          =   495
         Left            =   7680
         TabIndex        =   99
         Top             =   9360
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&f"
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
         MICON           =   "frmRetensiAsuransiCallHasil.frx":30281
         PICN            =   "frmRetensiAsuransiCallHasil.frx":3029D
         PICH            =   "frmRetensiAsuransiCallHasil.frx":33E83
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   98
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   97
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1425
         TabIndex        =   96
         Top             =   1500
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   360
         TabIndex        =   95
         Top             =   1500
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7245
         TabIndex        =   92
         Top             =   1485
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   6165
         TabIndex        =   91
         Top             =   1485
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   7245
         TabIndex        =   89
         Top             =   1155
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6165
         TabIndex        =   88
         Top             =   1155
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1425
         TabIndex        =   87
         Top             =   1518
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   86
         Top             =   1875
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1425
         TabIndex        =   85
         Top             =   528
         Width           =   60
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoRangka"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   84
         Top             =   528
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   83
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1425
         TabIndex        =   82
         Top             =   1215
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1965
         TabIndex        =   81
         Top             =   1905
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3045
         TabIndex        =   80
         Top             =   1905
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6165
         TabIndex        =   79
         Top             =   855
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   7245
         TabIndex        =   78
         Top             =   855
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   7245
         TabIndex        =   77
         Top             =   525
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6165
         TabIndex        =   76
         Top             =   525
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   6000
         TabIndex        =   5
         Top             =   9480
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".::Poses Call Customer ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4567
         TabIndex        =   4
         Top             =   60
         Width           =   2400
      End
      Begin VB.Line Line24 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11520
         Y1              =   9960
         Y2              =   9960
      End
      Begin VB.Line Line23 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11520
         Y1              =   9240
         Y2              =   9240
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11295
      End
   End
End
Attribute VB_Name = "frmRetensiAsuransiCallHasil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
Private rsFind      As New ADODB.Recordset
Private rsOli       As New ADODB.Recordset
Private rsGridJob   As New ADODB.Recordset
Private rsGridPart  As New ADODB.Recordset
Private rsD         As New ADODB.Recordset
Private rsProg      As New ADODB.Recordset
Private rsDsc       As New ADODB.Recordset
Private rso         As ADODB.Recordset
Private rsP         As ADODB.Recordset
Private isShow      As String
Private isAkt       As String
Private ntxt(12)    As String
Private nWOHO       As String
Private isModel     As String
Private nBook       As String
Private nProg       As String
Private ToTPart     As Double
Private TotOli      As Double
Private nPart       As Double
Private nJob        As Double
Private DiscProg    As Double
Private isBtl       As Boolean
Private isClose     As Boolean
Private ValidProg   As Boolean
Private n           As Integer
Private nKey        As Integer
Private rsRepl      As New ADODB.Recordset
Private sSQL        As String
Private Tri         As New ADODB.Recordset
Private KodeReQ     As String
Private isFalse     As String
Private idAsuransi  As Integer
Private isDataValid As Boolean

Public Property Let myShow(value As String)
    isShow = value
End Property
Public Property Let myAkt(value As String)
    isAkt = value
End Property
Public Property Let MyData(value As ADODB.Recordset)
    Set rsD = value
End Property

Public Property Let MyIdAsuransi(value As Integer)
    idAsuransi = value
End Property


Private Sub cmdSave_Click()
1       On Error GoTo Erro
2       If isValid = True Then
3           Select Case UCase(isAkt)
                Case "NEW CUSTOMER"
4                   KodeReQ = "": Set rso = Nothing: Set rso = New ADODB.Recordset: rsD.MoveFirst
                     
5                   AdaBegin = False: KetEro = ""
6                   Call objHT.OpenTransactionBengkel
7                   Call objHT.OpenTransactionHO
8                   AdaBegin = True
9                   sSQL = "UPDATE FUASURANSI SET NoPolish = '" & NulSID(txt(13).Text) & "', KodeAsuransi = " & CInt(DcAsuransi.Columns(2).value) & ",MasaberlakuAwal = '" & Format(tglpolish1.value, "mm/dd/yyyy") & "',MasaberlakuAkhir = '" & Format(tglpolish2.value, "mm/dd/yyyy") & "',Tenor = " & CInt(cboTenor.Text) & ", " & _
                               "JenisAsuransi1 = '" & s2d(Trim(isAs(cbojns(0)))) & "',JenisAsuransi2 = '" & s2d(Trim(isAs(cbojns(1)))) & "',JenisAsuransi3 = '" & s2d(Trim(isAs(cbojns(2)))) & "',JenisAsuransi4 = '" & s2d(Trim(isAs(cbojns(3)))) & "',JenisAsuransi5 = '" & s2d(Trim(isAs(cbojns(6)))) & "',JenisAsuransi6 = '" & s2d(Trim(isAs(cbojns(5)))) & "',JenisAsuransi7 = '" & s2d(Trim(isAs(cbojns(6)))) & "', " & _
                               "JenisAsuransi8 = '" & s2d(Trim(isAs(cbojns(7)))) & "',Terhubung = '" & IIf(Opt(0).value = True, 1, 0) & "',Deal = '" & IIf(Opt(2).value = True, 1, 0) & "',Alasan = '" & Trim(s2d(txt(11).Text)) & "' WHERE KODEREQ = '" & NulSID(rsD!KodeReQ) & "'"
                                
10                  Call objHT.InsertData_Bengkel(sSQL)
11                  sSQL = ""
12                  If Opt(2).value = True Then
13                      sSQL = "UPDATE OSS..Asuransi SET NoPolis = '" & NulSID(txt(13).Text) & "',MasaberlakuAwal = '" & Format(tglpolish1.value, "mm/dd/yyyy") & "',MasaberlakuAkhir = '" & Format(tglpolish2.value, "mm/dd/yyyy") & "', " & _
                               " banjir = '" & Ch(0).value & "',Gempa = '" & Ch(2).value & "',SRCC = '" & Ch(1).value & "',Teroris = '" & Ch(3).value & "'," & _
                               "TanggungP3 = " & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & ",CelakaDiri = " & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "," & _
                               "areaAsuransi = '" & Trim(DcAsuransi.Columns(1).value) & "',tenorAsuransi = '" & CInt(cboTenor.Text) & "',KodeAsuransi = '" & Trim(DcAsuransi.Columns(2).value) & "',JenisAsuransi = '" & isXs(cbojns(0).Text) & "',JenisAsuransi2 = '" & isXs(cbojns(1).Text) & "',JenisAsuransi3 = '" & isXs(cbojns(2).Text) & "',JenisAsuransi4 = '" & isXs(cbojns(3).Text) & "',JenisAsuransi5 = '" & isXs(cbojns(4).Text) & "',JenisAsuransi6 = '" & isXs(cbojns(5).Text) & "',JenisAsuransi7 = '" & isXs(cbojns(6).Text) & "',JenisAsuransi8 = '" & isXs(cbojns(7).Text) & "',CelakaDiriPass = " & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & " " & _
                               " WHERE NOSPK = '" & NulSID(rsD!KodeReQ) & "' and norangka = '" & NulSID(rsD!noRangka) & "'"
                               
14                        Call objHT.InsertData_Ho(sSQL)
15                  Else
                        
16                  End If
17                  Call objHT.CloseTransactionBengkel
18                  Call objHT.CloseTransactionHO
19                  AdaBegin = False
20                  MsgBox "Data Berhasil diproses!", vbInformation
21                  Unload Me
                    
22              Case NulSID("Upgrade Customer")
23                  KodeReQ = "": Set rso = Nothing: Set rso = New ADODB.Recordset: rsD.MoveFirst
                    
24                  sSQL = "UPDATE FUASURANSI SET NoPolish = '" & NulSID(txt(13).Text) & "', KodeAsuransi = " & CInt(DcAsuransi.Columns(2).value) & ",MasaberlakuAwal = '" & Format(tglpolish1.value, "mm/dd/yyyy") & "',MasaberlakuAkhir = '" & Format(tglpolish2.value, "mm/dd/yyyy") & "',Tenor = " & CInt(cboTenor.Text) & ", " & _
                               "JenisAsuransi1 = '" & s2d(Trim(isAs(cbojns(0)))) & "',JenisAsuransi2 = '" & s2d(Trim(isAs(cbojns(1)))) & "',JenisAsuransi3 = '" & s2d(Trim(isAs(cbojns(2)))) & "',JenisAsuransi4 = '" & s2d(Trim(isAs(cbojns(3)))) & "',JenisAsuransi5 = '" & s2d(Trim(isAs(cbojns(6)))) & "',JenisAsuransi6 = '" & s2d(Trim(isAs(cbojns(5)))) & "',JenisAsuransi7 = '" & s2d(Trim(isAs(cbojns(6)))) & "', " & _
                               "JenisAsuransi8 = '" & s2d(Trim(isAs(cbojns(7)))) & "',Terhubung = '" & IIf(Opt(0).value = True, 1, 0) & "',Deal = '" & IIf(Opt(2).value = True, 1, 0) & "',Alasan = '" & Trim(s2d(txt(11).Text)) & "' WHERE KODEREQ = '" & NulSID(rsD!KodeReQ) & "'"
                                
25                  AdaBegin = False: KetEro = ""
26                  Call objHT.OpenTransactionBengkel
27                  Call objHT.OpenTransactionHO
28                  AdaBegin = True
29                  Call objHT.InsertData_Bengkel(sSQL)
30                  sSQL = ""
31                  If Opt(2).value = True Then
32                      sSQL = "UPDATE OSS..Asuransi SET NoPolis = '" & NulSID(txt(13).Text) & "',MasaberlakuAwal = '" & Format(tglpolish1.value, "mm/dd/yyyy") & "',MasaberlakuAkhir = '" & Format(tglpolish2.value, "mm/dd/yyyy") & "', " & _
                               " banjir = '" & Ch(0).value & "',Gempa = '" & Ch(2).value & "',SRCC = '" & Ch(1).value & "',Teroris = '" & Ch(3).value & "'," & _
                               "TanggungP3 = " & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & ",CelakaDiri = " & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "," & _
                               "areaAsuransi = '" & Trim(DcAsuransi.Columns(1).value) & "',tenorAsuransi = '" & CInt(cboTenor.Text) & "',KodeAsuransi = '" & Trim(DcAsuransi.Columns(2).value) & "',JenisAsuransi = '" & isXs(cbojns(0).Text) & "',JenisAsuransi2 = '" & isXs(cbojns(1).Text) & "',JenisAsuransi3 = '" & isXs(cbojns(2).Text) & "',JenisAsuransi4 = '" & isXs(cbojns(3).Text) & "',JenisAsuransi5 = '" & isXs(cbojns(4).Text) & "',JenisAsuransi6 = '" & isXs(cbojns(5).Text) & "',JenisAsuransi7 = '" & isXs(cbojns(6).Text) & "',JenisAsuransi8 = '" & isXs(cbojns(7).Text) & "',CelakaDiriPass = " & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & " " & _
                               " WHERE NOSPK = '" & NulSID(rsD!KodeReQ) & "' and norangka = '" & NulSID(rsD!noRangka) & "'"
                               
33                        Call objHT.InsertData_Ho(sSQL)
                          
34                        sSQL = "UPDATE [OSS].[dbo].[UpgradeAsuransi] SET [KodeAsuransi] = '" & Trim(DcAsuransi.Columns(2).value) & "',[jenisAsuransi] = '" & isXs(cbojns(0).Text) & "'" & _
                                 " ,[jenisAsuransi2] = '" & isXs(cbojns(1).Text) & "',[jenisAsuransi3] = '" & isXs(cbojns(2).Text) & "',[jenisAsuransi4] = '" & isXs(cbojns(3).Text) & "',[jenisAsuransi5] = '" & isXs(cbojns(4).Text) & "',[jenisAsuransi6] = '" & isXs(cbojns(5).Text) & "',[jenisAsuransi7] = '" & isXs(cbojns(6).Text) & "',[jenisAsuransi8] = '" & isXs(cbojns(7).Text) & "',[Banjir] = '" & Ch(0).value & "',[Gempa]='" & Ch(2).value & "',[SRCC]='" & Ch(1).value & "',[Teroris] = '" & Ch(3).value & "',[TanggungP3] = " & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & ",[CelakaDiri] = " & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "" & _
                                 " ,[CelakaDiriPass] = " & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ",[tenorAsuransi] = '" & CInt(cboTenor.Text) & "'" & _
                                 " WHERE kodereq = '" & NulSID(rsD!KodeReQ) & "' and norangka = '" & NulSID(rsD!noRangka) & "'"
35                        Call objHT.InsertData_Ho(sSQL)
                          
36                  Else
                        
37                  End If
38                  Call objHT.CloseTransactionBengkel
39                  Call objHT.CloseTransactionHO
40                  AdaBegin = False
41                  MsgBox "Data Berhasil diproses!", vbInformation
42                  Unload Me
43              Case NulSID("Renew Customer")
44                  KodeReQ = "": Set rso = Nothing: Set rso = New ADODB.Recordset: rsD.MoveFirst
                     
45                  AdaBegin = False: KetEro = ""
46                  Call objHT.OpenTransactionBengkel
47                  Call objHT.OpenTransactionHO
48                   AdaBegin = True
49                        sSQL = "UPDATE FUASURANSI SET NoPolish = '" & NulSID(txt(13).Text) & "', KodeAsuransi = " & CInt(DcAsuransi.Columns(2).value) & ",MasaberlakuAwal = '" & Format(tglpolish1.value, "mm/dd/yyyy") & "',MasaberlakuAkhir = '" & Format(tglpolish2.value, "mm/dd/yyyy") & "',Tenor = " & CInt(cboTenor.Text) & ", " & _
                               "JenisAsuransi1 = '" & s2d(Trim(isAs(cbojns(0)))) & "',JenisAsuransi2 = '" & s2d(Trim(isAs(cbojns(1)))) & "',JenisAsuransi3 = '" & s2d(Trim(isAs(cbojns(2)))) & "',JenisAsuransi4 = '" & s2d(Trim(isAs(cbojns(3)))) & "',JenisAsuransi5 = '" & s2d(Trim(isAs(cbojns(6)))) & "',JenisAsuransi6 = '" & s2d(Trim(isAs(cbojns(5)))) & "',JenisAsuransi7 = '" & s2d(Trim(isAs(cbojns(6)))) & "', " & _
                               "JenisAsuransi8 = '" & s2d(Trim(isAs(cbojns(7)))) & "',Terhubung = '" & IIf(Opt(0).value = True, 1, 0) & "',Deal = '" & IIf(Opt(2).value = True, 1, 0) & "',Alasan = '" & Trim(s2d(txt(11).Text)) & "' WHERE KODEREQ = '" & NulSID(rsD!KodeReQ) & "'"
                                
50                  Call objHT.InsertData_Bengkel(sSQL)
51                  sSQL = ""
52                  If Opt(2).value = True Then
53                      sSQL = "UPDATE OSS..Asuransi SET NoPolis = '" & NulSID(txt(13).Text) & "',MasaberlakuAwal = '" & Format(tglpolish1.value, "mm/dd/yyyy") & "',MasaberlakuAkhir = '" & Format(tglpolish2.value, "mm/dd/yyyy") & "', " & _
                               " banjir = '" & Ch(0).value & "',Gempa = '" & Ch(2).value & "',SRCC = '" & Ch(1).value & "',Teroris = '" & Ch(3).value & "'," & _
                               "TanggungP3 = " & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & ",CelakaDiri = " & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "," & _
                               "areaAsuransi = '" & Trim(DcAsuransi.Columns(1).value) & "',tenorAsuransi = '" & CInt(cboTenor.Text) & "',KodeAsuransi = '" & Trim(DcAsuransi.Columns(2).value) & "',JenisAsuransi = '" & isXs(cbojns(0).Text) & "',JenisAsuransi2 = '" & isXs(cbojns(1).Text) & "',JenisAsuransi3 = '" & isXs(cbojns(2).Text) & "',JenisAsuransi4 = '" & isXs(cbojns(3).Text) & "',JenisAsuransi5 = '" & isXs(cbojns(4).Text) & "',JenisAsuransi6 = '" & isXs(cbojns(5).Text) & "',JenisAsuransi7 = '" & isXs(cbojns(6).Text) & "',JenisAsuransi8 = '" & isXs(cbojns(7).Text) & "',CelakaDiriPass = " & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & " " & _
                               " WHERE NOSPK = '" & NulSID(rsD!KodeReQ) & "' and norangka = '" & NulSID(rsD!noRangka) & "'"
                               
54                        Call objHT.InsertData_Ho(sSQL)
55                  Else
                        
56                  End If
57                  Call objHT.CloseTransactionBengkel
58                  Call objHT.CloseTransactionHO
59                  AdaBegin = False
60                  MsgBox "Data Berhasil diproses!", vbInformation
61                  Unload Me
62          End Select
63      Else
64            MsgBox "Data Yang Anda Inputkan Belum Lengkap !" & vbCrLf & isFalse, vbInformation
                 
65      End If
66      Exit Sub
Erro:
67      If AdaBegin Then
68        Call objHT.CloseTransactionBengkel_err
69        Call objHT.CloseTransactionHO_err
70      End If
       
71        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSave_Click of Form frmRetensiAsuransiCallHasil"

End Sub

Private Function isAs(vux As String) As String
    isAs = ""
    isAs = Replace(vux, "--pilih--", "")
End Function

Private Function isXs(vue As String) As String
    isXs = ""
    Select Case UCase(vue)
        Case "--PILIH--"
            isXs = "0"
        Case "ALL RISK"
            isXs = "2"
        Case "TOTAL LOST ONLY"
            isXs = "1"
    End Select
End Function

Private Function isValid() As Boolean
1       isFalse = ""
2       isValid = True
3       If Opt(1).value = True Then         'Tidak Terhubung
4           If txt(11).Text = "" Or Len(txt(11).Text) = 0 Then
5               isValid = False
6               isFalse = "[Alasan]"
7               Exit Function
8           End If
9       ElseIf Opt(0).value = True Then     'Terhubung
10          If Opt(3).value = True Then     'Tidak Deal
11              If txt(11).Text = "" Or Len(txt(11).Text) = 0 Then
12                  isValid = False
13                  isFalse = "[Alasan]"
14                  Exit Function
15              End If
16          ElseIf Opt(2).value = True Then 'Deal
17              If cboTenor.Text = "--pilih--" Then
18                  isValid = False
19                  isFalse = "[Tenor]"
20                  Exit Function
21              End If
22              If cboTenor.Text <> "--pilih--" Then
23                  For C = 0 To (CInt(cboTenor.Text) - 1)
24                      If cbojns(C).Text = "--pilih--" Then
25                          isValid = False
26                          isFalse = "[Jenis Asuransi Tahun " & C + 1 & "]"
27                          Exit Function
28                      End If
29                  Next C
30                  For C = (CInt(cboTenor.Text)) To 7
31                      If cbojns(C).Text <> "--pilih--" Then
32                          isValid = False
33                          isFalse = "[Jenis Asuransi Tahun " & C + 1 & " melebihi tenor]"
34                          Exit Function
35                      End If
36                  Next C
37              End If
      '          If cboPos(0).Text = "--pilih--" Then
      '              isValid = False
      '              isFalse = "[Pihak ke-3]"
      '              Exit Function
      '          End If
      '          If cboPos(1).Text = "--pilih--" Then
      '              isValid = False
      '              isFalse = "[Kecelakaan diri Driver]"
      '              Exit Function
      '          End If
      '          If cboPos(2).Text = "--pilih--" Then
      '              isValid = False
      '              isFalse = "[Kecelakaan diri Penumpang]"
      '              Exit Function
      '          End If
38              If DcAsuransi.Text = "" Then
39                  isValid = False
40                  isFalse = "[Asuransi]"
41              End If
42          Else
43              If Opt(2).value = False And Opt(3).value = False Then
44                  isValid = False
45                  isFalse = "[Deal or Tidak]"
46              End If
47          End If
48      End If
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdValidasi_Click()
    If MsgBox("Apakah yakin untuk validasi data?" & vbCrLf & "Data yang tervalidasi tidak dapat diubah.", vbYesNo) = vbYes Then
        sSQL = "UPDATE [OSS].[dbo].[UpgradeAsuransi] SET valid = 1" & _
                           " WHERE kodereq = '" & NulSID(rsD!KodeReQ) & "' and norangka = '" & NulSID(rsD!noRangka) & "'"
                    Call objO.InsertDataHO(sSQL)
        MsgBox "Data Berhasil divalidasi!", vbInformation
              Unload Me
    End If
End Sub

Private Sub DcAsuransi_Click()
    'MsgBox DcAsuransi.Columns(0).value & " | " & DcAsuransi.Columns(1).value & " | " & DcAsuransi.Columns(2).value
End Sub

Private Sub Form_Load()
    Me.Caption = TitleBP
   ' Move 3240, objA.MyTop + 385
    Me.Top = 0
    Move 0, 0
    Me.Width = Layar.Width - 600
    Me.Height = Layar.Height - 800
    Call clearAll
    If idAsuransi > 0 Then
        'DcAsuransi.Text = objO.FetchDataHO_String("SELECT NamaAsuransi + ' ' + WilayahAsuransi as NamaAsuransi, WilayahAsuransi, KodeAsuransi FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 AND KODEASURANSI = '" & idAsuransi & "'")
        DcAsuransi.value = objO.FetchDataHO_String("SELECT NamaAsuransi + ' ' + WilayahAsuransi as NamaAsuransi, WilayahAsuransi, KodeAsuransi FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 AND KODEASURANSI = '" & idAsuransi & "'")
    End If
    FrTenor.Visible = False
    FrJenis.Visible = False
    FrThird.Visible = False
    Opt(2).Enabled = False
    Opt(3).Enabled = False
    If rsD.RecordCount > 0 Then
        rsD.MoveFirst
        Call BindToCust
    End If
    
    isDataValid = False
    
    If NulSID(isAkt) = NulSID("upgrade customer") Then
        cmdValidasi.Visible = True
    Else
        cmdValidasi.Visible = False
    End If
    
    ''sSQL = "select a.*,b.terhubung,b.deal,b.alasan,b.aktifitas from [" & Log.LogServ & "].bengkel" & Right(gDb, 2) & ".dbo.FUAsuransi b with (nolock) left join oss..asuransi a on a.nospk = b.kodereq where b.kodereq = '" & NulSID(rsD!KodeReQ) & "' and b.norangka = '" & NulSID(rsD!noRangka) & "'"
    sSQL = "select b.*,a.jenisasuransi,a.jenisasuransi2,a.jenisasuransi3,a.jenisasuransi4,a.jenisasuransi5,a.jenisasuransi6,a.jenisasuransi7,a.jenisasuransi8,a.tenorasuransi,a.nopolis as nopolisasuransi,a.MASABERLAKUAWAL as MASABERLAKUAWALasuransi ,a.MASABERLAKUakhir as MASABERLAKUakhirasuransi from [" & Log.LogServ & "].bengkel" & Right(gDb, 2) & ".dbo.FUAsuransi b with (nolock) left join oss..asuransi a on a.nospk = b.kodereq where b.kodereq = '" & NulSID(rsD!KodeReQ) & "' and b.norangka = '" & NulSID(rsD!noRangka) & "'"
    Set rsP = objO.FetchDataHO(sSQL)
    If Not rsP.EOF Then
        rsP.MoveFirst
        If NulSID(rsP!nopolisasuransi) <> "" Then
            isDataValid = True
        Else
            sSQL = "SELECT VALID FROM OSS..UPGRADEASURANSI WITH (NOLOCK) WHERE KODEREQ = '" & NulSID(rsD!KodeReQ) & "' AND NORANGKA = '" & NulSID(rsD!noRangka) & "'"
            If objO.FetchDataHO_String(sSQL) = "" Then
                isDataValid = False
            Else
            ''If objO.FetchDataHO_String(sSQL) = False Or objO.FetchDataHO_String(sSQL) = "" Then
                If objO.FetchDataHO_String(sSQL) = False Then
                    isDataValid = False
                Else
                    isDataValid = True
                End If
            End If
        End If
        If NulSID(rsP!terhubung) = 1 Then
            Opt(0).value = True
        Else
            Opt(1).value = True
        End If
        If NulSID(rsP!deal) = 1 Then
            Opt(2).value = True
        Else
            Opt(3).value = True
        End If
        
        txt(11).Text = NulSID(rsP!alasan)
        
        DcAsuransi.Text = objO.FetchDataHO_String("SELECT NamaAsuransi + ' ' + WilayahAsuransi as NamaAsuransi, WilayahAsuransi, KodeAsuransi FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 AND KODEASURANSI = '" & NulSID(rsP!KodeAsuransi) & "'")
        'DcAsuransi.value = DcAsuransi.Text
        cboTenor.Text = IIf(NulSID(rsP!tenorasuransi) = "", "--pilih--", NulSID(rsP!tenorasuransi))
        cbojns(0).Text = IIf(NulSID(rsP!jenisasuransi) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi) = "1", "Total Lost Only", "All Risk"))
        cbojns(1).Text = IIf(NulSID(rsP!jenisasuransi2) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi2) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi2) = "", "--pilih--", NulSID(rsP!jenisasuransi2))
        cbojns(2).Text = IIf(NulSID(rsP!jenisasuransi3) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi3) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi3) = "", "--pilih--", NulSID(rsP!jenisasuransi3))
        cbojns(3).Text = IIf(NulSID(rsP!jenisasuransi4) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi4) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi4) = "", "--pilih--", NulSID(rsP!jenisasuransi4))
        cbojns(4).Text = IIf(NulSID(rsP!jenisasuransi5) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi5) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi5) = "", "--pilih--", NulSID(rsP!jenisasuransi5))
        cbojns(5).Text = IIf(NulSID(rsP!jenisasuransi6) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi6) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi6) = "", "--pilih--", NulSID(rsP!jenisasuransi6))
        cbojns(6).Text = IIf(NulSID(rsP!jenisasuransi7) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi7) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi7) = "", "--pilih--", NulSID(rsP!jenisasuransi7))
        cbojns(7).Text = IIf(NulSID(rsP!jenisasuransi8) = "0", "--pilih--", IIf(NulSID(rsP!jenisasuransi8) = "1", "Total Lost Only", "All Risk")) 'IIf(NulSID(rsP!jenisasuransi8) = "", "--pilih--", NulSID(rsP!jenisasuransi8))
        
        Ch(0).value = IIf(NulSID(rsP!banjir) = "", 0, NulSID(rsP!banjir))
        Ch(1).value = IIf(NulSID(rsP!srcc) = "", 0, NulSID(rsP!srcc)) 'NulSID(rsP!srcc)
        Ch(2).value = IIf(NulSID(rsP!gempa) = "", 0, NulSID(rsP!gempa)) 'NulSID(rsP!gempa)
        Ch(3).value = IIf(NulSID(rsP!teroris) = "", 0, NulSID(rsP!teroris)) 'NulSID(rsP!teroris)
        
        cboPos(0).Text = IIf(Nul(rsP!tanggungp3) > 0, Format(Nul(rsP!tanggungp3), "###,###,###,###"), "--pilih--")
        cboPos(1).Text = IIf(Nul(rsP!celakadiri) > 0, Format(Nul(rsP!celakadiri), "###,###,###,###"), "--pilih--")
        cboPos(2).Text = IIf(Nul(rsP!celakadiripass) > 0, Format(Nul(rsP!celakadiripass), "###,###,###,###"), "--pilih--")
        
        txt(13).Text = NulSID(rsP!nopolisasuransi)
        If NulSID(rsP!aktifitas) = "UPGRADE" Then
            txt(13).BackColor = &HC0E0FF
            txt(13).Enabled = False
        Else
            txt(13).BackColor = &H80000005
            txt(13).Enabled = True
        End If
        tglpolish1.value = Format(rsP!MASABERLAKUAWALasuransi, "dd/mm/yyyy")
        tglpolish2.value = Format(rsP!MASABERLAKUAkhirasuransi, "dd/mm/yyyy")
        
        If isDataValid = True Then
            cmdSave.Enabled = False
            cmdValidasi.Enabled = False
            MsgBox "Data sudah divalidasi.", vbInformation
        End If
    End If
    
End Sub

Private Sub BindToCust()
    txt(0).Text = Trim(Nul(rsD("nama_pemesan")))
    txt(1).Text = Trim(Nul(rsD("Alamat")))
    'txt(2).Text = Trim(Nul(rsd("telepon")))
    txt(3).Text = Trim(Nul(rsD("nohp1")))
    txt(4).Text = Trim(Nul(rsD("nohp2")))
    txt(5).Text = Trim(Nul(rsD("norangka")))
    txt(6).Text = Trim(Nul(rsD("nomesin")))
    txt(7).Text = Trim(Nul(rsD("namakategory")))
    txt(12).Text = Trim(Nul(rsD("warna")))
'    If Trim(Nul(rsd("nopolisi"))) <> "" Then
'        Call pecahNoPol(Trim(Nul(rsd("nopolisi"))), nDpn, nTgh, nBlk)
'        txt(8).Text = nDpn
'        txt(9).Text = nTgh
'        txt(10).Text = nBlk
'    Else
'        txt(8).Text = ""
'        txt(9).Text = ""
'        txt(10).Text = ""
'    End If
End Sub

Private Sub clearAll()
  For C = 0 To 9
      txt(C).Text = ""
  Next C
  For C = 0 To 3
      Opt(C).value = False
      Ch(C).value = 0
  Next C
  For C = 0 To 7
      cbojns(C).Clear
      cbojns(C).AddItem "--pilih--"            '(1)
      cbojns(C).AddItem "All Risk"            '(1)
      cbojns(C).AddItem "Total Lost Only"     '(2)
      cbojns(C).Text = cbojns(C).LIST(0)
  Next C
  With cboTenor
      .Clear
      .AddItem "--pilih--"
      .AddItem "1"
      .AddItem "2"
      .AddItem "3"
      .AddItem "4"
      .AddItem "5"
      .AddItem "6"
      .AddItem "7"
      .AddItem "8"
      .Text = .LIST(0)
  End With
    
    
  sSQL = "": sSQL = "select DISTINCT pertanggungan from simulasi_kredit..M_ThirdPartyLiability order by pertanggungan"
  Set Tri = Nothing: Set Tri = New ADODB.Recordset
  Set Tri = objO.FetchDataHO(sSQL)
  If Tri.RecordCount > 0 Then
      Tri.MoveFirst
      With cboPos(0)
          .Clear
          .AddItem "--pilih--"
          Do Until Tri.EOF
              .AddItem Format(Nul(Tri("pertanggungan")), "###,###,###,###")
              Tri.MoveNext
          Loop
          .Text = .LIST(0)
      End With
  End If
    
  sSQL = "": sSQL = "select DISTINCT pertanggungan from simulasi_kredit..M_PersonalAccident order by pertanggungan"
  Set Tri = Nothing: Set Tri = New ADODB.Recordset
  Set Tri = objO.FetchDataHO(sSQL)
  If Tri.RecordCount > 0 Then
      Tri.MoveFirst
      With cboPos(1)
          .Clear
          .AddItem "--pilih--"
          Do Until Tri.EOF
              .AddItem Format(Nul(Tri("pertanggungan")), "###,###,###,###")
              Tri.MoveNext
          Loop
          .Text = .LIST(0)
      End With
      Tri.MoveFirst
      With cboPos(2)
          .Clear
          .AddItem "--pilih--"
          Do Until Tri.EOF
              .AddItem Format(Nul(Tri("pertanggungan")), "###,###,###,###")
              Tri.MoveNext
          Loop
          .Text = .LIST(0)
      End With
  End If
    
  AdoAs.ConnectionString = SQLConBengkel99
  AdoAs.RecordSource = "SELECT NamaAsuransi + ' ' + WilayahAsuransi AS NamaAsuransi, WilayahAsuransi, KodeAsuransi FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 ORDER BY NamaAsuransi,WilayahAsuransi "

  DcAsuransi.DataSourceList = AdoAs
  'DcAsuransi.DataFieldToDisplay = "NamaAsuransi"
  DcAsuransi.DataFieldList = "NamaAsuransi" '"kode"
  DcAsuransi.Columns(0).Width = 4200
  DcAsuransi.Columns(1).Width = 3000
  DcAsuransi.Columns(2).Width = 1000
  DcAsuransi.Columns(0).Caption = "Nama Asuransi"
  DcAsuransi.Columns(1).Caption = "Wilayah Asuransi"
  DcAsuransi.Columns(2).Caption = "Kode"

End Sub



Private Sub Opt_Click(Index As Integer)
    Select Case Index
        Case 0
            If Opt(Index).value = True Then
                Opt(2).Enabled = True
                Opt(3).Enabled = True
                If Opt(2).value = True Then
                    FrTenor.Visible = True
                    FrJenis.Visible = True
                    FrThird.Visible = True
                End If
                txt(11).Text = ""
            End If
        Case 1
            If Opt(Index).value = True Then
                Opt(2).Enabled = False
                Opt(3).Enabled = False
                FrTenor.Visible = False
                FrJenis.Visible = False
                FrThird.Visible = False
            End If
        Case 2
            If Opt(Index).value = True Then
                FrTenor.Visible = True
                FrJenis.Visible = True
                FrThird.Visible = True
                frDetailPolish.Visible = True
                txt(11).Text = ""
            End If
        Case 3
            If Opt(Index).value = True Then
                FrTenor.Visible = False
                FrJenis.Visible = False
                FrThird.Visible = False
                frDetailPolish.Visible = False
               ' txt(11).SetFocus
            End If
    End Select
End Sub

Private Sub tglpolish1_LostFocus()
    If tglpolish1.value <> "" And cboTenor.Text <> "" Then
        tglpolish2.value = DateAdd("yyyy", cboTenor.Text, tglpolish1)
    End If
End Sub
