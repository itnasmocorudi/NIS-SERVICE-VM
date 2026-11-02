VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmSIDMasterPelanggan 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14250
   Icon            =   "frmSIDMasterPelanggan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSIDMasterPelanggan.frx":29C12
   ScaleHeight     =   10500
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   480
      TabIndex        =   8
      Top             =   0
      Width           =   13695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H80000008&
         Height          =   8850
         Left            =   120
         ScaleHeight     =   8820
         ScaleWidth      =   11370
         TabIndex        =   10
         Top             =   375
         Width           =   11400
         Begin VB.PictureBox PicMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   20000
            Left            =   15
            ScaleHeight     =   19995
            ScaleWidth      =   11295
            TabIndex        =   11
            Top             =   105
            Width           =   11295
            Begin VB.Frame FrDataPemakai 
               BackColor       =   &H80000016&
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   4440
               Left            =   20
               TabIndex        =   121
               Top             =   12120
               Width           =   11190
               Begin VB.Frame FrAlmS 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000016&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1335
                  Left            =   4800
                  TabIndex        =   179
                  Top             =   1515
                  Visible         =   0   'False
                  Width           =   2730
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
                     Height          =   320
                     IMEMode         =   3  'DISABLE
                     Index           =   77
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   183
                     Top             =   1005
                     Width           =   2715
                  End
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
                     Height          =   320
                     IMEMode         =   3  'DISABLE
                     Index           =   76
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   182
                     Top             =   675
                     Width           =   2715
                  End
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
                     Height          =   320
                     IMEMode         =   3  'DISABLE
                     Index           =   75
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   181
                     Top             =   345
                     Width           =   2715
                  End
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
                     Height          =   320
                     IMEMode         =   3  'DISABLE
                     Index           =   0
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   180
                     Top             =   15
                     Width           =   2715
                  End
               End
               Begin VB.ComboBox cboOlahS 
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
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   3255
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
                  Index           =   74
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   63
                  Top             =   3930
                  Width           =   2190
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
                  Index           =   73
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   62
                  Top             =   3600
                  Width           =   2190
               End
               Begin VB.CheckBox ChSurat 
                  BackColor       =   &H80000016&
                  Caption         =   "Data Pemakai sama dengan Data Pelanggan"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00202020&
                  Height          =   255
                  Left            =   75
                  TabIndex        =   49
                  Top             =   240
                  Width           =   5655
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
                  Index           =   49
                  Left            =   7875
                  MaxLength       =   25
                  TabIndex        =   51
                  Top             =   555
                  Width           =   3075
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
                  Index           =   48
                  Left            =   1755
                  MaxLength       =   50
                  TabIndex        =   50
                  Top             =   555
                  Width           =   3510
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
                  Index           =   52
                  Left            =   7875
                  MaxLength       =   150
                  TabIndex        =   55
                  Top             =   885
                  Width           =   3075
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
                  Index           =   50
                  Left            =   2700
                  MaxLength       =   30
                  TabIndex        =   53
                  Top             =   885
                  Width           =   2560
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
                  Index           =   53
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   59
                  Top             =   2565
                  Width           =   2190
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
                  Height          =   810
                  IMEMode         =   3  'DISABLE
                  Index           =   54
                  Left            =   4800
                  MaxLength       =   150
                  MultiLine       =   -1  'True
                  TabIndex        =   68
                  Top             =   2850
                  Width           =   2985
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
                  Index           =   51
                  Left            =   5280
                  MaxLength       =   30
                  TabIndex        =   54
                  Top             =   885
                  Width           =   2580
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
                  Index           =   55
                  Left            =   4800
                  MaxLength       =   3
                  TabIndex        =   69
                  Top             =   3675
                  Width           =   570
               End
               Begin VB.ComboBox cboTitleS 
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
                  Left            =   1755
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   885
                  Width           =   930
               End
               Begin VB.ComboBox cboAgamaS 
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
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   57
                  Top             =   1905
                  Width           =   2205
               End
               Begin VB.ComboBox cboPekS 
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
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   58
                  Top             =   2235
                  Width           =   2205
               End
               Begin VB.ComboBox cboHobiS 
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
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   2910
                  Width           =   2205
               End
               Begin VB.ComboBox cboProvS 
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   64
                  Top             =   1575
                  Width           =   2700
               End
               Begin VB.ComboBox cboKotaS 
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   65
                  Top             =   1890
                  Width           =   2700
               End
               Begin VB.ComboBox cboKecS 
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   66
                  Top             =   2205
                  Width           =   2700
               End
               Begin VB.ComboBox cboKelS 
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   2520
                  Width           =   2700
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
                  Index           =   56
                  Left            =   5580
                  MaxLength       =   3
                  TabIndex        =   70
                  Top             =   3675
                  Width           =   570
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
                  Index           =   57
                  Left            =   4800
                  MaxLength       =   10
                  TabIndex        =   71
                  Top             =   4005
                  Width           =   1350
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
                  Index           =   58
                  Left            =   8760
                  MaxLength       =   4
                  TabIndex        =   72
                  Top             =   1575
                  Width           =   645
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
                  Index           =   61
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   75
                  Top             =   2235
                  Width           =   2085
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
                  Index           =   59
                  Left            =   9420
                  MaxLength       =   10
                  TabIndex        =   73
                  Top             =   1575
                  Width           =   1425
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
                  Index           =   60
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   74
                  Top             =   1905
                  Width           =   2085
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
                  Index           =   62
                  Left            =   8760
                  MaxLength       =   100
                  TabIndex        =   76
                  Top             =   2565
                  Width           =   2085
               End
               Begin TDBDate6Ctl.TDBDate DtLahirS 
                  Height          =   315
                  Left            =   1230
                  TabIndex        =   56
                  Top             =   1575
                  Width           =   1530
                  _Version        =   65536
                  _ExtentX        =   2699
                  _ExtentY        =   556
                  Calendar        =   "frmSIDMasterPelanggan.frx":2BEE2
                  Caption         =   "frmSIDMasterPelanggan.frx":2C00E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmSIDMasterPelanggan.frx":2C06A
                  Keys            =   "frmSIDMasterPelanggan.frx":2C088
                  Spin            =   "frmSIDMasterPelanggan.frx":2C0E6
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   12640511
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
                  ForeColor       =   0
                  Format          =   "dd/mm/yyyy"
                  HighlightText   =   0
                  IMEMode         =   3
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxDate         =   2958465
                  MinDate         =   2
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
                  ValueVT         =   6815745
                  Value           =   39183
                  CenturyMode     =   0
               End
               Begin ucXPButton.XPButton cmdBackS 
                  Height          =   345
                  Left            =   7530
                  TabIndex        =   184
                  Top             =   1530
                  Visible         =   0   'False
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  BTYPE           =   3
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
                  BCOL            =   14215660
                  BCOLO           =   14215660
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "frmSIDMasterPelanggan.frx":2C10E
                  PICN            =   "frmSIDMasterPelanggan.frx":2C12A
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   178
                  Top             =   3990
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preferensi"
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
                  Left            =   75
                  TabIndex        =   177
                  Top             =   3990
                  Width           =   1020
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   176
                  Top             =   3660
                  Width           =   60
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   175
                  Top             =   3315
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suku"
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
                  Left            =   75
                  TabIndex        =   174
                  Top             =   3660
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Olahraga"
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
                  Left            =   75
                  TabIndex        =   173
                  Top             =   3315
                  Width           =   885
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
                  Index           =   3
                  Left            =   1635
                  TabIndex        =   163
                  Top             =   570
                  Width           =   60
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No. KTP"
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
                  Left            =   75
                  TabIndex        =   162
                  Top             =   570
                  Width           =   735
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " Data Pemakai"
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
                  Left            =   0
                  TabIndex        =   161
                  Top             =   0
                  Width           =   1410
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
                  Index           =   5
                  Left            =   7750
                  TabIndex        =   160
                  Top             =   570
                  Width           =   60
               End
               Begin VB.Label Label6 
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
                  Left            =   1635
                  TabIndex        =   159
                  Top             =   900
                  Width           =   60
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama Customer"
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
                  Left            =   75
                  TabIndex        =   158
                  Top             =   900
                  Width           =   1545
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "NPWP"
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
                  Left            =   7020
                  TabIndex        =   157
                  Top             =   570
                  Width           =   555
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Lahir"
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
                  Index           =   56
                  Left            =   75
                  TabIndex        =   156
                  Top             =   1605
                  Width           =   855
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
                  Index           =   42
                  Left            =   1125
                  TabIndex        =   155
                  Top             =   1590
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Agama"
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
                  Index           =   55
                  Left            =   75
                  TabIndex        =   154
                  Top             =   1905
                  Width           =   675
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
                  Index           =   41
                  Left            =   1125
                  TabIndex        =   153
                  Top             =   1905
                  Width           =   60
               End
               Begin VB.Line Line7 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  X1              =   1635
                  X2              =   10950
                  Y1              =   120
                  Y2              =   120
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   152
                  Top             =   2910
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hobi"
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
                  Left            =   75
                  TabIndex        =   151
                  Top             =   2925
                  Width           =   435
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
                  Left            =   1125
                  TabIndex        =   150
                  Top             =   2250
                  Width           =   60
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pekerjaan"
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
                  Left            =   75
                  TabIndex        =   149
                  Top             =   2250
                  Width           =   1005
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Data Diri"
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
                  Height          =   225
                  Index           =   54
                  Left            =   0
                  TabIndex        =   148
                  Top             =   1275
                  Width           =   3480
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
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
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   53
                  Left            =   3495
                  TabIndex        =   147
                  Top             =   1275
                  Width           =   4365
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
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
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   52
                  Left            =   7875
                  TabIndex        =   146
                  Top             =   1275
                  Width           =   3075
               End
               Begin VB.Shape Shape9 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   0
                  Top             =   1500
                  Width           =   3480
               End
               Begin VB.Shape Shape8 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   3495
                  Top             =   1500
                  Width           =   4365
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
                  Index           =   40
                  Left            =   4700
                  TabIndex        =   145
                  Top             =   1605
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Provinsi"
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
                  Index           =   51
                  Left            =   3550
                  TabIndex        =   144
                  Top             =   1620
                  Width           =   795
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
                  Index           =   39
                  Left            =   4700
                  TabIndex        =   143
                  Top             =   1920
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kab/Kota"
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
                  Index           =   50
                  Left            =   3550
                  TabIndex        =   142
                  Top             =   1935
                  Width           =   915
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
                  Index           =   38
                  Left            =   4700
                  TabIndex        =   141
                  Top             =   2235
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kecamatan"
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
                  Index           =   49
                  Left            =   3550
                  TabIndex        =   140
                  Top             =   2235
                  Width           =   1080
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
                  Index           =   37
                  Left            =   4700
                  TabIndex        =   139
                  Top             =   2565
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kelurahan"
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
                  Index           =   48
                  Left            =   3550
                  TabIndex        =   138
                  Top             =   2565
                  Width           =   990
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
                  Index           =   36
                  Left            =   4700
                  TabIndex        =   137
                  Top             =   2835
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jalan /"
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
                  Index           =   47
                  Left            =   3550
                  TabIndex        =   136
                  Top             =   2835
                  Width           =   690
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kompleks;"
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
                  Index           =   46
                  Left            =   3550
                  TabIndex        =   135
                  Top             =   3075
                  Width           =   1005
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No"
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
                  Index           =   45
                  Left            =   3550
                  TabIndex        =   134
                  Top             =   3285
                  Width           =   255
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":           /"
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
                  Index           =   35
                  Left            =   4700
                  TabIndex        =   133
                  Top             =   3690
                  Width           =   840
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RT/RW"
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
                  Index           =   44
                  Left            =   3550
                  TabIndex        =   132
                  Top             =   3690
                  Width           =   660
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
                  Index           =   34
                  Left            =   4695
                  TabIndex        =   131
                  Top             =   4035
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kode Pos"
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
                  Index           =   43
                  Left            =   3555
                  TabIndex        =   130
                  Top             =   4035
                  Width           =   885
               End
               Begin VB.Shape Shape7 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   7875
                  Top             =   1500
                  Width           =   3075
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
                  Index           =   33
                  Left            =   8650
                  TabIndex        =   129
                  Top             =   1620
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "PSTN"
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
                  Index           =   42
                  Left            =   7940
                  TabIndex        =   128
                  Top             =   1620
                  Width           =   495
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
                  Index           =   32
                  Left            =   8650
                  TabIndex        =   127
                  Top             =   1935
                  Width           =   60
               End
               Begin VB.Label Label11 
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
                  Index           =   41
                  Left            =   7940
                  TabIndex        =   126
                  Top             =   1935
                  Width           =   435
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
                  Index           =   31
                  Left            =   8650
                  TabIndex        =   125
                  Top             =   2265
                  Width           =   60
               End
               Begin VB.Label Label11 
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
                  Index           =   40
                  Left            =   7940
                  TabIndex        =   124
                  Top             =   2265
                  Width           =   435
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
                  Index           =   30
                  Left            =   8650
                  TabIndex        =   123
                  Top             =   2595
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-mail"
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
                  Index           =   33
                  Left            =   7940
                  TabIndex        =   122
                  Top             =   2595
                  Width           =   630
               End
            End
            Begin VB.Frame FrDtPlgan 
               BackColor       =   &H80000016&
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   5955
               Left            =   30
               TabIndex        =   12
               Top             =   -45
               Width           =   11190
               Begin MSMask.MaskEdBox txtNpwp 
                  Height          =   330
                  Left            =   8085
                  TabIndex        =   213
                  Top             =   810
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   582
                  _Version        =   393216
                  MaxLength       =   20
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##.###.###.#-###.###"
                  PromptChar      =   "_"
               End
               Begin VB.ComboBox cboKewarganegaraan 
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
                  ItemData        =   "frmSIDMasterPelanggan.frx":2C335
                  Left            =   1965
                  List            =   "frmSIDMasterPelanggan.frx":2C337
                  Style           =   2  'Dropdown List
                  TabIndex        =   205
                  Top             =   495
                  Width           =   1125
               End
               Begin MSMask.MaskEdBox MaskEdBox9 
                  Height          =   270
                  Left            =   6180
                  TabIndex        =   204
                  Top             =   1845
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
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
                  Index           =   23
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   23
                  Top             =   2880
                  Width           =   2075
               End
               Begin MSMask.MaskEdBox MaskEdBox8 
                  Height          =   270
                  Left            =   3345
                  TabIndex        =   203
                  Top             =   3600
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
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
                  Index           =   22
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   25
                  Top             =   3570
                  Width           =   2075
               End
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
                  Index           =   21
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   24
                  Top             =   3225
                  Width           =   2075
               End
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
                  Index           =   20
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   21
                  Top             =   2145
                  Width           =   2075
               End
               Begin VB.Frame FrAlm 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000016&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1335
                  Left            =   4800
                  TabIndex        =   165
                  Top             =   2145
                  Visible         =   0   'False
                  Width           =   3045
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Index           =   66
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   35
                     Top             =   1005
                     Width           =   2950
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Index           =   65
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   34
                     Top             =   675
                     Width           =   2950
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Index           =   64
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   33
                     Top             =   345
                     Width           =   2950
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Index           =   63
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   32
                     Top             =   15
                     Width           =   2950
                  End
                  Begin MSMask.MaskEdBox MaskEdBox3 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   196
                     Top             =   30
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox MaskEdBox1 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   197
                     Top             =   360
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox MaskEdBox2 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   198
                     Top             =   705
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox MaskEdBox4 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   199
                     Top             =   1035
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
               End
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
                  Index           =   19
                  Left            =   8760
                  MaxLength       =   15
                  TabIndex        =   193
                  Top             =   4110
                  Width           =   2085
               End
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
                  Index           =   18
                  Left            =   8760
                  MaxLength       =   30
                  TabIndex        =   190
                  Top             =   3765
                  Width           =   2085
               End
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
                  Index           =   17
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   186
                  Top             =   3150
                  Width           =   2085
               End
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
                  Index           =   72
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   27
                  Top             =   4245
                  Width           =   2075
               End
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
                  Index           =   71
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   26
                  Top             =   3915
                  Width           =   2075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   795
                  IMEMode         =   3  'DISABLE
                  Index           =   8
                  Left            =   4800
                  MaxLength       =   150
                  MultiLine       =   -1  'True
                  TabIndex        =   40
                  Top             =   3480
                  Width           =   2985
               End
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
                  Index           =   9
                  Left            =   4800
                  MaxLength       =   3
                  TabIndex        =   41
                  Top             =   4290
                  Width           =   570
               End
               Begin VB.ComboBox cboKec 
                  BackColor       =   &H0080FFFF&
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   2190
                  Width           =   2700
               End
               Begin VB.ComboBox cboKel 
                  BackColor       =   &H0080FFFF&
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   2505
                  Width           =   2700
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Index           =   10
                  Left            =   5580
                  MaxLength       =   3
                  TabIndex        =   42
                  Top             =   4290
                  Width           =   570
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Index           =   11
                  Left            =   4800
                  MaxLength       =   6
                  TabIndex        =   43
                  Top             =   1830
                  Width           =   1350
               End
               Begin VB.ComboBox cboProv 
                  BackColor       =   &H0080FFFF&
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   1875
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   22
                  Top             =   2535
                  Width           =   2075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   8085
                  MaxLength       =   150
                  TabIndex        =   19
                  Top             =   1170
                  Width           =   3075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   5490
                  MaxLength       =   30
                  TabIndex        =   18
                  Top             =   1170
                  Width           =   2580
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   2910
                  MaxLength       =   30
                  TabIndex        =   17
                  Top             =   1170
                  Width           =   2560
               End
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
                  Index           =   3
                  Left            =   8085
                  MaxLength       =   20
                  TabIndex        =   15
                  Top             =   825
                  Width           =   3075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   1965
                  MaxLength       =   16
                  TabIndex        =   14
                  Top             =   825
                  Width           =   3510
               End
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
                  Index           =   16
                  Left            =   8760
                  MaxLength       =   100
                  TabIndex        =   48
                  Top             =   2820
                  Width           =   2085
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Index           =   14
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   46
                  Top             =   2160
                  Width           =   2085
               End
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
                  Left            =   9420
                  MaxLength       =   10
                  TabIndex        =   45
                  Top             =   1830
                  Width           =   1425
               End
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
                  Index           =   15
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   47
                  Top             =   2490
                  Width           =   2085
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Index           =   12
                  Left            =   8760
                  MaxLength       =   4
                  TabIndex        =   44
                  Top             =   1830
                  Width           =   645
               End
               Begin VB.ComboBox cboTitle 
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
                  Left            =   1965
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   1170
                  Width           =   930
               End
               Begin VB.ComboBox cboTipe 
                  BackColor       =   &H0080FFFF&
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
                  Left            =   1965
                  Style           =   2  'Dropdown List
                  TabIndex        =   13
                  Top             =   165
                  Width           =   2490
               End
               Begin TDBDate6Ctl.TDBDate DtLahir 
                  Height          =   315
                  Left            =   1245
                  TabIndex        =   20
                  Top             =   1815
                  Width           =   1530
                  _Version        =   65536
                  _ExtentX        =   2699
                  _ExtentY        =   556
                  Calendar        =   "frmSIDMasterPelanggan.frx":2C339
                  Caption         =   "frmSIDMasterPelanggan.frx":2C465
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmSIDMasterPelanggan.frx":2C4C1
                  Keys            =   "frmSIDMasterPelanggan.frx":2C4DF
                  Spin            =   "frmSIDMasterPelanggan.frx":2C53D
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
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
                  ForeColor       =   0
                  Format          =   "dd/mm/yyyy"
                  HighlightText   =   0
                  IMEMode         =   3
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxDate         =   2958465
                  MinDate         =   2
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
                  ValueVT         =   6815745
                  Value           =   39183
                  CenturyMode     =   0
               End
               Begin ucXPButton.XPButton cmdBack 
                  Height          =   345
                  Left            =   7455
                  TabIndex        =   166
                  Top             =   3735
                  Visible         =   0   'False
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  BTYPE           =   3
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
                  BCOL            =   14215660
                  BCOLO           =   14215660
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "frmSIDMasterPelanggan.frx":2C565
                  PICN            =   "frmSIDMasterPelanggan.frx":2C581
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSMask.MaskEdBox MaskEdBox5 
                  Height          =   270
                  Left            =   3360
                  TabIndex        =   200
                  Top             =   2160
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskEdBox6 
                  Height          =   270
                  Left            =   3360
                  TabIndex        =   201
                  Top             =   2565
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskEdBox7 
                  Height          =   270
                  Left            =   3345
                  TabIndex        =   202
                  Top             =   3255
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin VB.ComboBox cboAgama 
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
                  Height          =   315
                  Left            =   1245
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   2130
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboPek 
                  BackColor       =   &H0080FFFF&
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
                  Left            =   1245
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboHobi 
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
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   2250
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboOlah 
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
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   2235
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboKota 
                  BackColor       =   &H0080FFFF&
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
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   1875
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H80000016&
                  Height          =   540
                  Left            =   7245
                  TabIndex        =   209
                  Top             =   225
                  Width           =   3885
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
                     Index           =   1
                     Left            =   2040
                     MaxLength       =   12
                     TabIndex        =   210
                     Top             =   135
                     Width           =   1740
                  End
                  Begin VB.Label Label23 
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
                     Left            =   1935
                     TabIndex        =   212
                     Top             =   195
                     Width           =   60
                  End
                  Begin VB.Label Label24 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Kode Customer"
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
                     Left            =   120
                     TabIndex        =   211
                     Top             =   195
                     Width           =   1470
                  End
               End
               Begin VB.Label Label2 
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
                  Left            =   1860
                  TabIndex        =   208
                  Top             =   540
                  Width           =   60
               End
               Begin VB.Label Label6 
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
                  Left            =   1650
                  TabIndex        =   207
                  Top             =   900
                  Width           =   60
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kewarganegaraan"
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
                  Left            =   45
                  TabIndex        =   206
                  Top             =   555
                  Width           =   1770
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
                  Left            =   8640
                  TabIndex        =   195
                  Top             =   4170
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Telp"
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
                  Left            =   7935
                  TabIndex        =   194
                  Top             =   4140
                  Width           =   420
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
                  Left            =   8640
                  TabIndex        =   192
                  Top             =   3795
                  Width           =   60
               End
               Begin VB.Label Label11 
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
                  Index           =   5
                  Left            =   7935
                  TabIndex        =   191
                  Top             =   3795
                  Width           =   555
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Data PIC"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   4
                  Left            =   7860
                  TabIndex        =   189
                  Top             =   3495
                  Width           =   3075
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
                  Left            =   8640
                  TabIndex        =   188
                  Top             =   3180
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No Fax"
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
                  Left            =   7935
                  TabIndex        =   187
                  Top             =   3180
                  Width           =   675
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   172
                  Top             =   4290
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preferensi"
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
                  Left            =   60
                  TabIndex        =   171
                  Top             =   4305
                  Width           =   1020
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   170
                  Top             =   3975
                  Width           =   60
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   169
                  Top             =   3630
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suku"
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
                  Left            =   60
                  TabIndex        =   168
                  Top             =   3975
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Olahraga"
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
                  Left            =   60
                  TabIndex        =   167
                  Top             =   3630
                  Width           =   885
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-mail"
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
                  Index           =   28
                  Left            =   7935
                  TabIndex        =   120
                  Top             =   2850
                  Width           =   630
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
                  Index           =   22
                  Left            =   8655
                  TabIndex        =   119
                  Top             =   2850
                  Width           =   60
               End
               Begin VB.Label Label11 
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
                  Index           =   27
                  Left            =   7935
                  TabIndex        =   118
                  Top             =   2520
                  Width           =   435
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
                  Index           =   21
                  Left            =   8655
                  TabIndex        =   117
                  Top             =   2520
                  Width           =   60
               End
               Begin VB.Label Label11 
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
                  Index           =   26
                  Left            =   7935
                  TabIndex        =   116
                  Top             =   2190
                  Width           =   435
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
                  Index           =   20
                  Left            =   8655
                  TabIndex        =   115
                  Top             =   2190
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "PSTN"
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
                  Index           =   25
                  Left            =   7935
                  TabIndex        =   114
                  Top             =   1875
                  Width           =   495
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
                  Index           =   19
                  Left            =   8655
                  TabIndex        =   113
                  Top             =   1875
                  Width           =   60
               End
               Begin VB.Shape Shape4 
                  BackColor       =   &H80000016&
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   7875
                  Top             =   1785
                  Width           =   3075
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kode Pos"
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
                  Index           =   24
                  Left            =   3555
                  TabIndex        =   112
                  Top             =   1875
                  Width           =   885
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
                  Index           =   18
                  Left            =   4680
                  TabIndex        =   111
                  Top             =   1920
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RT/RW"
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
                  Left            =   3555
                  TabIndex        =   110
                  Top             =   4320
                  Width           =   660
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":           /"
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
                  Left            =   4695
                  TabIndex        =   109
                  Top             =   4320
                  Width           =   840
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No"
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
                  Left            =   3555
                  TabIndex        =   108
                  Top             =   3915
                  Width           =   255
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kompleks;"
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
                  Left            =   3555
                  TabIndex        =   107
                  Top             =   3705
                  Width           =   1005
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jalan /"
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
                  Left            =   3555
                  TabIndex        =   106
                  Top             =   3465
                  Width           =   690
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
                  Left            =   4695
                  TabIndex        =   105
                  Top             =   3465
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kelurahan"
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
                  Left            =   3555
                  TabIndex        =   104
                  Top             =   3180
                  Width           =   990
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
                  Left            =   4695
                  TabIndex        =   103
                  Top             =   3180
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kecamatan"
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
                  Left            =   3555
                  TabIndex        =   102
                  Top             =   2850
                  Width           =   1080
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
                  Left            =   4695
                  TabIndex        =   101
                  Top             =   2880
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kab/Kota"
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
                  Left            =   3555
                  TabIndex        =   100
                  Top             =   2550
                  Width           =   915
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
                  Left            =   4695
                  TabIndex        =   99
                  Top             =   2535
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Provinsi"
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
                  Left            =   3555
                  TabIndex        =   98
                  Top             =   2235
                  Width           =   795
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
                  Left            =   4695
                  TabIndex        =   97
                  Top             =   2220
                  Width           =   60
               End
               Begin VB.Shape Shape3 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   3495
                  Top             =   1785
                  Width           =   4365
               End
               Begin VB.Shape Shape12 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   0
                  Top             =   1785
                  Width           =   3480
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
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
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   12
                  Left            =   7875
                  TabIndex        =   96
                  Top             =   1515
                  Width           =   3075
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
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
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   11
                  Left            =   3495
                  TabIndex        =   95
                  Top             =   1515
                  Width           =   4365
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Data Diri"
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
                  Height          =   225
                  Index           =   0
                  Left            =   0
                  TabIndex        =   94
                  Top             =   1515
                  Width           =   3480
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pekerjaan"
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
                  Left            =   60
                  TabIndex        =   93
                  Top             =   2565
                  Width           =   1005
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
                  Index           =   4
                  Left            =   1125
                  TabIndex        =   92
                  Top             =   2565
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hobi"
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
                  Left            =   60
                  TabIndex        =   91
                  Top             =   3285
                  Width           =   435
               End
               Begin VB.Label Label12 
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
                  Left            =   1125
                  TabIndex        =   90
                  Top             =   3285
                  Width           =   60
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  X1              =   1635
                  X2              =   11190
                  Y1              =   120
                  Y2              =   120
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
                  Left            =   1125
                  TabIndex        =   89
                  Top             =   2205
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Agama"
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
                  Left            =   60
                  TabIndex        =   88
                  Top             =   2205
                  Width           =   675
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
                  Left            =   1125
                  TabIndex        =   87
                  Top             =   1845
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Lahir"
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
                  Left            =   60
                  TabIndex        =   86
                  Top             =   1860
                  Width           =   855
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "NPWP"
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
                  Left            =   6540
                  TabIndex        =   85
                  Top             =   870
                  Width           =   555
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama Customer"
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
                  Left            =   60
                  TabIndex        =   84
                  Top             =   1230
                  Width           =   1545
               End
               Begin VB.Label Label6 
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
                  Left            =   1860
                  TabIndex        =   83
                  Top             =   1230
                  Width           =   60
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
                  Index           =   0
                  Left            =   7965
                  TabIndex        =   82
                  Top             =   870
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipe Customer"
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
                  Left            =   420
                  TabIndex        =   81
                  Top             =   225
                  Width           =   1410
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " Data Pelanggan"
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
                  Left            =   0
                  TabIndex        =   80
                  Top             =   0
                  Width           =   1590
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Passport / KITAS"
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
                  Left            =   45
                  TabIndex        =   79
                  Top             =   885
                  Width           =   1680
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
                  Left            =   1860
                  TabIndex        =   78
                  Top             =   885
                  Width           =   60
               End
               Begin VB.Label Label12 
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
                  Left            =   1845
                  TabIndex        =   77
                  Top             =   225
                  Width           =   60
               End
            End
         End
      End
      Begin ucXPButton.XPButton cmdAdd 
         Height          =   495
         Left            =   300
         TabIndex        =   0
         Top             =   9030
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmSIDMasterPelanggan.frx":2C78C
         PICN            =   "frmSIDMasterPelanggan.frx":2C7A8
         PICH            =   "frmSIDMasterPelanggan.frx":2CBF4
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
         TabIndex        =   7
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
         MICON           =   "frmSIDMasterPelanggan.frx":2D0BC
         PICN            =   "frmSIDMasterPelanggan.frx":2D0D8
         PICH            =   "frmSIDMasterPelanggan.frx":2D5AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCreate 
         Height          =   495
         Left            =   1455
         TabIndex        =   2
         Top             =   9045
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&r"
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
         MICON           =   "frmSIDMasterPelanggan.frx":2DAF7
         PICN            =   "frmSIDMasterPelanggan.frx":2DB13
         PICH            =   "frmSIDMasterPelanggan.frx":2E048
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdOpen 
         Height          =   495
         Left            =   5025
         TabIndex        =   6
         Top             =   9375
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&o"
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
         MICON           =   "frmSIDMasterPelanggan.frx":2E63B
         PICN            =   "frmSIDMasterPelanggan.frx":2E657
         PICH            =   "frmSIDMasterPelanggan.frx":2EB18
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   1440
         TabIndex        =   3
         Top             =   9360
         Visible         =   0   'False
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
         MICON           =   "frmSIDMasterPelanggan.frx":2F06E
         PICN            =   "frmSIDMasterPelanggan.frx":2F08A
         PICH            =   "frmSIDMasterPelanggan.frx":2F59D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   9360
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&c"
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
         MICON           =   "frmSIDMasterPelanggan.frx":2FB3F
         PICN            =   "frmSIDMasterPelanggan.frx":2FB5B
         PICH            =   "frmSIDMasterPelanggan.frx":3009F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdDelete 
         Height          =   495
         Left            =   3840
         TabIndex        =   5
         Top             =   9360
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
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
         MICON           =   "frmSIDMasterPelanggan.frx":30696
         PICN            =   "frmSIDMasterPelanggan.frx":306B2
         PICH            =   "frmSIDMasterPelanggan.frx":30C3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   9360
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&e"
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
         MICON           =   "frmSIDMasterPelanggan.frx":31276
         PICN            =   "frmSIDMasterPelanggan.frx":31292
         PICH            =   "frmSIDMasterPelanggan.frx":3173D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   8880
         TabIndex        =   185
         Top             =   9360
         Width           =   1170
         _ExtentX        =   2064
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
         MICON           =   "frmSIDMasterPelanggan.frx":31C62
         PICN            =   "frmSIDMasterPelanggan.frx":31C7E
         PICH            =   "frmSIDMasterPelanggan.frx":32191
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: MASTER PELANGGAN ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   4380
         TabIndex        =   164
         Top             =   60
         Width           =   2775
      End
      Begin VB.Label lblJml 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10320
         TabIndex        =   9
         Top             =   60
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11400
         Y1              =   9240
         Y2              =   9240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11400
         Y1              =   9960
         Y2              =   9960
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11400
      End
   End
End
Attribute VB_Name = "frmSIDMasterPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private rsFind  As ADODB.Recordset
Private rsMove  As ADODB.Recordset
Private rsBy    As New ADODB.Recordset
Private onFind  As Boolean
Private onF12   As Boolean
Private xWAC    As Boolean
Private passBy  As String
Private nValid  As String
Private xAs()   As String
Private kliPr   As Boolean
Private kliKo   As Boolean
Private kliKc   As Boolean
Private kliKl   As Boolean
Private kliPrp  As Boolean
Private kliKop  As Boolean
Private kliKcp  As Boolean
Private kliKlp  As Boolean
Private kliPrs  As Boolean
Private kliKos  As Boolean
Private kliKcs  As Boolean
Private kliKls  As Boolean

Private nDpn As String
Private nTgh As String
Private nBlk As String
Private kodeLgnLama As String
Private isEditData As Boolean
Private isLoad As Boolean
Private noRangka As String
Private isBolehArmst As Boolean

Public Sub isEdit(Optional value As Boolean = False)
    isEditData = value
End Sub

Public Property Let setKodeLgnLama(value As String)
    kodeLgnLama = NulSID(value)
End Property

Public Property Let setNoRangka(value As String)
    noRangka = NulSID(value)
End Property

Public Property Let MyRsBy(value As ADODB.Recordset)
    Set rsBy = value
End Property

Public Property Let MyPassBy(value As String)
    passBy = value
End Property

Public Property Get rsMyMove() As ADODB.Recordset
    Set rsMyMove = rsMove
End Property

Public Property Let WAC(value As Boolean)
    xWAC = value
End Property

Public Property Let Asura(value As String)
    xAs() = Split(value, ";")
End Property

Private Sub cboAgamaS_GotFocus()
    If cboAgamaS.BackColor = fColor Then
        cboAgamaS.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboAgamaS.Locked = False
    End If
End Sub

Private Sub cboKewarganegaraan_Change()
    If UCase(cboTipe.text) = "PERORANGAN" Then
        If UCase(cboKewarganegaraan.text) = "WNI" Then
            Label5(0).Caption = "No. KTP"
        ElseIf UCase(cboKewarganegaraan.text) = "WNA" Then
            Label5(0).Caption = "Passport / KITAS"
        End If
    End If
End Sub

Private Sub cboKewarganegaraan_Click()
    Call cboKewarganegaraan_Change
End Sub

Private Sub cboOlah_GotFocus()
    If cboOlah.BackColor = fColor Then
        cboOlah.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboOlah.Locked = False
    End If
End Sub

Private Sub cmdSimpan_Click()
         Dim kodeLgn As String
1        On Error GoTo cmdSimpan_Click_Error
          Dim boleh As Boolean
          Dim lengkapi As String
          Dim isBegin As Boolean
          Dim vStatus As String
2         boleh = True
3         lengkapi = ""
          
4         If UCase(cboTipe.text) <> "PERORANGAN" And NulSID(cboTipe.text) <> "" Then
5             If NulSID(Txt(18)) = "" Then
6                 boleh = False
7                 lengkapi = lengkapi & " Nama PIC,"
8             End If
9             If NulSID(Txt(19)) = "" Then
10                boleh = False
11                lengkapi = lengkapi & " Telp PIC,"
12            End If
13        End If
          
14        If NulSID(Txt(1)) = "" Then
15            boleh = False
16            lengkapi = lengkapi & " Kode Customer,"
17        End If
          
          ''NAMA LENGKAP
18        If NulSID(Txt(4) & Txt(5) & Txt(6)) = "" Then
19            boleh = False
20            lengkapi = lengkapi & " Nama Customer,"
21        End If
          
          ''EMAIL
      '    If nulsid(txt(16)) = "" Then
      '        boleh = False
      '        lengkapi = lengkapi & " Email,"
      '    End If
          
          ''ALAMAT (provinsi, kab/kota, kecamatan, kelurahan, jalan)
      '    If cmdBack.Visible = False Then
      '        If cboProv.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Provinsi,"
      '        End If
      '
      '        If cboKota.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Kota,"
      '        End If
      '
      '        If cboKec.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Kecamatan,"
      '        End If
      '
      '        If cboKel.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Kelurahan,"
      '        End If
      '
      '        If nulsid(txt(8)) = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Jalan/Kompleks,"
      '        End If
      '    Else
22            If NulSID(Txt(63)) = "" Then
23                boleh = False
24                lengkapi = lengkapi & " Provinsi,"
25            End If
              
26            If NulSID(Txt(64)) = "" Then
27                boleh = False
28                lengkapi = lengkapi & " Kota,"
29            End If
              
30            If NulSID(Txt(65)) = "" Then
31                boleh = False
32                lengkapi = lengkapi & " Kecamatan,"
33            End If
              
34            If NulSID(Txt(66)) = "" Then
35                boleh = False
36                lengkapi = lengkapi & " Kelurahan,"
37            End If
              
38            If NulSID(Txt(8)) = "" Then
39                boleh = False
40                lengkapi = lengkapi & " Jalan/Kompleks,"
41            End If
      '    End If
          
42        If NulSID(Txt(2)) = "" Or Val(NulSID(Txt(2))) = "0" Then
43            MsgBox Label5(0).Caption & " wajib diisi sebelum VALIDASI ORDER", vbInformation
44            If Label5(0).Caption = "No. KTP" Then
45                Txt(2) = "0000000000000000"
46            End If
47        End If
          
48        If NulSID(Txt(14)) = "" And NulSID(Txt(15)) = "" Then
49            boleh = False
50            lengkapi = lengkapi & " No HP,"
51        End If
          
52        If NulSID(DtLahir.value) <> "" Then
53            If DateDiff("D", Format(Now, "dd/mm/yyyy"), Format(DtLahir.value, "dd/mm/yyyy")) > 0 Then
54                boleh = False
55                lengkapi = lengkapi & " Tanggal Lahir melebihi tanggal hari ini,"
56            End If
57        End If
          
58        If boleh = True Then
              'If Get_ConnectionAcces(NisDb, NisUSER, NisPass, NisKomp) Then
                  Dim rsCari As ADODB.Recordset
59                sSQL = "SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(Txt(63)) & "' and KOTA = '" & NulSID(Txt(64)) & "' and KECAMATAN = '" & NulSID(Txt(65)) & "' and KELURAHAN = '" & NulSID(Txt(66)) & "' and KODEPOS = '" & NulSID(Txt(11)) & "' "
60                Set rsCari = objO.FetchData(sSQL)
61                If rsCari.EOF Then
62                    MsgBox "PROPINSI / KOTA / KECAMATAN / KELURAHAN / KODEPOS tidak cocok", vbInformation
63                    Exit Sub
64                End If
65                If Not (NulSID(Txt(2)) = "" Or Val(NulSID(Txt(2))) = "0") Then
66                    sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE idcust = '" & NulSID(Txt(1)) & "' "
67                    Set rsCari = objO.FetchData(sSQL)
68                    If Not rsCari.EOF Then
69                        If NulSID(Txt(2)) <> NulSID(rsCari!noKTP) Then
70                            sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NOKTP = '" & NulSID(Txt(2)) & "' "
71                            Set rsCari = objO.FetchData(sSQL)
72                            If Not rsCari.EOF Then
                                  'MsgBox "No KTP sudah ada, tidak boleh sama.", vbInformation\
73                                MsgBox Label5(0) & " sudah ada, tidak boleh sama.", vbInformation
74                                Exit Sub
75                            End If
76                        End If
77                    End If
78                End If
                  
                  ''20180130 KELOLOSAN KODELGN LAMA TIDAK TERISI DI BRIDGING
                  ''------------------------------------------------------------------------
79                If kodeLgnLama = "" Then
80                    sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE idcust = '" & NulSID(Txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' "
81                    Set rsCari = objO.FetchData(sSQL)
82                    If Not rsCari.EOF Then
83                        kodeLgnLama = NulSID(rsCari!kodeLgn)
84                    Else
85                        kodeLgnLama = NulSID(Txt(1))
86                    End If
87                Else
88                    sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE kodelgn = '" & NulSID(kodeLgnLama) & "' and idcust = '" & NulSID(Txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' "
89                    Set rsCari = objO.FetchData(sSQL)
90                    If rsCari.EOF Then
                          ''jika tidak ada kodelgn lama dan idcust , cek di bridging untuk kodelgn baru
91                        sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE idcust = '" & NulSID(Txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' "
92                        Set rsCari = objO.FetchData(sSQL)
93                        If Not rsCari.EOF Then
94                            kodeLgnLama = NulSID(rsCari!kodeLgn)
95                        Else
96                            kodeLgnLama = NulSID(Txt(1))
97                        End If
98                    End If
99                End If
                  
100               If kodeLgnLama <> NulSID(Txt(1)) And kodeLgnLama <> "" Then
101                   kodeLgn = kodeLgnLama
102               Else
103                   kodeLgn = NulSID(Txt(1))
104               End If
                  
105               If isEditData = False Then
106                   sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE idcust = '" & NulSID(Txt(1)) & "' "
107                   Set rsCari = objO.FetchData(sSQL)
108                   If Not rsCari.EOF Then
109                       MsgBox "Id Cust sudah ada atas nama " & NulSID(NulSID(rsCari!namacust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan lakukan proses simpan ulang.", vbInformation
110                       Txt(1).text = LastID(cboTipe.text)
111                       Exit Sub
112                   End If
                      
113                   Set rsCari = objO.FetchDataACC("SELECT NamaLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & NulSID(Txt(1)) & "'")
114                   If rsCari.RecordCount > 0 Then
115                       MsgBox "Id Cust sudah ada atas nama " & NulSID(NulSID(rsCari!Namalgn)) & vbCrLf & "Silahkan lakukan proses simpan ulang.", vbInformation
116                       Txt(1).text = LastID(cboTipe.text)
117                       Exit Sub
118                   End If
                      
                      ''MENGECEK NO KTP / NO NPWP
119                   If UCase(cboTipe.text) = "PERORANGAN" Then
                          ''KTP TIDAK BOLEH KEMBAR
120                       If NulSID(Txt(2)) = "" Or Val(NulSID(Txt(2))) = 0 Then
121                           Txt(2) = ""
122                       Else
123                           sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NoKTP = '" & NulSID(Txt(2)) & "' "
124                           Set rsCari = objO.FetchData(sSQL)
125                           If Not rsCari.EOF Then
                                  'MsgBox "No KTP sudah ada atas nama " & NulSID(NulSID(rsCari!NamaCust1) & " " & NulSID(rsCari!NamaCust2) & " " & NulSID(rsCari!NamaCust3)) & vbCrLf & "Silahkan diisi dengan data lain.", vbInformation
126                               MsgBox Label5(0) & " sudah ada atas nama " & NulSID(NulSID(rsCari!namacust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan diisi dengan data lain.", vbInformation
                                  'txt(1).Text = LastID(cboTipe.Text)
127                               Exit Sub
128                           End If
129                       End If
130                   Else
                          ''NPWP TIDAK BOLEH KEMBAR
131                       If NulSID(txtNpwp) = "" Or Val(NulSID(txtNpwp)) = 0 Then
132                           txtNpwp.text = "00.000.000.0-000.000"
133                       Else
134                           sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE noNPWP = '" & NulSID(txtNpwp) & "' "
135                           Set rsCari = objO.FetchData(sSQL)
136                           If Not rsCari.EOF Then
137                               MsgBox "No NPWP sudah ada atas nama " & NulSID(NulSID(rsCari!namacust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan diisi dengan data lain.", vbInformation
                                  'txt(1).Text = LastID(cboTipe.Text)
138                               Exit Sub
139                           End If
140                       End If
141                   End If
142               End If
                  
143               If NulSID(txtNpwp) = "" Or Val(NulSID(txtNpwp)) = 0 Then
144                   txtNpwp.text = "00.000.000.0-000.000"
145               ElseIf InStr(1, txtNpwp.text, "_") Then
146                   MsgBox "No NPWP belum lengkap.", vbInformation
                      'txt(1).Text = LastID(cboTipe.Text)
147                   Exit Sub
148               End If
                    
                  Dim namaLengkap As String
                  Dim Alamat As String
                  Dim provinsi As String
                  Dim Kota As String
                  Dim kec As String
                  Dim kel As String
149               provinsi = "": Kota = "": kec = "": kel = ""
                  
150               Alamat = NulSID(Txt(8))
                  
                  'If cmdBack.Visible = True Then
151                   Alamat = Alamat & " Kel." & NulSID(Txt(66)) & " Kec." & NulSID(Txt(65)) & " " & NulSID(Txt(64)) & " Prov." & NulSID(Txt(63))
152                   provinsi = UCase(NulSID(Txt(63)))
153                   Kota = UCase(NulSID(Txt(64)))
154                   kec = UCase(NulSID(Txt(65)))
155                   kel = UCase(NulSID(Txt(66)))
                      
      '            Else
      '                Alamat = Alamat & " Kel." & nulsid(cboKel.Text) & " Kec." & nulsid(cboKec.Text) & " " & nulsid(cboKota.Text) & " Prov." & nulsid(cboProv.Text)
      '                provinsi = UCase(cboProv.Text)
      '                Kota = UCase(cboKota.Text)
      '                kec = UCase(cboKec.Text)
      '                kel = UCase(cboKel.Text)
      '            End If
                  
156               namaLengkap = NulSID(NulSID(Txt(4)) & " " & NulSID(Txt(5)) & " " & NulSID(Txt(6)))
                  
                  
                  
157               If kodeLgnLama = "" Then
158                   kodeLgnLama = NulSID(Txt(1))
159               End If
                  
                  '20/12/2012  Cust Non Axtif
                  Dim rsNon As ADODB.Recordset
160               Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & NulSID(kodeLgn) & "' and nonaktifflag  = 1")
161               If rsNon.RecordCount > 0 Then
162                   MsgBox "Kode Langgan : " & Txt(1) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut", vbCritical + vbOKOnly, "Master Pelanggan Non Aktif"
163                   pelangganAda = False
164                   Exit Sub
165               End If
                  
                  ''sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NamaCust1 = '" & nulsid(txt(4)) & "' and NamaCust2 = '" & nulsid(txt(5)) & "' and NamaCust3 = '" & nulsid(txt(6)) & "' and Pekerjaan = '" & UCase(nulsid(txt(7))) & "' and Alamat = '" & nulsid(txt(8)) & "' "
                  ''20171227 || 7.  Saat save, untuk pengecekan data duplikasi yang sebelumnya menggunakan Nama+Alamat+Pekerjaan, untuk Alamat, nama jalan/kompleks/no tidak dicek
                  Dim namaSama As String
166               If NulSID(Txt(4)) = "" Then
167                   MsgBox "Nama Customer 1 harus diisi terlebih dahulu", vbInformation
168                   Exit Sub
169               End If
                  
170               namaSama = "(namacust1+NamaCust2+NamaCust3 LIKE '%" & IIf(NulSID(Txt(4)) <> "", NulSID(Txt(4)) & "%", "") & IIf(NulSID(Txt(5)) <> "", NulSID(Txt(5)) & "%", "") & IIf(NulSID(Txt(6)) <> "", NulSID(Txt(6)) & "%", "") & "%')"
                  
                  'sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NamaCust1 = '" & NulSID(Txt(4)) & "' and NamaCust2 = '" & NulSID(Txt(5)) & "' and NamaCust3 = '" & NulSID(Txt(6)) & "' and Pekerjaan = '" & UCase(NulSID(Txt(7))) & "' and Propinsi = '" & NulSID(Txt(63)) & "' and Kota = '" & NulSID(Txt(64)) & "' and Kecamatan = '" & NulSID(Txt(65)) & "' and Kelurahan = '" & NulSID(Txt(66)) & "'"
171               sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE " & namaSama & " and Pekerjaan = '" & UCase(NulSID(Txt(7))) & "' and Propinsi = '" & NulSID(Txt(63)) & "' and Kota = '" & NulSID(Txt(64)) & "' and Kecamatan = '" & NulSID(Txt(65)) & "' and Kelurahan = '" & NulSID(Txt(66)) & "'"
                  
172               Set rsNon = objO.FetchData(sSQL)
173               If rsNon.RecordCount > 0 And isEditData = False Then
174                   MsgBox "Nama Customer, Alamat dan Pekerjaan sudah ada dengan ID Pelanggan " & rsNon!idcust & vbCrLf & "Silahkan ganti dengan data yang lain.", vbInformation
                      'pelangganAda = False
175                   Exit Sub
176               End If
                  
177               If NulSID(Txt(2)) <> "" And Val(NulSID(Txt(2))) <> 0 And isEditData = True Then
178                   sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NoKTP = '" & NulSID(Txt(2)) & "' and idcust <> '" & NulSID(Txt(1)) & "' "
179                   Set rsNon = objO.FetchData(sSQL)
180                   If rsNon.RecordCount > 0 Then
                          'MsgBox "No KTP " & NulSID(txt(2)) & " sudah ada." & vbCrLf & "Silahkan ganti dengan data yang lain.", vbInformation
181                       MsgBox Label5(0) & " " & NulSID(Txt(2)) & " sudah ada." & vbCrLf & "Silahkan ganti dengan data yang lain.", vbInformation
                          'pelangganAda = False
182                       Exit Sub
183                   End If
184               End If
                  
      '            If Left(Trim(Text(57)), 3) = "AFL" Then
      '                MsgBox "Tidak diperkenankan membuat Kode Pelanggan Afiliasi.", vbCritical + vbOKOnly, "Master Pelanggan Salah"
      '                pelangganAda = False
      '                Exit Sub
      '            End If
185               isBegin = False
186               Call objHT.OpenTransactionAll
187               isBegin = True
188               If UCase(cboTipe.text) = UCase("Pemerintahan") Then
189                   vStatus = "P"
190               Else
191                   vStatus = "S"
192               End If
                  
193              Set rsSearch = objO.FetchDataACC("select KodeLgn from ARMST WITH(NOLOCK) where KodeLgn = '" & NulSID(kodeLgn) & "'")
194              If rsSearch.EOF Then
                     '' insert data armst
195                  sSQL1 = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                     "        StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,DBSumber,Keterangan,ClientName,KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak, SIUP, No_Mou ) "
196                  sSQL1 = sSQL1 & " VALUES ( '" & NulSID(kodeLgn) & "', " & _
                     "                  '" & s2d(NulSID(namaLengkap)) & "', " & _
                     "                  '" & s2d(Left(NulSID(kodeLgn), 3)) & "', " & _
                     "                  '" & s2d(NulSID(Txt(8))) & "', " & _
                     "                  '" & s2d(NulSID(Txt(66)) + ", " + NulSID(Txt(65))) & "', " & _
                     "                  '" & s2d(NulSID(Txt(64)) + ", " + NulSID(Txt(63))) & "', " & _
                     "                  '" & s2d(NulSID(Txt(14))) & "', " & _
                     "                  '" & s2d(NulSID(Txt(12)) & NulSID(Txt(13))) & "', " & _
                     "                  '" & s2d(NulSID(txtNpwp)) & "', " & _
                     "                  'D0', 0, " & _
                     "                  '" & s2d(NulSID(Txt(18))) & "', 'GRP', " & _
                     "                  '" & s2d(vStatus) & "', '', " & _
                     "                  " & "0" & ", 'Customer " & cboTipe.text & "', 0, " & _
                     "                  '" & gSA & "', GETDATE(),0,0,0, " & _
                     "                  '" & Trim(gDb) & "','Data Baru','" & CompName & "' ,'" & gSA & "',Getdate(), " & _
                     "                  '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', " & _
                     "                  " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                     "                  " & 0 & ", '" & 0 & "', '', '', '' ) "   'A.Rev1711291-SingleId

197                  sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3" _
                          & " ,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry,HubungDengan,Telepon1)" _
                         & " VALUES ('" & NulSID(kodeLgn) & "','" & NulSID(namaLengkap) & "','" & Left(NulSID(kodeLgn), 3) & "'" _
                         & ",'" & s2d(NulSID(Txt(8))) & "','" & s2d(NulSID(Txt(66)) + ", " + NulSID(Txt(65))) & "','" & s2d(NulSID(Txt(64)) + ", " + NulSID(Txt(63))) & "','" & Replace(Replace(NulSID(Txt(12) & Txt(13)) & "," & NulSID(Txt(14)) & "," & NulSID(Txt(15)), ">", ""), " ", "") & "','" & NulSID(Txt(17)) & "'" _
                         & ",'" & NulSID(txtNpwp) & "','D0','D0','D0','D0','D0','" & vStatus & "','" & gSA & "','" & Format(isnow, "MM/DD/YYYY") & "','" & NulSID(Txt(18)) & "','" & NulSID(Txt(19)) & "')"
                     
198                   Call objHT.InsertData_ACC(sSQL1)
199                   Call objHT.InsertData_ACC(sSQL)
                       
200              Else
201                  If isBolehArmst = True Then
202                      sSQL1 = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                                                           "StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,No_MOU,DBSumber,Keterangan,ClientName, " & _
                                                                           "KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak,SIUP ) " & _
                                           "SELECT KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                                                           "StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, GETDATE(),FPGabungan,NonPPn,Afiliasi,No_MOU, '" & Trim(gDb) & "','Sebelum Diedit','" & CompName & "' ,'" & gSA & "'," & _
                                           "Getdate(),TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak,SIUP from Armst WHERE KodeLgn='" & NulSID(kodeLgn) & "' "
         
                          
203                      sSQL = "UPDATE ARMST SET NamaLgn = '" & NulSID(namaLengkap) & "'," _
                              & " KodeWil = '" & Left(NulSID(kodeLgn), 3) & "',Alamat1 = '" & s2d(NulSID(Txt(8))) & "',Alamat2 = '" & s2d(NulSID(Txt(66)) + ", " + NulSID(Txt(65))) & "',Alamat3 = '" & s2d(NulSID(Txt(64)) + ", " + NulSID(Txt(63))) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                             & " Telepon = '" & Replace(Replace(NulSID(Txt(12) & Txt(13)) & "," & NulSID(Txt(14)) & "," & NulSID(Txt(15)), ">", ""), " ", "") & "',Npwp = '" & NulSID(txtNpwp) & "'," _
                             & " TOPGRP = 'D0',StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "',Nofax = '" & NulSID(Txt(17)) & "',HubungDengan = '" & NulSID(Txt(18)) & "',Telepon1 = '" & NulSID(Txt(19)) & "'" _
                             & " where KodeLgn = '" & NulSID(kodeLgn) & "'"
204                       Call objHT.InsertData_ACC(sSQL1)
205                       Call objHT.InsertData_ACC(sSQL)
                           
206                  End If
207              End If
                 
                  
208               Set rsSearch = objHT.FetchDataBengkel("select IDCust from SINGLEIDORI..NEWCUSTOMER WITH(NOLOCK) where IDCust = '" & NulSID(Txt(1)) & "'")
209               If rsSearch.EOF Then
210                   sSQL = "INSERT INTO SINGLEIDORI..NEWCUSTOMER (IDCust,TypeCustomer,NoKTP,noNPWP,Title,NamaCust1,NamaCust2,NamaCust3,tglLahir,Agama,Pekerjaan,Propinsi,Kota,Kecamatan,Kelurahan,KodePos,alamat,RT,RW,IdTelp,telp,HP1,HP2,Email,Hobi,Olahraga,Suku,Preferensi,FPGab,NonPPn,TopVhc,BatasVhc,TopGrp,BatasGrp,TopBp,BatasBp,TopPart,BatasPart,TglEntry,UserEntry,NoFax,namaPIC,telpPIC,kewarganegaraan)"
211                   sSQL = sSQL & " VALUES ('" & NulSID(Txt(1)) & "','" & UCase(NulSID(cboTipe.text)) & "','" & UCase(NulSID(Txt(2))) & "','" & NulSID(txtNpwp) & "','" & UCase(NulSID(cboTitle.text)) & "','" & UCase(NulSID(Txt(4))) & "','" & UCase(NulSID(Txt(5))) & "','" & UCase(NulSID(Txt(6))) & "','" & Format(DtLahir.value, "mm/dd/yyyy") & "','" & UCase(Txt(20).text) & "','" & IIf(UCase(NulSID(Txt(7))) = "LAINNYA", UCase(NulSID(Txt(23))), UCase(NulSID(Txt(7)))) & "','" & provinsi & "','" & Kota & "','" & kec & "','" & kel & "',"
212                   sSQL = sSQL & " '" & UCase(NulSID(Txt(11))) & "','" & UCase(NulSID(Txt(8))) & "','" & UCase(NulSID(Txt(9))) & "','" & UCase(NulSID(Txt(10))) & "','" & UCase(NulSID(Txt(12))) & "','" & UCase(NulSID(Txt(13))) & "','" & UCase(NulSID(Txt(14))) & "','" & UCase(NulSID(Txt(15))) & "','" & UCase(NulSID(Txt(16))) & "',"
213                   sSQL = sSQL & " '" & UCase(NulSID(Txt(21).text)) & "','" & UCase(NulSID(Txt(22).text)) & "','" & UCase(NulSID(Txt(71))) & "','" & UCase(NulSID(Txt(72))) & "',0,0,'D0',0,'D0',0,'D0',0,'D0',0,GETDATE(),'" & gSA & "','" & NulSID(Txt(17)) & "','" & NulSID(Txt(18)) & "','" & NulSID(Txt(19)) & "','" & NulSID(cboKewarganegaraan.text) & "')"
214               Else
215                   sSQL = "UPDATE SINGLEIDORI..NEWCUSTOMER SET NoKTP = '" & UCase(NulSID(Txt(2))) & "',noNPWP = '" & UCase(NulSID(txtNpwp)) & "',Title = '" & UCase(cboTitle.text) & "',NamaCust1='" & UCase(NulSID(Txt(4))) & "',NamaCust2='" & UCase(NulSID(Txt(5))) & "',NamaCust3='" & UCase(NulSID(Txt(6))) & "',tglLahir='" & Format(DtLahir.value, "mm/dd/yyyy") & "',Agama='" & UCase(Txt(20).text) & "',Pekerjaan='" & IIf(UCase(NulSID(Txt(7))) = "LAINNYA", UCase(NulSID(Txt(23))), UCase(NulSID(Txt(7)))) & "',Propinsi='" & provinsi & "',Kota='" & Kota & "',Kecamatan='" & kec & "',"
216                   sSQL = sSQL & " Kelurahan='" & kel & "',KodePos='" & UCase(NulSID(Txt(11))) & "',alamat='" & UCase(NulSID(Txt(8))) & "',RT='" & UCase(NulSID(Txt(9))) & "',RW='" & UCase(NulSID(Txt(10))) & "',IdTelp='" & UCase(NulSID(Txt(12))) & "',telp='" & UCase(NulSID(Txt(13))) & "',HP1='" & UCase(NulSID(Txt(14))) & "',HP2='" & UCase(NulSID(Txt(15))) & "',Email='" & UCase(NulSID(Txt(16))) & "',Hobi='" & UCase(NulSID(Txt(21).text)) & "',Olahraga='" & UCase(NulSID(Txt(22).text)) & "',Suku='" & UCase(NulSID(Txt(71))) & "',Preferensi='" & UCase(NulSID(Txt(72))) & "',NoFax = '" & NulSID(Txt(17)) & "',namaPIC= '" & NulSID(Txt(18)) & "',telpPIC= '" & NulSID(Txt(19)) & "',kewarganegaraan = '" & NulSID(cboKewarganegaraan.text) & "',UserUpdate = '" & NulSID(gSA) & "',TglUpdate = GETDATE(),DealerUpdate = '" & Right(gDb, 2) & "', AppUpdate = 'SVC' WHERE IDCust = '" & NulSID(Txt(1)) & "' "
217               End If
                  
218               Call objHT.InsertData_Bengkel(sSQL)
                  
219               Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM SINGLEIDORI..Bridging WHERE idcust = '" & NulSID(Txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' ")
220               If rsSearch.EOF Then
221                   sSQL = "INSERT INTO SINGLEIDORI..Bridging(KodeDealer,kodeLgn,IDCust,TglEntry,UserEntry) VALUES ('" & Right(gDb, 2) & "','" & kodeLgnLama & "','" & NulSID(Txt(1)) & "',getdate(),'" & gSA & "')"
222                   Call objHT.InsertData_Bengkel(sSQL)
223               End If
                  
                  
                  
224               Call objHT.CloseTransactionAll
                  
225               gBantuan = Trim(Txt(1))
      '            Call hapusForm
      '            Text(57) = ""
                  
226               With frmtCustomer
227                   If .Visible Then
228                       .text(2) = Trim(gBantuan)
229                       Unload Me
230                   End If
231               End With
                  
232               With frmBooking
233                   If .Visible Then
234                       .text(2) = Trim(gBantuan)
235                       Unload Me
236                   End If
237               End With
238               MsgBox "Data berhasil disimpan", vbInformation
239               Unload Me
              'End If
240       Else
241           MsgBox "Silahkan lengkapi data berikut:" & vbCrLf & lengkapi, vbInformation
242           Exit Sub
243       End If

244      On Error GoTo 0
245      Exit Sub
          

246      On Error GoTo 0
247      Exit Sub

cmdSimpan_Click_Error:
248       If isBegin Then
249           Call objHT.CloseTransactionAll_Err
250       End If
251       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmSIDMasterPelanggan"
End Sub

Private Sub DTLahir_GotFocus()
    If DtLahir.BackColor = fColor Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Activate()
    If isEditData Then
        
        
        Txt(4).BackColor = &H80FFFF
        Txt(5).BackColor = &H80FFFF
        Txt(6).BackColor = &H80FFFF
        
        Txt(14).BackColor = &H80FFFF
        
        
        cboTipe.Enabled = False
        Txt(1).Enabled = False
        If NulSID(Txt(1).text) <> "" And isLoad = False Then
            Call callPanggilDataNewCustomer
        End If
    End If
End Sub

Private Sub Form_Load()
    isLoad = False
    isBolehArmst = True
    Me.Caption = "" 'TitleBP
    Move 0, objA.MyTop + 370
    Me.Width = objA.MyLeft
   ' Me.Height = fHeight - 260
   gMainKeyCari = ""
    'Frame1.Left = (((objA.MyLeft - 480) - Val(Frame1.Width)) / 2) + 495
    Call SetControlForm("load")
    'Call ViewPassing
    Call Load_Cbo
    FrAlm.Visible = True
    If isEditData Then
        
        
        Txt(4).BackColor = &H80FFFF '&HC0E0FF
        Txt(5).BackColor = &H80FFFF ' &HC0E0FF
        Txt(6).BackColor = &H80FFFF ' &HC0E0FF
        
        cboTipe.Enabled = False
        Txt(1).Enabled = False
        If NulSID(Txt(1).text) <> "" And isLoad = False Then
            Call callPanggilDataNewCustomer
        End If
    End If
    'isLoad = True
End Sub

Private Sub callPanggilDataNewCustomer()
          Dim rsCust As ADODB.Recordset
1         sSQL = "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE no_chassis = '" & NulSID(noRangka) & "'"
2         Set rsCust = objO.FetchData(sSQL)
3         If Not rsCust.EOF Then
4             kodeLgnLama = ""
5             kodeLgnLama = NulSID(rsCust!FS_KD_PLG)
6             sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE KodeDealer = '" & Right(gDb, 2) & "' AND IDCUST= '" & NulSID(Txt(1).text) & "' AND KODELGN = '" & NulSID(kodeLgnLama) & "'"
7             Set rsCust = objO.FetchData(sSQL)
8             If rsCust.EOF Then
9                 isBolehArmst = False
10            End If
11        Else
12            isBolehArmst = False
13        End If
          
          
14        sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST= '" & NulSID(Txt(1).text) & "'"
15        Set rsCust = objO.FetchData(sSQL)
16        If Not rsCust.EOF Then
17            Txt(1) = NulSID(rsCust!idcust)
18            Txt(2) = NulSID(rsCust!noKTP)
19            If NulSID(rsCust!noNPWP) <> "" Then
                  'txtNpwp.Mask = ""
20                txtNpwp.text = setNpwpSID(Replace(Replace(NulSID(rsCust!noNPWP), ".", ""), "-", ""))
                  'txtNpwp.Mask = "00.000.000.0-000.000" '"##.###.###.#-###.###"
      '        Else
      '        txtNpwp.Text = ""
21            Else
22                txtNpwp.text = "00.000.000.0-000.000"
23            End If
24            Txt(4) = NulSID(rsCust!namacust1)
25            Txt(5) = NulSID(rsCust!namacust2)
26            Txt(6) = NulSID(rsCust!namacust3)
              
27            Txt(4).Enabled = False
28            Txt(5).Enabled = False
29            Txt(6).Enabled = False
              
30            DtLahir.value = NulSID(rsCust!tglLahir)
              
31            Set rsHelp = objO.FetchData("SELECT DISTINCT agama FROM [SingleIdOri]..agama WHERE agama = '" & NulSID(rsCust!Agama) & "'")
32            If Not rsHelp.EOF Then
                  'cboAgama.Text = nulsid(rsCust!Agama)
33                Txt(20).text = NulSID(rsCust!Agama)
34            End If
35            cboKewarganegaraan.Clear
36            cboKewarganegaraan.AddItem ("WNI")
37            cboKewarganegaraan.AddItem ("WNA")
              
38            If NulSID(rsCust!kewarganegaraan) <> "" Then
39                cboKewarganegaraan.text = NulSID(rsCust!kewarganegaraan)
40            End If
41            cboTipe.text = NulSID(rsCust!TypeCustomer)
42            Call cboTipe_Change
              
43            If UCase(cboTipe.text) <> "PERORANGAN" And NulSID(cboTipe.text) <> "" Then
44                Txt(18).Enabled = True
45                Txt(19).Enabled = True
                  
46            Else
47                Txt(18).text = ""
48                Txt(19).text = ""
49                Txt(18).Enabled = False
50                Txt(19).Enabled = False
51                Txt(18).BackColor = fColor
52                Txt(19).BackColor = fColor
53            End If
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE PROPINSI = '" & nulsid(rsCust!Propinsi) & "'")
      '        If Not rsHelp.EOF Then
      '            cboProv.text = UCase(nulsid(rsCust!Propinsi))
      '            Call cboProv_Change
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE kota = '" & nulsid(rsCust!Kota) & "'")
      '            If Not rsHelp.EOF Then
      '                cboKota.text = UCase(nulsid(rsCust!Kota))
      '                Call cboKota_Change
      '            End If
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE kota = '" & nulsid(rsCust!Kota) & "' and KECAMATAN = '" & nulsid(rsCust!Kecamatan) & "'")
      '            If Not rsHelp.EOF Then
      '                cboKec.text = UCase(nulsid(rsCust!Kecamatan))
      '                Call cboKec_Change
      '            End If
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE KECAMATAN = '" & nulsid(rsCust!Kecamatan) & "' and kelurahan = '" & nulsid(rsCust!kelurahan) & "'")
      '            If Not rsHelp.EOF Then
      '                cboKel.text = (nulsid(rsCust!kelurahan))
      '            Call cboKel_Change
      '            End If
      '
      '        Else
      '            'Call hideComboProv(False)]
      '            FrAlm.Visible = False
54                cmdBack.Visible = False
55                Txt(63) = UCase(NulSID(rsCust!Propinsi))
56                Txt(64) = UCase(NulSID(rsCust!Kota))
57                Txt(65) = UCase(NulSID(rsCust!kecamatan))
58                Txt(66) = UCase(NulSID(rsCust!kelurahan))
      '        End If
               
59            Txt(8) = NulSID(rsCust!Alamat)
60            Txt(9) = Format(NulSID(rsCust!rt), "0##") 'nulsid(rsCust!RT)
61            Txt(10) = Format(NulSID(rsCust!rw), "0##")
62            Txt(11) = NulSID(rsCust!KodePos)
              
63            Txt(12) = NulSID(rsCust!IdTelp)
64            Txt(13) = NulSID(rsCust!telp)
65            Txt(14) = NulSID(rsCust!hp1)
66            Txt(15) = NulSID(rsCust!hp2)
67            Txt(16) = NulSID(rsCust!email)
              
68            Set rsHelp = objO.FetchData("SELECT DISTINCT Pekerjaan FROM [SingleIdOri]..Pekerjaan WHERE Pekerjaan = '" & NulSID(rsCust!Pekerjaan) & "'")
69            If Not rsHelp.EOF Then
70                Txt(7) = UCase(NulSID(rsCust!Pekerjaan))

71            Else
72                Txt(7) = "LAINNYA"
                 ' Call cboPek_Change
                 ' txt(7) = UCase(nulsid(rsCust!Pekerjaan))
73                Txt(23) = UCase(NulSID(rsCust!Pekerjaan))
74                Txt(23).BackColor = &HFFFFFF
75            End If
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT hobi FROM [SingleIdOri]..hobi WHERE hobi = '" & nulsid(rsCust!Hobi) & "'")
      '        If Not rsHelp.EOF Then
      '            cboHobi.Text = UCase(nulsid(rsCust!Hobi))
      '        End If
              
76            Txt(21) = UCase(NulSID(rsCust!hobi))
77            Txt(22) = UCase(NulSID(rsCust!olahraga))
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT olahraga FROM [SingleIdOri]..olahraga WHERE olahraga = '" & nulsid(rsCust!olahraga) & "'")
      '        If Not rsHelp.EOF Then
      '            cboOlah.Text = UCase(nulsid(rsCust!olahraga))
      '        End If

78            Set rsHelp = objO.FetchData("SELECT DISTINCT Title FROM [SingleIdOri]..Title WHERE Title = '" & NulSID(rsCust!Title) & "'")
79            If Not rsHelp.EOF Then
80                cboTitle.text = UCase(NulSID(rsCust!Title))
81            End If
              
82            Txt(71) = UCase(NulSID(rsCust!Suku))
83            Txt(72) = UCase(NulSID(rsCust!Preferensi))
              
84            Txt(17) = UCase(NulSID(rsCust!NoFax))
85            Txt(18) = UCase(NulSID(rsCust!NamaPIC))
86            Txt(19) = UCase(NulSID(rsCust!TelpPIC))
87             isLoad = True
88        End If
End Sub

Private Sub SetControlForm(nSET As String)
On Error GoTo ErrNull
    With Me
    Select Case nSET
        Case "load"
            '.txt(1).BackColor = &HC0E0FF
            .Txt(1).text = ""
            .Txt(1).Enabled = False
            For i = 2 To 72
                '.txt(i).BackColor = tColor
                .Txt(i).text = ""
            Next i
            
            .cboAgama.Clear
            .cboPek.Clear
            .cboHobi.Clear
            .cboOlah.Clear
            
            .cboProv.Clear
            .cboKota.Clear
            .cboKec.Clear
            .cboKel.Clear
            
            '.cboProvP.Clear
'            .cboKotaP.Clear
'            .cboKecP.Clear
'            .cboKelP.Clear
            
            .cboAgamaS.Clear
            .cboPekS.Clear
            .cboHobiS.Clear
            .cboOlahS.Clear
            
            .cboProvS.Clear
            .cboKotaS.Clear
            .cboKecS.Clear
            .cboKelS.Clear
            
            .cboTitle.BackColor = fColor
            .cboTipe.BackColor = fColor
            
            .cboAgama.BackColor = fColor
            .cboPek.BackColor = fColor
            .cboHobi.BackColor = fColor
            .cboOlah.BackColor = fColor
            
            .cboProv.BackColor = fColor
            .cboKota.BackColor = fColor
            .cboKec.BackColor = fColor
            .cboKel.BackColor = fColor
            
'            .cboProvP.BackColor = fColor
'            .cboKotaP.BackColor = fColor
'            .cboKecP.BackColor = fColor
'            .cboKelP.BackColor = fColor
            
            .cboTitleS.BackColor = fColor
           ' txt(7).BackColor = fColor
            .cboAgamaS.BackColor = fColor
            .cboPekS.BackColor = fColor
            .cboHobiS.BackColor = fColor
            
            .cboProvS.BackColor = fColor
            .cboKotaS.BackColor = fColor
            .cboKecS.BackColor = fColor
            .cboKelS.BackColor = fColor
            .DtLahir.BackColor = fColor
'            .DtDdate.BackColor = fColor
'            .DtPolish0.BackColor = fColor
'            .DtPolish1.BackColor = fColor
            .DtLahirS.BackColor = fColor
            .ChSurat.ForeColor = fColorz
'            .chPolish.ForeColor = fColorz
'            .chPolish.value = 0
            .ChSurat.value = 0
            .DtLahir.text = "__/__/____"
'            .DtDdate.text = "__/__/____"
'            .DtPolish0.text = "__/__/____"
'            .DtPolish1.text = "__/__/____"
            .DtLahirS.text = "__/__/____"
            .cmdSave.Visible = True
            .cmdCreate.Visible = False
            .cmdAdd.Visible = True
            .cmdAdd.Enabled = True
            .cmdCreate.Enabled = False
            .cmdEdit.Enabled = False
            .cmdSave.Enabled = False
            .cmdDelete.Enabled = False
            .cmdCancel.Enabled = False
            .CmdOpen.Enabled = False
            .cmdExit.Enabled = True
            
        Case "load_update"
            .Txt(1).BackColor = tColor
            For i = 2 To 43
                .Txt(i).BackColor = fColor
            Next i
'            .DtDdate.BackColor = fColor
            'SID .DtDOTam.BackColor = fColor
'            .DtPolish0.BackColor = fColor
'            .DtPolish1.BackColor = fColor
'            .chPolish.ForeColor = fColorz
            .ChSurat.ForeColor = fColorz
            'SID .optTipe(0).ForeColor = fColorz
            'SID .optTipe(1).ForeColor = fColorz
            
            .cmdSave.Visible = True
            .cmdCreate.Visible = False
            .cmdEdit.Visible = True
            .cmdAdd.Visible = False
            .cmdCreate.Enabled = False
            .cmdEdit.Enabled = True
            .cmdSave.Enabled = False
            .cmdDelete.Enabled = True
            .cmdCancel.Enabled = True
            .CmdOpen.Enabled = True
            .cmdExit.Enabled = True
            
        Case "add"
            Call Load_Cbo
            For i = 1 To 74
                Txt(i).BackColor = tColorx              'Kuning
                Select Case i
                    Case 2, 3, 15, 16, 30, 37, 38, 46, 47, 48, 49, 53, 61, 62, 71, 72, 73, 74
                        Txt(i).BackColor = tColor       'Putih
                    Case 1, 7, 11, 22, 23, 25, 29, 31, 32, 33, 34, 35, 36, 42, 57
                        Txt(i).BackColor = fColor       'Orange
                End Select
                Txt(i).text = ""
            Next i
            
            If Txt(31).text <> "" Then
'                .DtPolish0.BackColor = tColorx
'                .DtPolish0.text = "__/__/____"
'                .DtPolish1.BackColor = tColorx
'                .DtPolish1.text = "__/__/____"
            Else
'                .DtPolish0.BackColor = tColor
'                .DtPolish0.text = "__/__/____"
'                .DtPolish1.BackColor = tColor
'                .DtPolish1.text = "__/__/____"
            End If
            
            .cboTitle.BackColor = tColor
            .cboTipe.BackColor = tColor
            .cboAgama.BackColor = tColor
            .cboPek.BackColor = tColor
            .cboHobi.BackColor = tColor
            .cboProv.BackColor = tColor
            .cboKota.BackColor = tColor
            .cboKec.BackColor = tColor
            .cboKel.BackColor = tColor
'            .cboProvP.BackColor = tColor
'            .cboKotaP.BackColor = tColor
'            .cboKecP.BackColor = tColor
'            .cboKelP.BackColor = tColor
            .cboTitleS.BackColor = tColor
            .cboAgamaS.BackColor = tColor
            .cboPekS.BackColor = tColor
            .cboHobiS.BackColor = tColor
            .cboProvS.BackColor = tColor
            .cboKotaS.BackColor = tColor
            .cboKecS.BackColor = tColor
            .cboKelS.BackColor = tColor
            .DtLahir.BackColor = tColor
'            .DtDdate.BackColor = fColor
'            .DtPolish0.BackColor = tColor
'            .DtPolish1.BackColor = tColor
            .DtLahirS.BackColor = tColor
            .ChSurat.ForeColor = tColorz
'            .chPolish.ForeColor = tColorz
'            .chPolish.value = 0
            .ChSurat.value = 0
            .DtLahir.text = "__/__/____"
'            .DtDdate.text = "__/__/____"
'            .DtPolish0.text = "__/__/____"
'            .DtPolish1.text = "__/__/____"
            .DtLahirS.text = "__/__/____"
            
            .cmdSave.Visible = False
            .cmdCreate.Visible = True
            .cmdAdd.Enabled = False
            .cmdCreate.Enabled = True
            .cmdEdit.Enabled = False
            .cmdSave.Enabled = False
            .cmdCancel.Enabled = True
            .cmdDelete.Enabled = False
            .CmdOpen.Enabled = False
            .cmdExit.Enabled = False
            .cboTipe.SetFocus
            
        Case "edit"
            Txt(1).BackColor = fColor
            If xWAC = True Then Txt(1).BackColor = tColor
            For i = 2 To 4
                Txt(i).BackColor = tColorx
            Next i
            For i = 5 To 6
                Txt(i).BackColor = tColor
            Next i
            For i = 7 To 14
                Txt(i).BackColor = tColorx
                If i = 9 Then Txt(9).BackColor = tColor
            Next i
            For i = 17 To 21
                If i <> 18 Then Txt(i).BackColor = tColorx
            Next i
            If Left(UCase(Txt(14).text), 4) = "ZOTH" Then
                Txt(16).BackColor = tColor
            Else
                Txt(16).BackColor = fColor
            End If
            Txt(19).BackColor = fColor
            Txt(23).BackColor = tColor
            If Txt(24).text <> "" Then
                For i = 31 To 33
                    Txt(i).BackColor = tColorx
                Next i
                For i = 34 To 35
                    Txt(i).BackColor = tColor
                Next i
                For i = 36 To 43
                    Txt(i).BackColor = tColorx
                    If i = 40 Or i = 41 Then Txt(i).BackColor = tColor
                Next i
'                .DtPolish0.BackColor = tColorx
'                .DtPolish1.BackColor = tColorx
            Else
                For i = 31 To 36
                    Txt(i).BackColor = tColor
                Next i
                For i = 37 To 43
                    Txt(i).BackColor = tColorx
                    If i = 40 Or i = 41 Then Txt(i).BackColor = tColor
                Next i
'                .DtPolish0.BackColor = tColor
'                .DtPolish1.BackColor = tColor
            End If
            
'            .DtDdate.BackColor = tColor
            'SID .DtDOTam.BackColor = tColor
'            .chPolish.ForeColor = tColorz
            .ChSurat.ForeColor = tColorz
            'SID .optTipe(0).ForeColor = tColorz
            'SID .optTipe(1).ForeColor = tColorz
            
            .cmdSave.Visible = True
            .cmdCreate.Visible = False
            .cmdAdd.Enabled = False
            .cmdEdit.Enabled = False
            .cmdCreate.Enabled = False
            .cmdSave.Enabled = True
            .cmdCancel.Enabled = True
            .cmdDelete.Enabled = False
            .CmdOpen.Enabled = False
            .cmdExit.Enabled = False
    End Select
    End With
'txt(43).SetFocus
ErrNull:
End Sub

Private Sub Load_Cbo()

    cboKewarganegaraan.Clear
    cboKewarganegaraan.AddItem ("WNI")
    cboKewarganegaraan.AddItem ("WNA")
    Dim rsA As New ADODB.Recordset
    '===================================================
    '   TIPE
    '===================================================
    Set rsA = objSID.FetchDataSID("Select * from TipeCustomer With (NoLock) Order By Kode")
    With cboTipe
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Tipe")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   AGAMA
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Agama) as Agama from Agama With (NoLock)")
    With cboAgama
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Agama")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboAgamaS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Agama")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   PEKERJAAN
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Pekerjaan) as Pekerjaan from Pekerjaan With (NoLock) Order By Pekerjaan")
    With cboPek
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Pekerjaan")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboPekS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Pekerjaan")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   HOBI
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Hobi) as Hobi from Hobi With (NoLock) Order By Hobi")
    With cboHobi
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Hobi")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboHobiS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Hobi")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   OLAHRAGA
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Olahraga) as Olahraga from Olahraga With (NoLock) Order By Olahraga")
    With cboOlah
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Olahraga")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboOlahS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Olahraga")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   PROVINSI
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Distinct UPPER(Propinsi) AS Propinsi from TopologyDaerah With (NoLock) Order By Propinsi")
    With cboProv
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Propinsi")
                rsA.MoveNext
            Loop
        End If
        .AddItem "LAINNYA"
    End With

    With cboProvS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Propinsi")
                rsA.MoveNext
            Loop
        End If
        .AddItem "LAINNYA"
    End With
    
'    With cboProvP
'        .Clear
'        If rsA.RecordCount > 0 Then
'            rsA.MoveFirst
'            Do Until rsA.EOF
'                .AddItem rsA("Propinsi")
'                rsA.MoveNext
'            Loop
'        End If
'        .AddItem "LAINNYA"
'    End With
End Sub

Private Sub cboAgama_Click()
    cboPek.SetFocus
End Sub

Private Sub cboAgama_GotFocus()
    If cboAgama.BackColor = fColor Then
        cboAgama.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboAgama.Locked = False
    End If
End Sub

Private Sub cboPek_Click()
    Call cboPek_Change
End Sub

Private Sub cboPek_Change()
    If UCase(cboPek.text) = "LAINNYA" Then
        Txt(7).BackColor = tColor
        Txt(7).SetFocus
    Else
        Txt(7).BackColor = fColor
        Sendkeys ("{TAB}")
    End If
End Sub

Private Sub cboHobi_Click()
    Sendkeys ("{TAB}")
End Sub

Private Sub cboProv_Click()
    If cboProv.Tag = "" Then
        kliPr = True
        Call cboProv_Change
        cboKel.Clear
        cboKec.Clear
    End If
End Sub

Private Sub cboProv_Change()
    If kliPr = True Then
        If cboProv.text <> "" Then
            Dim rsX As New ADODB.Recordset, z As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProv.text)) & "' Order By Kota")
            With cboKota
                .Clear: z = 0: cboKota.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If z = 0 Then cboKota.Tag = rsX("Kota")
                        .AddItem rsX("Kota")
                        rsX.MoveNext
                        z = z + 1
                    Loop
                    kliKo = False
                    cboKota.text = cboKota.Tag
                    cboKota.Tag = ""
                    cboKota.SetFocus
                End If
            End With
        Else
            cboKota.Tag = "": cboKota.Clear
        End If
        If cboProv.text = "LAINNYA" Then
            FrAlm.Visible = True
            cmdBack.Visible = True
            Txt(63).SetFocus
        Else
            FrAlm.Visible = False
            cmdBack.Visible = False
        End If
        kliPr = False
    End If
End Sub

Private Sub cboKota_Click()
    If cboKota.Tag = "" Then
        kliKo = True
        Call cboKota_Change
    End If
End Sub

Private Sub cboKota_Change()
    If kliKo = True Then
        If cboKota.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKota.text)) & "' Order By Kecamatan")
            With cboKec
                .Clear: X = 0: cboKota.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKec.Tag = rsX("Kecamatan")
                        .AddItem rsX("Kecamatan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKec.text = cboKec.Tag
                    cboKec.Tag = ""
                    cboKec.SetFocus
                End If
            End With
        Else
            cboKec.Tag = "": cboKec.Clear
        End If
    kliKo = False
    End If
End Sub

Private Sub cboKec_Click()
    If cboKec.Tag = "" Then
        kliKc = True
        Call cboKec_Change
    End If
End Sub

Private Sub cboKec_Change()
    If kliKc = True Then
        If cboKec.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKec.text)) & "' Order By Kelurahan")
            With cboKel
                .Clear: X = 0: cboKel.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKel.Tag = rsX("Kelurahan")
                        .AddItem rsX("Kelurahan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKel.text = cboKel.Tag
                    cboKel.Tag = ""
                    cboKel.SetFocus
                End If
            End With
        Else
            cboKel.Tag = "": cboKel.Clear
        End If
    kliKc = False
    End If
End Sub

Private Sub cboKel_Click()
    If cboKel.Tag = "" Then
        kliKl = True
        Call cboKel_Change
    End If
End Sub

Private Sub cboKel_Change()
    If kliKl = True Then
        If cboKel.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah Where UPPER(Kelurahan) = '" & Trim(UCase(cboKel.text)) & "'")
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                Txt(11).text = rsX("KodePos")
                Txt(8).SetFocus
            Else
                Txt(11).text = ""
                Txt(8).SetFocus
            End If
        Else
            Txt(11).text = ""
        End If
        kliKl = False
    End If
End Sub

Private Sub cboPek_GotFocus()
    If cboPek.BackColor = fColor Then
        cboPek.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboPek.Locked = False
    End If
End Sub

Private Sub cboHobi_GotFocus()
    If cboHobi.BackColor = fColor Then
        cboHobi.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboHobi.Locked = False
    End If
End Sub

Private Sub cboTipe_GotFocus()
    If cboTipe.BackColor = fColor Then
        cboTipe.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboTipe.Locked = False
    End If
End Sub

Private Sub cboTitle_GotFocus()
    If cboTitle.BackColor = fColor Then
        cboTitle.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboTitle.Locked = False
    End If
End Sub

Private Sub cboProv_GotFocus()
    If cboProv.BackColor = fColor Then
        cboProv.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboProv.Locked = False
    End If
End Sub

Private Sub cboKota_GotFocus()
    If cboKota.BackColor = fColor Then
        cboKota.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboKota.Locked = False
    End If
End Sub

Private Sub cboKec_GotFocus()
    If cboKec.BackColor = fColor Then
        cboKec.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboKec.Locked = False
    End If
End Sub

Private Sub cboKel_GotFocus()
    If cboKel.BackColor = fColor Then
        cboKel.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboKel.Locked = False
    End If
End Sub
'
'Private Sub cboProvP_Click()
'    If cboProvP.Tag = "" Then
'        kliPrp = True
'        Call cboProvP_Change
'        cboKelP.Clear
'        cboKecP.Clear
'    End If
'End Sub
'
'Private Sub cboProvP_Change()
'    If kliPrp = True Then
'        If cboProvP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, z As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvP.Text)) & "' Order By Kota")
'            With cboKotaP
'                .Clear: z = 0: cboKotaP.Tag = ""
'                If rsX.RecordCount > 0 Then
'                    rsX.MoveFirst
'                    Do Until rsX.EOF
'                        If z = 0 Then cboKotaP.Tag = rsX("Kota")
'                        .AddItem rsX("Kota")
'                        rsX.MoveNext
'                        z = z + 1
'                    Loop
'                    kliKop = False
'                    cboKotaP.Text = cboKotaP.Tag
'                    cboKotaP.Tag = ""
'                    cboKotaP.SetFocus
'                End If
'            End With
'        Else
'            cboKotaP.Tag = "": cboKotaP.Clear
'        End If
'        If cboProvP.Text = "LAINNYA" Then
'            FrAlmP.Visible = True
'            cmdBackP.Visible = True
'            txt(67).SetFocus
'        Else
'            FrAlmP.Visible = False
'            cmdBackP.Visible = False
'        End If
'        kliPrp = False
'    End If
'End Sub
'
'Private Sub cboKotaP_Click()
'    If cboKotaP.Tag = "" Then
'        kliKop = True
'        Call cboKotaP_Change
'    End If
'End Sub
'
'Private Sub cboKotaP_Change()
'    If kliKop = True Then
'        If cboKotaP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, X As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaP.Text)) & "' Order By Kecamatan")
'            With cboKecP
'                .Clear: X = 0: cboKotaP.Tag = ""
'                If rsX.RecordCount > 0 Then
'                    rsX.MoveFirst
'                    Do Until rsX.EOF
'                        If X = 0 Then cboKecP.Tag = rsX("Kecamatan")
'                        .AddItem rsX("Kecamatan")
'                        rsX.MoveNext
'                        X = X + 1
'                    Loop
'                    cboKecP.Text = cboKecP.Tag
'                    cboKecP.Tag = ""
'                    cboKecP.SetFocus
'                End If
'            End With
'        Else
'            cboKecP.Tag = "": cboKecP.Clear
'        End If
'    kliKop = False
'    End If
'End Sub
'
'Private Sub cboKecP_Click()
'    If cboKecP.Tag = "" Then
'        kliKcp = True
'        Call cboKecP_Change
'    End If
'End Sub

'Private Sub cboKecP_Change()
'    If kliKcp = True Then
'        If cboKecP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, X As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecP.Text)) & "' Order By Kelurahan")
'            With cboKelP
'                .Clear: X = 0: cboKelP.Tag = ""
'                If rsX.RecordCount > 0 Then
'                    rsX.MoveFirst
'                    Do Until rsX.EOF
'                        If X = 0 Then cboKelP.Tag = rsX("Kelurahan")
'                        .AddItem rsX("Kelurahan")
'                        rsX.MoveNext
'                        X = X + 1
'                    Loop
'                    cboKelP.Text = cboKelP.Tag
'                    cboKelP.Tag = ""
'                    cboKelP.SetFocus
'                End If
'            End With
'        Else
'            cboKelP.Tag = "": cboKelP.Clear
'        End If
'    kliKcp = False
'    End If
'End Sub

'Private Sub cboKelP_Click()
'    If cboKelP.Tag = "" Then
'        kliKlp = True
'        Call cboKelP_Change
'    End If
'End Sub
'
'Private Sub cboKelP_Change()
'    If kliKlp = True Then
'        If cboKelP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, X As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah Where UPPER(Kelurahan) = '" & Trim(UCase(cboKelP.Text)) & "'")
'            If rsX.RecordCount > 0 Then
'                rsX.MoveFirst
'                txt(42).Text = rsX("KodePos")
'                txt(39).SetFocus
'            Else
'                txt(42).Text = ""
'                txt(39).SetFocus
'            End If
'        Else
'            txt(42).Text = ""
'        End If
'        kliKlp = False
'    End If
'End Sub
'
'Private Sub cboProvP_GotFocus()
'    If cboProvP.BackColor = fColor Then
'        cboProvP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboProvP.Locked = False
'    End If
'End Sub
'
'Private Sub cboKotaP_GotFocus()
'    If cboKotaP.BackColor = fColor Then
'        cboKotaP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboKotaP.Locked = False
'    End If
'End Sub
'
'Private Sub cboKecP_GotFocus()
'    If cboKecP.BackColor = fColor Then
'        cboKecP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboKecP.Locked = False
'    End If
'End Sub
'
'Private Sub cboKelP_GotFocus()
'    If cboKelP.BackColor = fColor Then
'        cboKelP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboKelP.Locked = False
'    End If
'End Sub

Private Sub cboHobiS_Click()
    Sendkeys ("{TAB}")
End Sub

Private Sub cboProvS_Click()
    If cboProvS.Tag = "" Then
        kliPrs = True
        Call cboProvS_Change
        cboKelS.Clear
        cboKecS.Clear
    End If
End Sub

Private Sub cboProvS_Change()
    If kliPrs = True Then
        If cboProvS.text <> "" Then
            Dim rsX As New ADODB.Recordset, z As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvS.text)) & "' Order By Kota")
            With cboKotaS
                .Clear: z = 0: cboKotaS.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If z = 0 Then cboKotaS.Tag = rsX("Kota")
                        .AddItem rsX("Kota")
                        rsX.MoveNext
                        z = z + 1
                    Loop
                    kliKos = False
                    cboKotaS.text = cboKotaS.Tag
                    cboKotaS.Tag = ""
                    cboKotaS.SetFocus
                End If
            End With
        Else
            cboKotaS.Tag = "": cboKotaS.Clear
        End If
        If cboProvS.text = "LAINNYA" Then
            FrAlmS.Visible = True
            cmdBackS.Visible = True
            Txt(67).SetFocus
        Else
            FrAlmS.Visible = False
            cmdBackS.Visible = False
        End If
        kliPrp = False
    End If
End Sub

Private Sub cboKotaS_Click()
    If cboKotaS.Tag = "" Then
        kliKos = True
        Call cboKotaS_Change
    End If
End Sub

Private Sub cboKotaS_Change()
    If kliKos = True Then
        If cboKotaS.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaS.text)) & "' Order By Kecamatan")
            With cboKecS
                .Clear: X = 0: cboKotaS.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKecS.Tag = rsX("Kecamatan")
                        .AddItem rsX("Kecamatan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKecS.text = cboKecS.Tag
                    cboKecS.Tag = ""
                    cboKecS.SetFocus
                End If
            End With
        Else
            cboKecS.Tag = "": cboKecS.Clear
        End If
    kliKos = False
    End If
End Sub

Private Sub cboKecS_Click()
    If cboKecS.Tag = "" Then
        kliKcs = True
        Call cboKecS_Change
    End If
End Sub

Private Sub cboKecS_Change()
    If kliKcs = True Then
        If cboKecS.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecS.text)) & "' Order By Kelurahan")
            With cboKelS
                .Clear: X = 0: cboKelS.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKelS.Tag = rsX("Kelurahan")
                        .AddItem rsX("Kelurahan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKelS.text = cboKelS.Tag
                    cboKelS.Tag = ""
                    cboKelS.SetFocus
                End If
            End With
        Else
            cboKelS.Tag = "": cboKelS.Clear
        End If
    kliKcs = False
    End If
End Sub

Private Sub cboKelS_Click()
    If cboKelS.Tag = "" Then
        kliKls = True
        Call cboKelS_Change
    End If
End Sub

Private Sub cboKelS_Change()
    If kliKls = True Then
        If cboKelS.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah Where UPPER(Kelurahan) = '" & Trim(UCase(cboKelS.text)) & "'")
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                Txt(57).text = rsX("KodePos")
                Txt(54).SetFocus
            Else
                Txt(57).text = ""
                Txt(54).SetFocus
            End If
        Else
            Txt(57).text = ""
        End If
        kliKls = False
    End If
End Sub
'
'Private Sub LoadCbo_KotaP()
'    Dim rsX As New ADODB.Recordset, z As Integer
'    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvP.Text)) & "' Order By Kota")
'    With cboKotaP
'        .Clear: z = 0: cboKotaP.Tag = ""
'        If rsX.RecordCount > 0 Then
'            rsX.MoveFirst
'            Do Until rsX.EOF
'                If z = 0 Then cboKotaP.Tag = rsX("Kota")
'                .AddItem rsX("Kota")
'                rsX.MoveNext
'                z = z + 1
'            Loop
'        End If
'    End With
'End Sub
'
'Private Sub LoadCbo_KecP()
'    Dim rsX As New ADODB.Recordset, z As Integer
'    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaP.Text)) & "' Order By Kecamatan")
'    With cboKecP
'        .Clear: z = 0: cboKecP.Tag = ""
'        If rsX.RecordCount > 0 Then
'            rsX.MoveFirst
'            Do Until rsX.EOF
'                If z = 0 Then cboKecP.Tag = rsX("Kecamatan")
'                .AddItem rsX("Kecamatan")
'                rsX.MoveNext
'                z = z + 1
'            Loop
'        End If
'    End With
'End Sub
'
'Private Sub LoadCbo_KelP()
'    Dim rsX As New ADODB.Recordset, z As Integer
'    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecP.Text)) & "' Order By Kelurahan")
'    With cboKelP
'        .Clear: z = 0: cboKelP.Tag = ""
'        If rsX.RecordCount > 0 Then
'            rsX.MoveFirst
'            Do Until rsX.EOF
'                If z = 0 Then cboKelP.Tag = rsX("Kelurahan")
'                .AddItem rsX("Kelurahan")
'                rsX.MoveNext
'                z = z + 1
'            Loop
'        End If
'    End With
'End Sub

Private Sub LoadCbo_KotaS()
    Dim rsX As New ADODB.Recordset, z As Integer
    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvS.text)) & "' Order By Kota")
    With cboKotaS
        .Clear: z = 0: cboKotaS.Tag = ""
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Do Until rsX.EOF
                If z = 0 Then cboKotaS.Tag = rsX("Kota")
                .AddItem rsX("Kota")
                rsX.MoveNext
                z = z + 1
            Loop
        End If
    End With
End Sub

Private Sub LoadCbo_KecS()
    Dim rsX As New ADODB.Recordset, z As Integer
    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaS.text)) & "' Order By Kecamatan")
    With cboKecS
        .Clear: z = 0: cboKecS.Tag = ""
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Do Until rsX.EOF
                If z = 0 Then cboKecS.Tag = rsX("Kecamatan")
                .AddItem rsX("Kecamatan")
                rsX.MoveNext
                z = z + 1
            Loop
        End If
    End With
End Sub

Private Sub LoadCbo_KelS()
    Dim rsX As New ADODB.Recordset, z As Integer
    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecS.text)) & "' Order By Kelurahan")
    With cboKelS
        .Clear: z = 0: cboKelS.Tag = ""
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Do Until rsX.EOF
                If z = 0 Then cboKelS.Tag = rsX("Kelurahan")
                .AddItem rsX("Kelurahan")
                rsX.MoveNext
                z = z + 1
            Loop
        End If
    End With
End Sub

Private Sub cboTitleS_GotFocus()
    If cboTitleS.BackColor = fColor Then
        cboTitleS.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboTitleS.Locked = False
    End If
End Sub

Private Sub cmdBack_Click()
    FrAlm.Visible = False
    cmdBack.Visible = False
    Call Load_Cbo
End Sub

'
'Private Sub isSavingData()
'    If IsValid = True Then
'        If objP.isExist_ItemKodeLgn(txt(1).Text, txt(19).Text) = False Then
'            '-----------------------------------------------------------------------------------------------------------------
'            '   INSERT CUSTDAT
'            '-----------------------------------------------------------------------------------------------------------------
''            Call objP.AddNew_MasterPelanggan(txt(1).text, txt(2).text, txtnpwp.text, txt(4).text, txt(5).text, _
''                                         txt(6).text, txt(7).text, txt(8).text, txt(9).text, IIf(optTipe(0).value = True, 1, 0), _
''                                         (txt(10).text & "-" & txt(11).text & (IIf(txt(12) <> "", "-" & txt(12).text, ""))), _
''                                         txt(13).text, txt(14).text, IIf(Left(txt(14).text, 4) = "ZOTH", txt(16).text, ""), txt(17).text, txt(19).text, txt(20).text, _
''                                         (IIf(DtDOTam.text = "__/__/____", "01/01/3000", DtDOTam.value)), _
''                                         (IIf(DtDdate.text = "__/__/____", "01/01/3000", DtDdate.value)), _
''                                         txt(21).text, txt(23).text, chPolish.value, txt(31).text, txt(32).text, _
''                                         txt(33).text, txt(34).text, txt(35).text, txt(36).text, _
''                                         (IIf(DtPolish0.text = "__/__/____", "01/01/3000", DtPolish0.value)), _
''                                         (IIf(DtPolish1.text = "__/__/____", "01/01/3000", DtPolish1.value)), _
''                                         ChSurat.value, txt(37).text, txt(38).text, txt(39).text, _
''                                         txt(40).text, txt(41).text, txt(42).text, txt(43).text)
'        Else
'            '-----------------------------------------------------------------------------------------------------------------
'            '   UPDATE CUSTDAT
'            '-----------------------------------------------------------------------------------------------------------------
'            Call objP.Update_MasterPelanggan(txt(1).Text, txt(2).Text, txtnpwp.Text, txt(4).Text, txt(5).Text, _
'                                         txt(6).Text, txt(7).Text, txt(8).Text, txt(9).Text, IIf(optTipe(0).value = True, 1, 0), _
'                                         (txt(10).Text & "-" & txt(11).Text & (IIf(txt(12) <> "", "-" & txt(12).Text, ""))), _
'                                         txt(13).Text, txt(14).Text, IIf(Left(txt(14).Text, 4) = "ZOTH", txt(16).Text, ""), txt(17).Text, txt(19).Text, txt(20).Text, _
'                                         (IIf(DtDOTam.Text = "__/__/____", "01/01/3000", DtDOTam.value)), _
'                                         (IIf(DtDdate.Text = "__/__/____", "01/01/3000", DtDdate.value)), _
'                                         txt(21).Text, txt(23).Text, chPolish.value, txt(31).Text, txt(32).Text, _
'                                         txt(33).Text, txt(34).Text, txt(35).Text, txt(36).Text, _
'                                         (IIf(DtPolish0.Text = "__/__/____", "01/01/3000", DtPolish0.value)), _
'                                         (IIf(DtPolish1.Text = "__/__/____", "01/01/3000", DtPolish1.value)), _
'                                         ChSurat.value, txt(37).Text, txt(38).Text, txt(39).Text, _
'                                         txt(40).Text, txt(41).Text, txt(42).Text, txt(43).Text)
'            '-----------------------------------------------------------------------------------------------------------------
'            '   UPDATE JOBDTA
'            '-----------------------------------------------------------------------------------------------------------------
'            Call objWO.Update_DataJobDTA(txt(2).Text, txtnpwp.Text, txt(4).Text, txt(5).Text, txt(6).Text, txt(7).Text, txt(8).Text, _
'                                        (txt(10).Text & "-" & txt(11).Text & (IIf(txt(12) <> "", "-" & txt(12).Text, ""))), _
'                                        txt(13).Text, txt(14).Text, txt(17).Text, txt(20).Text, txt(19).Text)
'        End If
'        ft.ShowForm = "ARMST"
'        For C = 1 To 9
'            ft.DataPelanggan = txt(C).Text
'        Next C
'        ft.DataPelanggan = IIf(optTipe(0).value = True, "0", "1")
'        ft.Show vbModal
'        Call SetControlForm("load_update")
'    Else
'        With msg
'            .PassingMsg = "Data Master Pelanggan Belum Lengkap !" & Chr(13) & nValid
'            .passingTipe = 2
'            .Show vbModal
'        End With
'    End If
'End Sub

Private Sub BindToControl_MasterAcc(rs As ADODB.Recordset)
    For i = 1 To 9
        Txt(i).text = ""
    Next i
    For i = 0 To 3
        Txt(i + 1).text = NulSID(rs.Fields(i).value)
    Next i
    Txt(6).text = IIf(Left(Replace(Replace(NulSID(rs.Fields(4).value), "-", ""), " ", ""), 2) = "08", "", Replace(Replace(NulSID(rs.Fields(4).value), "-", ""), " ", ""))
    Txt(7).text = IIf(Left(Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""), 2) <> "08", "", Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""))
    Txt(8).text = IIf(Left(Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""), 2) = "08", "", Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""))
    Txt(10).text = ""
    Txt(11).text = ""
    Txt(12).text = ""
    For i = 13 To 20
        Txt(i).text = ""
    Next i
    DtDOTam.text = "__/__/____"
    DtDdate.text = "__/__/____"
    For i = 21 To 29
        Txt(i).text = ""
    Next i
    chPolish.value = 0
    For i = 31 To 36
        Txt(i).text = ""
    Next i
    DtDOTam.text = "__/__/____"
    DtDdate.text = "__/__/____"
    ChSurat.value = 0
    For i = 37 To 43
        Txt(i).text = ""
    Next i
End Sub

Private Sub BindToControl_Model(rs As ADODB.Recordset)
    Txt(21).text = rs.Fields(0).value
    Txt(22).text = rs.Fields(1).value
    Txt(23).text = rs.Fields(2).value
End Sub

Private Sub BindToControl_Dealer(rs As ADODB.Recordset)
    Txt(28).text = rs.Fields(0).value
    Txt(29).text = rs.Fields(1).value
End Sub

Private Sub BindToControl_Warna(rs As ADODB.Recordset)
    Txt(24).text = rs.Fields(0).value
    Txt(25).text = rs.Fields(1).value
End Sub

Private Sub BindToControl_Asuransi(rs As ADODB.Recordset)
    Txt(30).text = NulSID(rs.Fields(0).value)
    Txt(31).text = NulSID(rs.Fields(1).value)
    Txt(32).text = NulSID(rs.Fields(2).value)
    Txt(33).text = NulSID(rs.Fields(3).value)
    Txt(34).text = NulSID(rs.Fields(4).value)
    Txt(35).text = NulSID(rs.Fields(5).value)
    Txt(36).text = NulSID(rs.Fields(6).value)
End Sub

Private Sub BindToControl_Master(rs As ADODB.Recordset, isJustCust As Boolean)
    For i = 0 To 8
        Txt(i + 1).text = rs.Fields(i).value
    Next i
    'SID optTipe(0).value = IIf(CInt(rs.Fields(9).value) = -1, True, False)
    'SID optTipe(1).value = IIf(CInt(rs.Fields(9).value) = 0, True, False)
    If isJustCust = False Then
        Call pecahNoPol(Trim(rs.Fields(10).value), nDpn, nTgh, nBlk)
        Txt(10).text = nDpn
        Txt(11).text = nTgh
        Txt(12).text = nBlk
        For i = 13 To 20
            Txt(i).text = rs.Fields(i - 2).value
        Next i
        If Left(UCase(Txt(21).text), 4) = "ZOTH" Then
            Txt(23).text = rs.Fields(47).value
        End If
        'SID DtDOTam.Text = IIf(Trim(rs.Fields(19).value) = "01/01/3000", "__/__/____", rs.Fields(19).value)
        'SID DtDdate.Text = IIf(Trim(rs.Fields(20).value) = "01/01/3000", "__/__/____", rs.Fields(20).value)
        For i = 21 To 29
            Txt(i).text = NulSID(rs.Fields(i).value)
        Next i
        chPolish.value = IIf(rs.Fields(30).value = True, 1, 0)
        For i = 31 To 36
            Txt(i).text = rs.Fields(i).value
        Next i
        DtPolish0.text = IIf(Trim(rs.Fields(37).value) = "01/01/3000", "__/__/____", rs.Fields(37).value)
        DtPolish1.text = IIf(Trim(rs.Fields(38).value) = "01/01/3000", "__/__/____", rs.Fields(38).value)
        ChSurat.value = IIf(rs.Fields(39).value = True, 1, 0)
        For i = 37 To 43
            Txt(i).text = rs.Fields(i + 3).value
        Next i
    Else
        Txt(10).text = ""
        Txt(11).text = ""
        Txt(12).text = ""
        For i = 13 To 20
            Txt(i).text = ""
        Next i
        'SID DtDOTam.Text = "__/__/____"
        'SID DtDdate.Text = "__/__/____"
        For i = 21 To 29
            Txt(i).text = ""
        Next i
        chPolish.value = 0
        For i = 31 To 36
            Txt(i).text = ""
        Next i
        'SID DtDOTam.Text = "__/__/____"
        DtDdate.text = "__/__/____"
        ChSurat.value = 0
        For i = 37 To 43
            Txt(i).text = ""
        Next i
    End If
End Sub

'Private Sub chPolish_Click()
'    If chPolish.value = 1 Then
'        If txt(4).Text = "" Then Exit Sub
'        For i = 37 To 47
'            If i = 37 Then
'                txt(i).Text = txt(4).Text & " " & txt(5).Text & IIf(txt(6).Text <> "", " " & txt(6).Text, "")
'            Else
'                txt(i).Text = txt(i - 31).Text
'            End If
'            txt(i).BackColor = fColor
'            If i = 38 Then
'                txt(i).Text = ""
'                txt(i).BackColor = tColor
'            End If
'        Next i
'        cboProvP.BackColor = fColor
'        cboKotaP.BackColor = fColor
'        cboKecP.BackColor = fColor
'        cboKelP.BackColor = fColor
'        If cboProv.Text = "LAINNYA" Or FrAlm.Visible = True Then
'            FrAlmP.Visible = True
'            cmdBackP.Visible = True
'            cmdBackP.Enabled = False
'            For i = 67 To 70
'                txt(i).Text = txt(i - 4).Text
'            Next i
'        Else
'            FrAlmP.Visible = False
'            cmdBackP.Visible = False
'
'            If cboProv.Text = "" Then GoTo GoKota
'            cboProvP.Tag = "ari"
'            cboProvP.Text = cboProv.Text
'            Call LoadCbo_KotaP
'            cboProvP.Tag = ""
'GoKota:
'            If cboKota.Text = "" Then GoTo GoKec
'            cboKotaP.Tag = "ari"
'            cboKotaP.Text = cboKota.Text
'            Call LoadCbo_KecP
'            cboKotaP.Tag = ""
'GoKec:
'            If cboKec.Text = "" Then GoTo GoKel
'            cboKecP.Tag = "ari"
'            cboKecP.Text = cboKec.Text
'            Call LoadCbo_KelP
'            cboKecP.Tag = ""
'GoKel:
'            If cboKel.Text = "" Then GoTo GoEnd
'            cboKelP.Tag = "ari"
'            cboKelP.Text = cboKel.Text
'            cboKelP.Tag = ""
'GoEnd:
'        End If
'        cboProvP.BackColor = fColor
'        cboKotaP.BackColor = fColor
'        cboKecP.BackColor = fColor
'        cboKelP.BackColor = fColor
'        txt(38).SetFocus
'    Else
'        For i = 37 To 47
'            txt(i).Text = ""
'            txt(i).BackColor = tColor
'            If i = 39 Then txt(i).BackColor = tColorx
'            If i = 40 Then txt(i).BackColor = tColorx
'            If i = 41 Then txt(i).BackColor = tColorx
'            If i = 43 Then txt(i).BackColor = tColorx
'            If i = 44 Then txt(i).BackColor = tColorx
'            If i = 45 Then txt(i).BackColor = tColorx
'            If i = 42 Then txt(i).BackColor = fColor
'        Next i
'        For i = 67 To 70
'            txt(i).Text = ""
'            txt(i).BackColor = tColor
'        Next i
'        cboProvP.BackColor = tColor
'        cboKotaP.BackColor = tColor
'        cboKecP.BackColor = tColor
'        cboKelP.BackColor = tColor
'        FrAlmP.Visible = False
'        cmdBackP.Visible = False
'        cboKotaP.Clear
'        cboKecP.Clear
'        cboKelP.Clear
'    End If
'End Sub

Private Sub chSurat_Click()
    If ChSurat.value = 1 Then
        For i = 37 To 43
            Txt(i).text = Txt(i - 35).text
        Next i
    Else
        For i = 37 To 43
            Txt(i).text = ""
        Next i
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 63, 64, 65, 66, 11
                
                If gMainKeyCari = "" Then gMainKeyCari = Txt(Index)
                gSearchOption = "TOPOLOGY"
                If frmCari.Visible Then
                    frmCari.Hide
                End If
                frmCari.Show (vbModal)
            Case 20
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDAGAMA"
                frmCari.Show (vbModal)
            Case 7
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDPEKERJAAN"
                frmCari.Show (vbModal)
            Case 21
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDHOBI"
                frmCari.Show (vbModal)
            Case 22
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDOLAHRAGA"
                frmCari.Show (vbModal)
            Case 1          'Custdat
                onFind = True
                onF12 = True
                ft.ShowForm = "SID"
                ft.Show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    If cmdCreate.Enabled = True Then
                        Call BindToControl_Master(rsFind, True)
                        cmdEdit.Enabled = False
                        cmdDelete.Enabled = False
                        CmdOpen.Enabled = False
                    Else
                        Call BindToControl_Master(rsFind, False)
                        cmdEdit.Enabled = True
                        cmdDelete.Enabled = True
                        CmdOpen.Enabled = True
                    End If
                    cmdAdd.Visible = False
                    cmdCancel.Enabled = True
                Else
                    If ft.IsActi = "CLOSE" Then
                        Txt(1).SetFocus
                    ElseIf ft.IsActi = "ADD" Then
                        Call CmdAdd_Click
                    End If
                End If
                onF12 = False
            Case 21          '" & DB_bengkel & "..Model
                onFind = True
                ft.ShowForm = "Model"
                ft.Show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Model(rsFind)
                    Txt(24).SetFocus
                Else
                    Txt(21).SetFocus
                End If
            Case 24          '" & DB_spk & "..InWarnaExt
                onFind = True
                ft.ShowForm = "Warna"
                ft.Show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Warna(rsFind)
                    Txt(26).SetFocus
                Else
                    Txt(24).SetFocus
                End If
            Case 28         '" & DB_bengkel & "..T_Dealer
                onFind = True
                ft.ShowForm = "Dealer"
                ft.Show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Dealer(rsFind)
                    Txt(30).SetFocus
                Else
                    Txt(28).SetFocus
                End If
            Case 30         'ACC11..ARMST(ASS)
                onFind = True
                ft.ShowForm = "AsuransiBp"
                ft.Show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Asuransi(rsFind)
                    Txt(37).SetFocus
                Else
                    Txt(30).SetFocus
                End If
        End Select
        onFind = False
    ElseIf KeyCode = 13 Then
        If Index = 11 Then
            If NulSID(Txt(Index)) <> "" Then
                Dim rsTopology As ADODB.Recordset
                Set rsTopology = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE KODEPOS = '" & NulSID(Txt(11)) & "'")
                If Not rsTopology.EOF Then
                    rsTopology.MoveFirst
                    Txt(63) = NulSID(rsTopology!Propinsi)
                    Txt(64) = NulSID(rsTopology!Kota) 'nulsid(DGrid.Columns(1))
                    Txt(65) = NulSID(rsTopology!kecamatan) 'nulsid(DGrid.Columns(2))
                    Txt(66) = NulSID(rsTopology!kelurahan) 'nulsid(DGrid.Columns(3))
                    Txt(11) = NulSID(rsTopology!KodePos)
                Else
                    MsgBox "Data kodepos tidak terdaftar.", vbInformation
                    Txt(63) = ""
                    Txt(64) = "" 'nulsid(DGrid.Columns(1))
                    Txt(65) = "" 'nulsid(DGrid.Columns(2))
                    Txt(66) = "" 'nulsid(DGrid.Columns(3))
                    Txt(11) = "" 'nulsid(DGrid.Columns(4))
                End If
            End If
        Else
            Call txt_Validate(Index, False)
        End If
    Else
        Select Case Index
            Case 2
                If cboTipe.text = "" Then
                    MsgBox "Tipe Customer Belum dipilih!"
                    cboTipe.SetFocus
                End If
'            Case 63, 64, 65, 66, 20, 7, 21, 22
'                KeyCode = 0
                 
        End Select
        
    End If
End Sub

'=====================================================================================
'============================= SETTING CONTROL =======================================
'=====================================================================================
Private Sub txt_GotFocus(Index As Integer)
    'Label1(3).Caption = Index: DoEvents
    Dim DiTab As Boolean
    DiTab = False
    If Txt(Index).BackColor = fColor Then
        Txt(Index).Locked = True
        Sendkeys ("{TAB}")
    Else
        Txt(Index).Locked = False
        If Index = 9 Then
            If Txt(8).text = "" Then Txt(8).SetFocus
        End If
    End If
    If gMainKeyCari <> "" Then
        Txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        DiTab = True
        Sendkeys ("{TAB}")
        'txt(Index).SetFocus
        'SendKeys ("{TAB}")
    End If
    
    If Index = 7 Then
        If NulSID(UCase(Txt(Index))) = "LAINNYA" Then
            'txt(23).SetFocus
            Txt(23).BackColor = &HFFFFFF
        End If
    End If
'    If DiTab Then
'        SendKeys ("{TAB}")
'    Else
'        SendKeys ("{HOME}+{END}")
'    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 19 Then
        Select Case KeyAscii
            Case 65 To 90
                'OK
            Case 97 To 122
                'OK
            Case 48 To 57
                'OK
            Case 8
                'OK
            Case 13
                'OK
            Case Else
                KeyAscii = 0
        End Select
    End If
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        If onFind = False Then
            If Index = 1 Or Index = 24 Or Index = 21 Or Index = 23 Then
                Select Case Index
                    Case 1
                        If Txt(1).text = "" Then Exit Sub
                        If onF12 = True Then Exit Sub
                        If cmdSave.Enabled = True Then Exit Sub
                        Dim rsExist As ADODB.Recordset
                        Set rsExist = objP.FetchListAllPelanggan("Where KodeLg = '" & Txt(1).text & "'")
                        If rsExist.RecordCount = 0 Then
                            Dim rsExistAcc As ADODB.Recordset
                            Set rsExistAcc = objP.FetchPelangganACC("Where KodeLgn = '" & Txt(1).text & "'")
                            If rsExistAcc.RecordCount > 0 Then
                                With msg
                                    .PassingMsg = "Data Master Pelanggan Kode : " & Txt(1).text & " Sudah terdaftar di Accounting, Tampilkan Data ?"
                                    .passingTipe = 1
                                    .Show vbModal
                                    If .GetMsg = True Then
                                        If cmdAdd.Enabled = True Or cmdEdit.Enabled = True Then
                                            If cmdCreate.Enabled = True Then
                                                Call BindToControl_MasterAcc(rsExistAcc)
                                                cmdEdit.Enabled = False
                                                cmdDelete.Enabled = False
                                                CmdOpen.Enabled = False
                                            Else
                                                Call BindToControl_MasterAcc(rsExistAcc)
                                                CmdOpen.Enabled = True
                                                cmdEdit.Enabled = True
                                                cmdDelete.Enabled = True
                                            End If
                                        Else
                                            Call BindToControl_MasterAcc(rsExistAcc)
                                            cmdEdit.Enabled = False
                                            cmdDelete.Enabled = False
                                            CmdOpen.Enabled = False
                                        End If
                                        cmdAdd.Enabled = False
                                        cmdAdd.Visible = False
                                        cmdCancel.Enabled = True
                                    Else
                                        Txt(1).text = ""
                                        If cmdCreate.Enabled = True Or cmdSave.Enabled = True Then Exit Sub
                                        Call SetControlForm("load")
                                    End If
                                End With
                            Else
                                'If cmdCreate.Enabled = True Or cmdSave.Enabled = True Then Exit Sub
                                With msg
                                    .PassingMsg = "Data Master Pelanggan Baru ?"
                                    .passingTipe = 1
                                    .Show vbModal
                                    If .GetMsg = True Then
                                        Dim nkode As String
                                        nkode = Txt(1).text
                                        Call SetControlForm("add")
                                        Txt(1).text = nkode
                                    Else
                                        Txt(1).text = ""
                                    End If
                                End With
                            End If
                        Else
                            With msg
                                .PassingMsg = "Data Master Pelanggan Kode : " & Txt(1).text & " Sudah terdaftar, Tampilkan Data ?"
                                .passingTipe = 1
                                .Show vbModal
                                If .GetMsg = True Then
                                    If cmdAdd.Enabled = True Or cmdEdit.Enabled = True Then
                                        If cmdCreate.Enabled = True Then
                                            Call BindToControl_Master(rsExist, True)
                                            cmdEdit.Enabled = False
                                            cmdDelete.Enabled = False
                                            CmdOpen.Enabled = False
                                        Else
                                            Call BindToControl_Master(rsExist, False)
                                            CmdOpen.Enabled = True
                                            cmdEdit.Enabled = True
                                            cmdDelete.Enabled = True
                                        End If
                                    Else
                                        Call BindToControl_Master(rsExist, True)
                                        cmdEdit.Enabled = False
                                        cmdDelete.Enabled = False
                                        CmdOpen.Enabled = False
                                    End If
                                    cmdAdd.Enabled = False
                                    cmdAdd.Visible = False
                                    cmdCancel.Enabled = True
                                Else
                                    Txt(1).text = ""
                                    If cmdCreate.Enabled = True Or cmdSave.Enabled = True Then Exit Sub
                                    Call SetControlForm("load")
                                End If
                            End With
                            'cmdOpen.SetFocus
                        End If
'                    Case 21
'                        'Dim objM As New clsModel
'                        Set rsFind = objM.FetchListNamaModel(" Where MDL = '" & txt(21).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Model(rsFind)
'                        Else
'                            txt(21).Text = ""
'                            txt(22).Text = ""
'                            txt(23).Text = ""
'                        End If
'                    Case 24
'                        Dim objW As New clsWarna
'                        Set rsFind = objW.FetchListWarna(" Where KodeWarna = '" & txt(24).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Warna(rsFind)
'                        Else
'                            txt(24).Text = ""
'                            txt(25).Text = ""
'                        End If
'                    Case 28
'                        Set rsFind = objD.FetchListNamaDealer(" Where FS_KD_Dealer = '" & txt(21).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Dealer(rsFind)
'                        Else
'                            txt(28).Text = ""
'                            txt(29).Text = ""
'                        End If
'                    Case 30
'                        Dim objA As New clsAsuransi
'                        Set rsFind = objA.FetchListAsuransiBp(" Where Kode = '" & txt(30).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Asuransi(rsFind)
'                        Else
'                            txt(30).Text = ""
'                            txt(31).Text = ""
'                            txt(32).Text = ""
'                            txt(33).Text = ""
'                            txt(34).Text = ""
'                            txt(35).Text = ""
'                            txt(36).Text = ""
'                        End If
                End Select
            Else
                If Index = 66 Then
                    Txt(8).SetFocus
                End If
            End If
        End If
    Else
        If Txt(Index).BackColor = &HC0E0FF Then
            KeyAscii = 0
            Exit Sub
        End If
        If KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            Select Case Index
                Case 3, 9, 10, 12, 13, 14, 15, 19, 40, 41, 43, 44, 45, 46, 55, 56, 58, 59, 60, 61
                    If KeyAscii <= 57 And KeyAscii >= 48 Then
                        KeyAscii = KeyAscii
                    ElseIf KeyAscii = 46 Then
                        KeyAscii = KeyAscii
                    Else
                        KeyAscii = 0
                    End If
                Case 2
                    If Label5(0).Caption <> "Passport / KITAS" Then
                        If KeyAscii <= 57 And KeyAscii >= 48 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii = 46 Then
                            KeyAscii = KeyAscii
                        Else
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii <= 57 And KeyAscii >= 48 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii = 46 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii <= 90 And KeyAscii >= 65 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii <= 122 And KeyAscii >= 97 Then
                            KeyAscii = KeyAscii
                        Else
                            KeyAscii = 0
                        End If
                    End If
'                Case 63, 64, 65, 66, 20, 7, 21, 22
'                    KeyAscii = 0
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        End If
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
     
    If gMainKeyCari <> "" Then
        Txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        'SendKeys ("{TAB}")
    End If
End Sub

'Private Sub DtDotam_GotFocus()
'    If DtDOTam.BackColor = fColor Then SendKeys "{TAB}"
'End Sub

'Private Sub DtDdate_GotFocus()
'    If DtDdate.BackColor = fColor Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtPolish0_GotFocus()
'    If DtPolish0.BackColor = fColor Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtPolish1_GotFocus()
'    If DtPolish1.BackColor = fColor Then SendKeys "{TAB}"
'End Sub

'Private Sub optTipe_GotFocus(Index As Integer)
'    If optTipe(Index).ForeColor = fColorz Then SendKeys "{TAB}"
'End Sub

'Private Sub chPolish_GotFocus()
'    If chPolish.ForeColor = fColorz Then SendKeys "{TAB}"
'End Sub

Private Sub chSurat_GotFocus()
    If ChSurat.ForeColor = fColorz Then Sendkeys "{TAB}"
End Sub
'
'Private Sub DtDotam_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtDdate_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtPolish0_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

'Private Sub DtPolish1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub optTipe_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

'Private Sub chPolish_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub chSurat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cboTipe_Click()
    Call cboTipe_Change
End Sub

Private Sub cboTipe_Change()
    If cboTipe.text <> "" Then
        Dim rsZ As New ADODB.Recordset
        Set rsZ = objSID.FetchDataSID("Select b.Title From TipeCustomer a Inner Join Title b On a.Kode=b.Tipe Where Upper(a.Tipe) = '" & UCase(cboTipe.text) & "'")
        With cboTitle
            .Clear
            If rsZ.RecordCount > 0 Then
                rsZ.MoveFirst
                Do Until rsZ.EOF
                    .AddItem rsZ("Title")
                    rsZ.MoveNext
                Loop
            End If
        End With
        If UCase(cboTipe.text) = "PERUSAHAAN UMUM" Then
            Label5(0).Caption = "No. TDP"
            txtNpwp.BackColor = &HFFFFFF ' &H80FFFF
            Txt(2).BackColor = &HFFFFFF
        ElseIf UCase(cboTipe.text) = "PERORANGAN" Then
            Label5(0).Caption = "No. KTP"
            Txt(2).BackColor = &HFFFFFF ' &H80FFFF
            txtNpwp.BackColor = &HFFFFFF
        Else
            Label5(0).Caption = "No. TDP"
            txtNpwp.BackColor = &HFFFFFF ' &H80FFFF
            Txt(2).BackColor = &HFFFFFF
        End If
        If isEditData = False Then
            If UCase(cboTipe.text) = "PERUSAHAAN UMUM" Then
                Label5(0).Caption = "No. TDP"
                Txt(1).text = LastID(cboTipe.text)
            ElseIf UCase(cboTipe.text) = "PERORANGAN" Then
                Label5(0).Caption = "No. KTP"
                Txt(1).text = LastID(cboTipe.text)
            Else
                Label5(0).Caption = "No. TDP"
                Txt(1).text = LastID(cboTipe.text)
            End If
            
            If UCase(cboTipe.text) <> "PERORANGAN" And NulSID(cboTipe.text) <> "" Then
                Txt(18).Enabled = True
                Txt(19).Enabled = True
                Txt(18).BackColor = &H80FFFF
                Txt(19).BackColor = &H80FFFF
                
            Else
                Txt(18).text = ""
                Txt(19).text = ""
                Txt(18).Enabled = False
                Txt(19).Enabled = False
                Txt(18).BackColor = &HC0E0FF
                Txt(19).BackColor = &HC0E0FF
                
            End If
        End If
        Txt(2).SetFocus
        
    End If
End Sub

Private Sub cboTitle_Click()
    Call cboTitle_Change
End Sub

Private Sub cboTitle_Change()
    If cboTipe.text = "" Then
        MsgBox "Tipe Customer belum dipilih!"
        cboTipe.SetFocus
    End If
    If Txt(4).Enabled Then Txt(4).SetFocus
End Sub

'Private Sub ViewPassing()
'    If passBy = "" Then Exit Sub
'    Select Case passBy
'        Case "BOK"
'            If rsBy.Fields(6).value = "" Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'            For C = 1 To 8
'                txt(C).Text = rsBy.Fields(C + 5).value
'            Next C
'            Call pecahNoPol(Trim(rsBy.Fields(5).value), nDpn, nTgh, nBlk)
'            txt(10).Text = nDpn
'            txt(11).Text = nTgh
'            txt(12).Text = nBlk
'            txt(13).Text = Trim(rsBy.Fields(14).value)
'            txt(21).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(21, 13)
'            txt(24).Text = Trim(rsBy.Fields(16).value)
'            Call txt_KeyPress(24, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''            If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then      'SID
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(28).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(28, 13)
'            txt(23).Text = Trim(rsBy.Fields(22).value)
'            Call txt_KeyPress(23, 13)
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsFindIn As ADODB.Recordset
'                Set rsFindIn = objP.FetchListAllPelanggan(" Where c.KodeLg = '" & txt(1).Text & "' And c.NoRangka = '" & txt(19).Text & "'")
'                If rsFindIn.RecordCount > 0 Then
'                    Call cmdCancel_Click
'                    Call BindToControl_Master(rsFindIn, False)
'                    cmdEdit.Enabled = True
'                    cmdDelete.Enabled = True
'                    cmdOpen.Enabled = True
'                    cmdAdd.Visible = False
'                    cmdCancel.Enabled = False
'                End If
'            End If
'        Case "NEWBOK"
'            On Error GoTo ErNul
'            If rsBy.Fields(0).value = "" Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'            For C = 1 To 9
'                txt(C).Text = rsBy.Fields(C - 1).value
'            Next C
'            'optTipe(CInt(IIf(rsBy.Fields(9).value = "0", 1, 0))).value = True  'SID
'            Call pecahNoPol(Trim(rsBy.Fields(10).value), nDpn, nTgh, nBlk)
'            txt(10).Text = nDpn
'            txt(11).Text = nTgh
'            txt(12).Text = nBlk
'            txt(13).Text = Trim(rsBy.Fields(11).value)
'            txt(14).Text = Trim(rsBy.Fields(12).value)
'            Call txt_KeyPress(14, 13)
'            txt(24).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(17, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''SID         If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(21).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(21, 13)
'            txt(23).Text = Trim(rsBy.Fields(23).value)
'            Call txt_KeyPress(23, 13)
'            For C = 31 To 36
'                txt(C).Text = Trim(rsBy.Fields(C).value)
'            Next C
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(37).value) = "01/01/3000", "__/__/____", rsBy.Fields(37).value)
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(38).value) = "01/01/3000", "__/__/____", rsBy.Fields(38).value)
'            For C = 37 To 43
'                txt(C).Text = Trim(rsBy.Fields(C + 3).value)
'            Next C
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsFindInx As ADODB.Recordset
'                Set rsFindInx = objP.FetchListAllPelanggan(" Where c.KodeLg = '" & txt(1).Text & "' And c.NoRangka = '" & txt(19).Text & "'")
'                If rsFindInx.RecordCount > 0 Then
'                    Call cmdCancel_Click
'                    Call BindToControl_Master(rsFindInx, False)
'                    cmdEdit.Enabled = True
'                    cmdDelete.Enabled = True
'                    cmdOpen.Enabled = True
'                    cmdAdd.Visible = False
'                    cmdCancel.Enabled = False
'                End If
'            End If
'        Case "EST"
'            On Error GoTo ErNul
'            If rsBy.Fields(5).value = "" Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'            For C = 1 To 9
'                txt(C).Text = rsBy.Fields(C + 4).value
'            Next C
'            If rsBy.Fields(4).value <> "" Then
'                Call pecahNoPol(Trim(rsBy.Fields(4).value), nDpn, nTgh, nBlk)
'                txt(10).Text = nDpn
'                txt(11).Text = nTgh
'                txt(12).Text = nBlk
'            Else
'                txt(10).Text = ""
'                txt(11).Text = ""
'                txt(12).Text = ""
'            End If
'            txt(13).Text = Trim(rsBy.Fields(14).value)
'            txt(14).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(14, 13)
'            txt(17).Text = Trim(rsBy.Fields(16).value)
'            Call txt_KeyPress(17, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''SID         If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(21).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(21, 13)
'            txt(23).Text = Trim(rsBy.Fields(22).value)
'            Call txt_KeyPress(23, 13)
'            For C = 31 To 36
'                txt(C).Text = Trim(rsBy.Fields(C - 8).value)
'            Next C
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(29).value) = "01/01/3000", "__/__/____", rsBy.Fields(29).value)
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(30).value) = "01/01/3000", "__/__/____", rsBy.Fields(30).value)
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsExist As ADODB.Recordset
'                Set rsExist = objP.FetchListAllPelanggan("Where C.KodeLg = '" & txt(1).Text & "' And NoRangka = '" & txt(19).Text & "'")
'                If rsExist.RecordCount > 0 Then
'                    Call BindToControl_Master(rsExist, False)
'                End If
'            End If
'        Case "WO"
'        Case "NewEST"
'
'            If rsBy.EOF = True Then
'                Call SetControlForm("add")
'                Exit Sub
'            Else
'                txt(1).Text = rsBy.Fields(0).value
'                If rsBy.Fields(0).value = "" Then
'                    Call SetControlForm("add")
'                Else
'                    Call SetControlForm("edit")
'                End If
'            End If
'            For C = 1 To 8
'                txt(C + 1).Text = rsBy.Fields(C).value
'            Next C
'            If rsBy.Fields(9).value = False Then
'                'SID Me.optTipe(1).value = True
'            Else
'                'SID Me.optTipe(0).value = True
'            End If
'            If rsBy.Fields(10).value <> "" Then
'                Call pecahNoPol(Trim(rsBy.Fields(10).value), nDpn, nTgh, nBlk)
'                txt(10).Text = nDpn
'                txt(11).Text = nTgh
'                txt(12).Text = nBlk
'            Else
'                txt(10).Text = ""
'                txt(11).Text = ""
'                txt(12).Text = ""
'            End If
'            txt(13).Text = Trim(rsBy.Fields(11).value)
'            txt(14).Text = Trim(rsBy.Fields(12).value)
'            Call txt_KeyPress(14, 13)
'            txt(17).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(17, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''SID         If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(21).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(21, 13)
'            txt(23).Text = Trim(rsBy.Fields(23).value)
'            Call txt_KeyPress(23, 13)
'            chPolish.value = IIf(rsBy.Fields(30).value = True, 1, 0)
'            For C = 31 To 36
'                txt(C).Text = Trim(rsBy.Fields(C).value)
'            Next C
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(37).value) = "01/01/3000", "__/__/____", rsBy.Fields(37).value)
'            DtPolish1.Text = IIf(Trim(rsBy.Fields(38).value) = "01/01/3000", "__/__/____", rsBy.Fields(38).value)
'            ChSurat.value = IIf(rsBy.Fields(39).value = True, 1, 0)
'            For C = 37 To 43
'                txt(C).Text = Trim(rsBy.Fields(C + 3).value)
'            Next C
'        Case "NEWBOKWAC"
'            On Error GoTo ErNul
'            'rsBy.Sort = "[" & rsBy.Fields(0).value & "]" & " Asc"
'            rsBy.Find "Pilih = 'TRUE'", , adSearchForward, 1
'            If rsBy.RecordCount = 0 Then Exit Sub
'            If rsBy.Fields(0).value = False Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'
'            txt(1).Text = ""
'            txt(2).Text = nulsid(rsBy("Nama"), "")
'            txtnpwp.Text = nulsid(rsBy("Alamat"), "")
'            txt(4).Text = nulsid(rsBy("Kota"), "")
'            txt(5).Text = ""
'            txt(6).Text = nulsid(rsBy("Telp"), "")
'            txt(7).Text = ""
'            txt(8).Text = ""
'            txt(9).Text = nulsid(rsBy("Email"), "")
'            'SID optTipe(0).value = False
'            'SID optTipe(1).value = False
'            Call pecahNoPol(Trim(rsBy("NoPolisi")), nDpn, nTgh, nBlk)
'            txt(10).Text = nDpn
'            txt(11).Text = nTgh
'            txt(12).Text = nBlk
'            txt(13).Text = Trim(nulsid(rsBy("ThMobil"), ""))
'            txt(14).Text = Trim(nulsid(rsBy("FullModel"), ""))
'            Call txt_KeyPress(14, 13)
'            txt(17).Text = ""
'            'Call txt_KeyPress(17, 13)
'            txt(18).Text = rsBy("Color")
'            txt(19).Text = Trim(nulsid(rsBy("NoRangka"), ""))
'            txt(20).Text = Trim(nulsid(rsBy("NoMesin"), ""))
'            'SID DtDOTam.Text = "__/__/____"
'            DtDdate.Text = "__/__/____"
'            txt(28).Text = Trim(nulsid(rsBy("KodeDealerWAC"), ""))
'            Call txt_KeyPress(28, 13)
'            'txt(23).Text = Trim(nulsid(rsBy("KodeAssBP"), ""))
''            Call txt_KeyPress(23, 13)
''            txt(31).Text = Trim(nulsid(rsBy("NamaPo"), ""))
''            txt(32).Text = Trim(nulsid(rsBy("AlamatPo"), ""))
''            txt(33).Text = Trim(nulsid(rsBy("KotaPo"), ""))
''            txt(34).Text = ""
''            txt(35).Text = Trim(nulsid(rsBy("TelpPo"), ""))
''            txt(36).Text = Trim(nulsid(rsBy("NomorPolish"), ""))
''            DtPolish0.Text = IIf(Trim(nulsid(rsBy("JangkaPolishAwal"), "01/01/3000")) = "01/01/3000", "__/__/____", rsBy("JangkaPolishAwal"))
''            DtPolish1.Text = IIf(Trim(nulsid(rsBy("JangkaPolishAkhir"), "01/01/3000")) = "01/01/3000", "__/__/____", rsBy("JangkaPolishAkhir"))
'            If xAs(0) <> "" Then
'                chPolish.value = 1
'            Else
'                chPolish.value = 0
'            End If
'            txt(23).Text = xAs(0)
'            Call txt_KeyPress(23, 13)
'            txt(31).Text = xAs(1)
'            txt(32).Text = xAs(2)
'            txt(33).Text = xAs(3)
'            txt(34).Text = ""
'            txt(35).Text = xAs(4)
'            txt(36).Text = xAs(5)
'            DtPolish0.Text = xAs(6)
'            DtPolish1.Text = xAs(7)
'            For C = 37 To 43
'                txt(C).Text = ""
'            Next C
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsFindIny As ADODB.Recordset
'                Set rsFindIny = objP.FetchListAllPelanggan(" Where c.KodeLg = '" & txt(1).Text & "' And c.NoRangka = '" & txt(19).Text & "'")
'                If rsFindIny.RecordCount > 0 Then
'                    Call cmdCancel_Click
'                    Call BindToControl_Master(rsFindIny, False)
'                    cmdEdit.Enabled = True
'                    cmdDelete.Enabled = True
'                    cmdOpen.Enabled = True
'                    cmdAdd.Visible = False
'                    cmdCancel.Enabled = False
'                End If
'            End If
'    End Select
'    On Error Resume Next
'    If xWAC = False Then
'        If Left(UCase(txt(21).Text), 4) = "ZOTH" Then
'            txt(23).Text = rsBy.Fields(47).value
'        End If
'    End If
'ErNul:
'End Sub
'
'Private Sub VSc_Change()
'    PicMove.Top = -(VSc.value)
'End Sub
'
'Private Sub VSc_Scroll()
'    PicMove.Top = -(VSc.value)
'End Sub
'
'Private Sub cmdSave_Click()
'    '-----------------------------------------------------------------------------------------------------------------
'    '   CEK DI ARMST KODE & NAMA
'    '-----------------------------------------------------------------------------------------------------------------
'    Dim rsCPL As ADODB.Recordset
'    Set rsCPL = objP.FetchCekPelangganACC(Trim(txt(1).Text))
'    If rsCPL.RecordCount > 0 Then
'        If Trim(rsCPL("NamaLgn")) = Trim(txt(2).Text) Then
'            Call isSavingData
'        Else
'            With msg
'                .PassingMsg = "Data Master Pelanggan Kode : " & txt(1).Text & "'" & Chr(13) & "Sudah Terdaftar Atas Nama : " & rsCPL("NamaLgn") & Chr(13) & "Silahkan ganti Kode Pelanggan!"
'                .passingTipe = 2
'                .Show vbModal
'            End With
'            Exit Sub
'        End If
'    Else
'        Call isSavingData
'    End If
'End Sub

Private Sub CmdAdd_Click()
    Call SetControlForm("add")
End Sub
'
'Private Sub cmdCreate_Click()
'    Dim rsRangka As ADODB.Recordset
'    '-------------------------------------------------------------
'    '   1. Cek NoRangka di CustDat
'    '-------------------------------------------------------------
'    Set rsRangka = objP.FetchExistRangka(txt(19).Text)
'    If rsRangka.RecordCount > 0 Then
'        With msg
'            .PassingMsg = "Nomor Rangka '" & txt(19).Text & "'" & Chr(13) & Chr(13) & "SUDAH TERDAFTAR !"
'            .passingTipe = 2
'            .Show vbModal
'        End With
'        Exit Sub
'    End If
'
'    Dim rsLgn As ADODB.Recordset
'    '-------------------------------------------------------------
'    '   Cek KodeLgn di ARMST
'    '-------------------------------------------------------------
'    Set rsLgn = objP.FetchPelangganACC(" Where KodeLgn = '" & txt(1).Text & "'")
'    If rsLgn.RecordCount > 0 Then
'    '-------------------------------------------------------------
'    '   2. Terdaftar di ARMST
'    '-------------------------------------------------------------
'        With msg
'            .PassingMsg = "'" & txt(1).Text & "' TERDAFTAR DI ACC!" & Chr(13) & "Nama : " & rsLgn.Fields(1).value & Chr(13) & "Alamat : " & rsLgn.Fields(2).value
'            .passingTipe = 2
'            .Show vbModal
'        End With
'        If Trim(rsLgn.Fields(1).value) = Trim(txt(2).Text) Then
'        '-------------------------------------------------------------
'        '   3. Ada | Jika nama sama, Cek di Custdat (Kode & Nama)
'        '-------------------------------------------------------------
'            Dim rsCRK As ADODB.Recordset
'            Set rsCRK = objP.FetchCekPelanggan(Trim(txt(1).Text), txt(2).Text)
'            If rsCRK.RecordCount > 0 Then
'                If Trim(rsCRK("NoRangka")) <> Trim(txt(19).Text) Then
'                '-------------------------------------------------------------
'                '  4. Ada | Jika No Rangka tidak sesuai, buat kode baru
'                '-------------------------------------------------------------
'                    With msg
'                        .PassingMsg = "Tambahkan No Rangka : " & Trim(txt(19).Text) & Chr(13) & " pada KodeLgn : " & Trim(txt(1).Text) & "?"
''''                        .PassingMsg = "No Rangka Tidak Sesuai : " & rsCRK("NoRangka") & Chr(13) & " Ganti Kode Pelanggan Baru"
'                        .passingTipe = 1
'                        .Show vbModal
'                        If .GetMsg = False Then
'                            Exit Sub
'                        Else
'                            Call isSavingData
'                        End If
'                    End With
'                Else
'                '-------------------------------------------------------------
'                '   4. Jika No Rangka sesuai, update data
'                '-------------------------------------------------------------
'                    With msg
'                        .PassingMsg = "'" & txt(1).Text & "' SUDAH TERDAFTAR !" & Chr(13) & Chr(13) & "YAKIN UPDATE DATA ???"
'                        .passingTipe = 1
'                        .Show vbModal
'                        If .GetMsg = False Then
'                            Exit Sub
'                        Else
'                            Call isSavingData
'                        End If
'                    End With
'                End If
'            Else
'            '-------------------------------------------------------------
'            '  Simpan Data
'            '-------------------------------------------------------------
'                Call isSavingData
'            End If
'        Else
'        '-------------------------------------------------------------
'        '   3. Gnti Kode Baru
'        '-------------------------------------------------------------
'            With msg
'                .PassingMsg = "'" & txt(1).Text & "' tidak sesuai ACC! (" & rsLgn("NamaLgn") & ")" & Chr(13) & " Ganti Kode Pelanggan Baru"
'                .passingTipe = 2
'                .Show vbModal
'            End With
'            txt(1).Text = ""
'            txt(1).SetFocus
'            Exit Sub
'        End If
'    Else
'    '-------------------------------------------------------------
'    '   2. Tidak terdaftar di ARMST
'    '-------------------------------------------------------------
'        Call isSavingData
'    End If
'End Sub
'
'Private Sub cmdEdit_Click()
'    If txt(19).Text = "" Then
'        With msg
'            .PassingMsg = "Nomor Rangka masih kosong, lakukan pembuatan (add) data pelanggan ?"
'            .passingTipe = 3
'            .Show vbModal
'        End With
'        Exit Sub
'    End If
'    Call SetControlForm("edit")
'End Sub
'
'Private Sub CmdDelete_Click()
'    With msg
'        .PassingMsg = "Hapus Master Pelanggan : " & txt(1).Text & ", No Rangka : " & txt(19).Text & " ?"
'        .passingTipe = 1
'        .Show vbModal
'        If .GetMsg = True Then
'            Call objP.Delete_ItemMasterPelanggan(txt(19).Text)
'            Call SetControlForm("load")
'        End If
'    End With
'End Sub
'
'Private Sub cmdOpen_Click()
'    '---------------------------------------------------------------------------------------------------------------------
'    '   02 SEPTEMBER 2015 |   REV. CEK ARMST & CUSTDAT (ANTISIPASI DOBEL KODELGN) |   ARIKDLG
'    '---------------------------------------------------------------------------------------------------------------------
'    '-----------------------------------------------------------------------------------------------------------------
'    '   CEK DI ARMST KODE & NAMA
'    '-----------------------------------------------------------------------------------------------------------------
'    Dim rsCPL As ADODB.Recordset
'    Set rsCPL = objP.FetchCekPelangganACC(Trim(txt(1).Text))
'    If rsCPL.RecordCount > 0 Then
'        If Trim(rsCPL("NamaLgn")) = Trim(txt(2).Text) Then
'            Set rsMove = objP.FetchListAllPelanggan(" Where KodeLg = '" & txt(1).Text & "' And NoRangka ='" & txt(19).Text & "'")
'            Unload Me
'        Else
'            With msg
'                .PassingMsg = "'" & txt(1).Text & "'" & Chr(13) & "Terdaftar Atas Nama : " & rsCPL("NamaLgn") & Chr(13) & "Silahkan ganti Kode Pelanggan!"
'                .passingTipe = 2
'                .Show vbModal
'            End With
'            txt(1).Text = ""
'            txt(1).SetFocus
'            Exit Sub
'        End If
'    End If
'End Sub

Private Sub cmdCancel_Click()
    Call SetControlForm("load")
End Sub

Private Sub cmdExit_Click()
   
    Unload Me
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Txt(Index).BackColor = fColor Then
        Txt(Index).SelStart = 0
        Txt(Index).SelLength = Len(Txt(Index).text)
        Clipboard.SetText Txt(Index).text
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Txt(Index).BackColor = fColor Then
        Txt(Index).SelStart = 0
        Txt(Index).SelLength = Len(Txt(Index).text)
        Clipboard.SetText Txt(Index).text
    End If
End Sub

'Private Sub Zom_Click()
'    onFind = True
'    onF12 = True
'    ft.ShowForm = "PLGN"
'    ft.Show vbModal
'    txt(1).SetFocus
'    onF12 = False
'End Sub

Private Function IsValid() As Boolean
    IsValid = True
    If Txt(15) = "" Then Txt(14) = ""
    If Txt(18) = "" Then Txt(17) = ""
    If Txt(22) = "" Then Txt(21) = ""
    If Txt(24) = "" Then Txt(23) = ""
        
    If Txt(1).text = "" Or Txt(2).text = "" Or txtNpwp.text = "" Or Txt(4).text = "" Then
        IsValid = False
        nValid = "[Nama/Alamat]"
        Exit Function
    End If
    If Txt(7).text = "" And Txt(8).text = "" Then
        IsValid = False
        nValid = "[Telp Pelanggan]"
        Exit Function
    End If
    If Txt(11).text = "" Or Txt(13).text = "" _
    Or Txt(21).text = "" Or Txt(24).text = "" Or Txt(26).text = "" Or Txt(27).text = "" _
    Or Txt(21).text = "" Then
        IsValid = False
        nValid = "[Data Kendaraan]"
        Exit Function
    End If
    If Txt(37).text = "" Or Txt(38).text = "" Or Txt(39).text = "" Then
        IsValid = False
        nValid = "[Data Surat]"
        Exit Function
    End If
    If Txt(23).text <> "" Then
        If Txt(31).text = "" Or Txt(32).text = "" Or Txt(33).text = "" Or Txt(36).text = "" _
        Or DtPolish0.text = "__/__/____" Or DtPolish1.text = "__/__/____" Then
            IsValid = False
            nValid = "[Data Polis]"
            Exit Function
        End If
        If DtPolish0.value >= DtPolish1.value Then
            IsValid = False
            nValid = "[Tanggal Polis]"
            Exit Function
        End If
    End If
    If Txt(42).text = "" And Txt(43).text = "" Then
        IsValid = False
        nValid = "[Telp Surat]"
        Exit Function
    End If
    If Len(Txt(19).text) < 9 Then
        IsValid = False
        nValid = "[No Rangka Tidak Valid]"
        Exit Function
    End If
    Dim rsWx As ADODB.Recordset
    Set rsWx = objWl.FetchItemWilayah(Left(Txt(1).text, 3))
    If rsWx.RecordCount = 0 Then
        IsValid = False
        nValid = "[Kode Wilayah pada KODE LANGGAN, tidak terdaftar]"
        Exit Function
    End If
End Function

Private Function LastID(tipe As String) As String
    Dim rsI As New ADODB.Recordset
    LastID = ""
    Set rsI = objSID.FetchLastID(tipe)
    If rsI.RecordCount > 0 Then
        rsI.MoveFirst
        LastID = rsI("Kode") & Format(rsI("LastUrut") + 1, "0000000")
    End If
End Function

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Dim rsCari As ADODB.Recordset
    
    If Index = 2 Then
 
    ElseIf Index = 20 Then
        If NulSID(Txt(Index)) <> "" Then
            Dim rsAgama As ADODB.Recordset
            Set rsAgama = objO.FetchData("SELECT * FROM [SINGLEIDORI]..agama WITH (NOLOCK) WHERE agama = '" & NulSID(Txt(Index)) & "'")
            If rsAgama.EOF Then
                Txt(Index) = ""
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDAGAMA"
                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 7 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..PEKERJAAN WITH (NOLOCK) WHERE pekerjaan = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                Txt(Index) = ""
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDPEKERJAAN"
                frmCari.Show (vbModal)
            End If
            If NulSID(UCase(Txt(Index))) <> "LAINNYA" Then
                Txt(23).BackColor = &HC0E0FF
                Txt(23) = ""
            End If
        End If
    ElseIf Index = 21 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..HOBI WITH (NOLOCK) WHERE HOBI = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                Txt(Index) = ""
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDHOBI"
                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 22 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..OLAHRAGA WITH (NOLOCK) WHERE OLAHRAGA = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                Txt(Index) = ""
                gMainKeyCari = Txt(Index)
                gSearchOption = "SIDOLAHRAGA"
                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 63 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = Txt(Index)
                Txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
                
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 64 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(Txt(63)) & "' and KOTA = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = Txt(Index)
                Txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 65 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(Txt(63)) & "' and KOTA = '" & NulSID(Txt(64)) & "' and KECAMATAN = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = Txt(Index)
                Txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 66 Then
        If NulSID(Txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(Txt(63)) & "' and KOTA = '" & NulSID(Txt(64)) & "' and KECAMATAN = '" & NulSID(Txt(65)) & "' and KELURAHAN = '" & NulSID(Txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = Txt(Index)
                Txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 9 Or Index = 10 Then
        Txt(Index) = Format(NulSID(Txt(Index)), "0##")
    End If
End Sub
