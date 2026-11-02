VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOSupportTL 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12750
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   12750
   Begin VB.Frame frameKirimTL 
      Height          =   5535
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CheckBox ckProblemSolve 
         Alignment       =   1  'Right Justify
         Caption         =   "Problem Solve"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   45
         Top             =   4320
         Width           =   1695
      End
      Begin VB.ComboBox cbRTJ 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmHOSupportTL.frx":0000
         Left            =   7440
         List            =   "frmHOSupportTL.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3240
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.ComboBox cbKirimpbt 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmHOSupportTL.frx":0019
         Left            =   7440
         List            =   "frmHOSupportTL.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtRootCause 
         Height          =   1095
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   3585
         Width           =   7215
      End
      Begin VB.TextBox txtPerbaikan 
         Height          =   1095
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   2040
         Width           =   7215
      End
      Begin VB.TextBox txtProblem 
         Height          =   1095
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   7215
      End
      Begin ucXPButton.XPButton cmdBatalKirim 
         Height          =   495
         Left            =   10200
         TabIndex        =   22
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         MICON           =   "frmHOSupportTL.frx":0032
         PICN            =   "frmHOSupportTL.frx":004E
         PICH            =   "frmHOSupportTL.frx":6647
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKirim 
         Height          =   495
         Left            =   9000
         TabIndex        =   23
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         MICON           =   "frmHOSupportTL.frx":CEF6
         PICN            =   "frmHOSupportTL.frx":CF12
         PICH            =   "frmHOSupportTL.frx":13715
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc adcProblem 
         Height          =   330
         Left            =   120
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbKategoriProblem 
         Height          =   375
         Left            =   7440
         TabIndex        =   33
         Top             =   480
         Width           =   3915
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   6906
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Kategori Problem -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbKategoriKomponen 
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         Top             =   1440
         Width           =   3915
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   6906
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Kategori Komponen -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc adcKomponen 
         Height          =   330
         Left            =   1440
         Top             =   4800
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin TDBDate6Ctl.TDBDate tglKirim 
         Height          =   375
         Left            =   10080
         TabIndex        =   51
         Top             =   3600
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
         Calendar        =   "frmHOSupportTL.frx":1A2BB
         Caption         =   "frmHOSupportTL.frx":1A3E7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOSupportTL.frx":1A453
         Keys            =   "frmHOSupportTL.frx":1A471
         Spin            =   "frmHOSupportTL.frx":1A4CF
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
         Text            =   "06/03/2020"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   43896
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tglRespon 
         Height          =   375
         Left            =   10080
         TabIndex        =   52
         Top             =   3960
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
         Calendar        =   "frmHOSupportTL.frx":1A4F7
         Caption         =   "frmHOSupportTL.frx":1A623
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOSupportTL.frx":1A68F
         Keys            =   "frmHOSupportTL.frx":1A6AD
         Spin            =   "frmHOSupportTL.frx":1A70B
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
         Text            =   "06/03/2020"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   43896
         CenturyMode     =   0
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Tanggal Kirim Ke NTL : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   54
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Tanggal Respon NTL : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   53
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "RTJ Atau TIDAK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7560
         TabIndex        =   44
         Top             =   2925
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   11160
         TabIndex        =   43
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   11160
         TabIndex        =   40
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "KIRIM PBT KE NTL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7560
         TabIndex        =   39
         Top             =   2085
         Width           =   3375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   7440
         Shape           =   4  'Rounded Rectangle
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   11160
         TabIndex        =   38
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "KATEGORI KOMPONEN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7560
         TabIndex        =   37
         Top             =   1125
         Width           =   3375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   7440
         Shape           =   4  'Rounded Rectangle
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   11160
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "KATEGORI PROBLEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7560
         TabIndex        =   34
         Top             =   165
         Width           =   3375
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   7440
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7080
         TabIndex        =   31
         Top             =   3225
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "PENYEBAB MASALAH ( ROOT CAUSE )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   3280
         Width           =   6615
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3225
         Width           =   7215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7080
         TabIndex        =   28
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "PERBAIKAN YANG DILAKUKAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   1725
         Width           =   6615
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   7215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7080
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "PROBLEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   160
         Width           =   6615
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   7215
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   7440
         Shape           =   4  'Rounded Rectangle
         Top             =   2880
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.Frame frameTidakKirim 
      Height          =   5415
      Left            =   12240
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox txtKeteranganTidakKirim 
         Height          =   2655
         Left            =   2160
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   7695
      End
      Begin ucXPButton.XPButton cmdBatalTidakKirim 
         Height          =   495
         Left            =   8520
         TabIndex        =   18
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         MICON           =   "frmHOSupportTL.frx":1A733
         PICN            =   "frmHOSupportTL.frx":1A74F
         PICH            =   "frmHOSupportTL.frx":20D48
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpanTidakKirim 
         Height          =   495
         Left            =   7320
         TabIndex        =   19
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         MICON           =   "frmHOSupportTL.frx":275F7
         PICN            =   "frmHOSupportTL.frx":27613
         PICH            =   "frmHOSupportTL.frx":2DE16
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inputkan hasil konfirmasi Teknisi / FO terkai Lead Time pekerjaan yang lama."
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
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   7665
      End
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   6360
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
      EnableSearchControl=   0   'False
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   5055
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbFO 
         Height          =   375
         Left            =   1680
         TabIndex        =   46
         Top             =   960
         Visible         =   0   'False
         Width           =   3255
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   5741
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   120
         Top             =   1080
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Batal"
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
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1370
         Width           =   810
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1370
         Width           =   930
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1370
         Width           =   810
      End
      Begin TDBDate6Ctl.TDBDate Tgl1 
         Height          =   330
         Left            =   1680
         TabIndex        =   7
         Top             =   570
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmHOSupportTL.frx":349BC
         Caption         =   "frmHOSupportTL.frx":34AE8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOSupportTL.frx":34B54
         Keys            =   "frmHOSupportTL.frx":34B72
         Spin            =   "frmHOSupportTL.frx":34BD0
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
      Begin TDBDate6Ctl.TDBDate Tgl2 
         Height          =   330
         Left            =   3480
         TabIndex        =   8
         Top             =   570
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmHOSupportTL.frx":34BF8
         Caption         =   "frmHOSupportTL.frx":34D24
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOSupportTL.frx":34D90
         Keys            =   "frmHOSupportTL.frx":34DAE
         Spin            =   "frmHOSupportTL.frx":34E0C
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ".:: LAPORAN SUPPORT ATL ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Support ATL"
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
         TabIndex        =   10
         Top             =   645
         Width           =   1500
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
         Left            =   3240
         TabIndex        =   9
         Top             =   645
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerButton btnKembali 
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BackColor       =   255
      BorderStyle     =   1
      Caption         =   "KEMBALI KE MENU"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   11760
      Top             =   3600
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton2 
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BackColor       =   33023
      HighlightColor  =   16777215
      BorderStyle     =   1
      Caption         =   "LAPORAN"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   0
   End
   Begin AIFCmp1.asxPowerButton cmdProses 
      Height          =   495
      Left            =   9720
      TabIndex        =   47
      Top             =   4125
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BackColor       =   49152
      BorderStyle     =   1
      Caption         =   "PROSES"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   0
   End
   Begin TrueOleDBGrid70.TDBGrid listSelesaiCuci 
      Height          =   3015
      Left            =   240
      TabIndex        =   48
      Top             =   960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "cbStatus"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10).DropDown=   "cbStatus"
      Columns(10).DropDown.vt=   8
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11).DropDown=   "TDBDropDown1"
      Columns(11).DropDown.vt=   8
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(10).DropDownList=1"
      Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   0
      DeadAreaBackColor=   12577002
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=28,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
      _StyleDefs(84)  =   "Named:id=33:Normal"
      _StyleDefs(85)  =   ":id=33,.parent=0"
      _StyleDefs(86)  =   "Named:id=34:Heading"
      _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=34,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=35:Footing"
      _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=36:Selected"
      _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=37:Caption"
      _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(95)  =   "Named:id=38:HighlightRow"
      _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=39:EvenRow"
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
   End
   Begin AIFCmp1.asxPowerButton cmdRefresh 
      Height          =   495
      Left            =   7800
      TabIndex        =   55
      Top             =   4125
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BackColor       =   16711680
      BorderStyle     =   1
      Caption         =   "REFRESH"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "BELUM INPUT HASIL = "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8280
      TabIndex        =   50
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblJumlahHasil 
      BackColor       =   &H00FFFFFF&
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   49
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   11640
      X2              =   11040
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   11640
      X2              =   10800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   11640
      X2              =   10680
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   11640
      X2              =   120
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   11640
      X2              =   120
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar SUPPORT ATL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmHOSupportTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsSelesaiCuci As ADODB.Recordset
Private rsTempSelesaiCuci As New ADODB.Recordset

Private rsAntriCuci As ADODB.Recordset
Private rsTempAntriCuci As New ADODB.Recordset

Private strSql As String
Private isPengering As Boolean

Dim i As Double 'Integer
Private Pilih As Double 'Integer
Private idSimpan As Double 'Integer
Private pesan As String

Private Sub FetchCombo()
    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = "SELECT DISTINCT a.[GROUP] kode, b.Nama FROM MEKANIK a WITH (NOLOCK) LEFT JOIN ManPower B WITH (NOLOCK) ON a.[GROUP] = b.Kode"
    cbFO.DataSourceList = Adodc1
    cbFO.DataFieldList = "kode"
    cbFO.Columns(0).Width = 1200
    cbFO.Columns(1).Width = 4000
    cbFO.Columns(0).Caption = "Kode Dealer"
    cbFO.Columns(1).Caption = "Nama Petugas"
End Sub

Private Function createAntriCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTempAntriCuci.State Then rsTempAntriCuci.Close
3         Set rsTempAntriCuci = Nothing
4         With rsTempAntriCuci
5             .Fields.Append "No", adDouble 'adInteger
6             .Fields.Append "Id", adDouble 'adInteger
7             .Fields.Append "NoPolisi", adVarChar, 25
8             .Fields.Append "SA", adVarChar, 5
9             .Fields.Append "NoWo", adVarChar, 15
10            .Fields.Append "JenisKendaraan", adVarChar, 50
11            .Fields.Append "TahunKendaraan", adVarChar, 25
12            .Fields.Append "Perbaikan", adVarChar, 200
13            .Fields.Append "Teknisi", adVarChar, 25
14            .Fields.Append "Tanggal", adVarChar, 25
15            .Fields.Append "Ditanggapi", adBoolean
16            .Open
17        End With
18        If rsTempAntriCuci.State Then createAntriCuci = True
19        Exit Function
          
vb_error:
20        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createSelesaiCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTempSelesaiCuci.State Then rsTempSelesaiCuci.Close
3         Set rsTempSelesaiCuci = Nothing
4         With rsTempSelesaiCuci
5             .Fields.Append "No", adDouble 'adInteger
6             .Fields.Append "Id", adDouble 'adInteger
7             .Fields.Append "NoPolisi", adVarChar, 25
8             .Fields.Append "NoWo", adVarChar, 15
9             .Fields.Append "JenisKendaraan", adVarChar, 50
10            .Fields.Append "Perbaikan", adVarChar, 200
11            .Fields.Append "Teknisi", adVarChar, 25
12            .Fields.Append "Tanggal", adVarChar, 25
13            .Fields.Append "Input", adBoolean
14            If Log.rsA = 1 Then
15                .Fields.Append "KodeModel", adVarChar, 50
16                .Fields.Append "RSA", adVarChar, 10
17                .Fields.Append "FO", adVarChar, 10
18            Else
19                .Fields.Append "SA", adVarChar, 5
20                .Fields.Append "TahunKendaraan", adVarChar, 25
21                .Fields.Append "Keterangan", adVarChar, 200
22                .Fields.Append "Status", adVarChar, 100
23            End If
24            .Open
25        End With
26        If rsTempSelesaiCuci.State Then createSelesaiCuci = True
27        Exit Function
          
vb_error:
28        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function


Private Sub refreshGridAntriCuci()
1         If Log.rsA = 1 Then
2             sSQL = "select n.id,j.no_wo,j.no_polisi,j.fmdl,j.fs_th_mobil,j.rep1,j.rep2,mech,front,tglentry "
3             sSQL = sSQL & " from needsupporttl n WITH(NOLOCK) "
4             sSQL = sSQL & " INNER JOIN jobdta J WITH(NOLOCK) ON j.no_wo = n.no_wo "
5             sSQL = sSQL & " WHERE n.kodedealer = '" & Right(gDb, 2) & "' and n.tglFu = '01/01/1900' order by tglentry asc"
6         Else
7             sSQL = "select n.id,j.no_wo,j.no_polisi,j.fmdl,j.fs_th_mobil,j.rep1,j.rep2,mech,front,tglentry from needsupporttl n INNER JOIN jobdta J WITH(NOLOCK) ON j.no_wo = n.no_wo WHERE n.tglFu = '01/01/1900' order by tglentry asc"
8         End If
          
          
9         Set rsAntriCuci = objO.FetchData(sSQL)

          Dim sisaWaktu As String
          Dim tipe As String
          Dim waktuSekarang As String
10        If createAntriCuci Then
11            If rsAntriCuci.RecordCount > 0 Then
12                rsAntriCuci.MoveFirst
13                i = 1
14                Do While Not rsAntriCuci.EOF
15                    With rsTempAntriCuci
16                        .AddNew
17                        !No = i '
18                        !id = CDbl(rsAntriCuci("id"))
19                        !Ditanggapi = False
20                        !NoPolisi = Trim(rsAntriCuci("no_polisi"))
21                        !sa = Trim(rsAntriCuci("FRONT"))
22                        !noWO = Trim(rsAntriCuci("no_wo"))
23                        !JenisKendaraan = Trim(rsAntriCuci("fmdl"))
24                        !TahunKendaraan = Trim(rsAntriCuci("fs_th_mobil"))
25                        !Perbaikan = Trim(rsAntriCuci("rep1")) & " ; " & Trim(rsAntriCuci("rep2"))
26                        !Teknisi = Trim(rsAntriCuci("MECH"))
27                        !tanggal = Format(Trim(rsAntriCuci("tglentry")), "DD/MM/YYYY HH:MM:SS")
28                        .Update
29                        .MoveNext
30                    End With
31                    rsAntriCuci.MoveNext
32                    i = i + 1
33                Loop
34            End If
35        End If
36        listAntriCuci.DataSource = rsTempAntriCuci
37        Call headColumnAntriCuci
38        lblJumlahAntri.Caption = rsTempAntriCuci.RecordCount
          'lblHariAC.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
End Sub

Private Sub headColumnAntriCuci()
    With listAntriCuci
        .Columns(0).Caption = "No"
        .Columns(0).DataField = "No"
        .Columns(0).Width = 500
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Caption = "Ditanggapi"
        .Columns(1).DataField = "Ditanggapi"
        .Columns(1).Width = 1000
        .Columns(1).Alignment = dbgCenter
        .Splits(0).Columns(1).Locked = False
        
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Splits(0).Columns(2).Locked = False
        
        .Columns(3).Caption = "No Polisi"
        .Columns(3).DataField = "NoPolisi"
        .Columns(3).Width = 1000
        .Splits(0).Columns(3).Locked = False
                
        .Columns(4).Caption = "Jenis Kendaraan"
        .Columns(4).DataField = "JenisKendaraan"
        .Columns(4).Width = 2000
        .Splits(0).Columns(4).Locked = False
        
        .Columns(5).Caption = "Tahun Kendaraan"
        .Columns(5).DataField = "TahunKendaraan"
        .Columns(5).Width = 1500
        .Splits(0).Columns(5).Locked = False
        
        .Columns(6).Caption = "Perbaikan"
        .Columns(6).DataField = "Perbaikan"
        .Columns(6).Width = 3000
        .Splits(0).Columns(6).Locked = False
        
        .Columns(7).Caption = "Teknisi"
        .Columns(7).DataField = "Teknisi"
        .Columns(7).Width = 700
        .Splits(0).Columns(7).Locked = False
        
        .Columns(8).Caption = "SA"
        .Columns(8).DataField = "SA"
        .Columns(8).Width = 700
        .Splits(0).Columns(8).Locked = False
        
        .Columns(9).Caption = "Tgl Request TL"
        .Columns(9).DataField = "Tanggal"
        .Columns(9).Width = 1200
        .Splits(0).Columns(9).Locked = False
    End With
End Sub

Private Sub refreshGridSelesaiCuci()
1         If Log.rsA = 1 Then
2             strSql = "select a.id,a.nopol as no_polisi,a.no_wo,a.tglInfo,b.FS_NM_DEALER,a.nopol,a.model,a.sumberTL,a.DetailProblem,"
3             strSql = strSql & " C.rep1 , C.REP2, a.tech, a.FO,a.tglentry"
4             strSql = strSql & " from NeedSupportTL a with (nolock)"
5             strSql = strSql & " left join T_DEALER b with (nolock) on 'bengkel'+a.KodeDealer = b.DBASECBG and b.FS_LOG_DEL = ''"
6             strSql = strSql & " left join JOBDTA c with (nolock) on c.FS_KD_DEALER = b.FS_KD_DEALER and c.NO_WO = a.no_wo"
              ''strSql = strSql & " where solve <> 'YA' and (kirimATL = 'YA' or DATEDIFF(day,tglentry,getdate())>=1) "
7             strSql = strSql & " where solve <> 'YA' and (kirimATL = 'YA') "
8             strSql = strSql & " order by a.tglentry asc"
9         Else
10            strSql = "select n.id,j.no_wo,j.no_polisi,j.fmdl,j.fs_th_mobil,j.rep1,j.rep2,mech,front,tglentry from needsupporttl n INNER JOIN jobdta J WITH(NOLOCK) ON j.no_wo = n.no_wo WHERE n.tglFu <> '01/01/1900' and tglInfo = '01/01/1900' order by tglentry asc"
11        End If
12        Set rsSelesaiCuci = objO.FetchDataHO(strSql)
13        Pilih = 0
14        If createSelesaiCuci Then
15            If rsSelesaiCuci.RecordCount > 0 Then
16                i = 1
17                rsSelesaiCuci.MoveFirst
18                Do While Not rsSelesaiCuci.EOF
19                    With rsTempSelesaiCuci
20                        .AddNew
21                        !No = i '
22                        !id = CDbl(rsSelesaiCuci("id"))
23                        !Input = False
24                        !NoPolisi = Trim(rsSelesaiCuci("no_polisi"))
25                        !noWO = Trim(rsSelesaiCuci("no_wo"))
26                        !JenisKendaraan = Trim(rsSelesaiCuci("model"))
27                        !Teknisi = Trim(rsSelesaiCuci("tech"))
28                        !tanggal = Format(Trim(rsSelesaiCuci("tglentry")), "DD/MM/YYYY HH:MM:SS")
29                        If Log.rsA = 1 Then
                             ' !KodeModel = Trim(rsSelesaiCuci("MDL"))
30                            !rsA = Trim(rsSelesaiCuci("sumberTL"))
31                            !FO = Trim(rsSelesaiCuci("FO"))
32                            If NulSID(rsSelesaiCuci("sumberTL")) = "PBT" Then
33                                !Perbaikan = Trim(rsSelesaiCuci("DetailProblem"))
34                            Else
35                                !Perbaikan = Trim(rsSelesaiCuci("rep1")) & " ; " & Trim(rsSelesaiCuci("rep2"))
36                            End If
37                        Else
38                            !Perbaikan = Trim(rsSelesaiCuci("rep1")) & " ; " & Trim(rsSelesaiCuci("rep2"))
39                            !TahunKendaraan = Trim(rsSelesaiCuci("fs_th_mobil"))
40                            !sa = Trim(rsSelesaiCuci("FRONT"))
41                            !keterangan = ""
42                            !Status = ""
43                        End If
44                        .Update
45                        .MoveNext
46                    End With
47                    i = i + 1
48                    rsSelesaiCuci.MoveNext
49                Loop
50            End If
51        End If
          
52        listSelesaiCuci.DataSource = rsTempSelesaiCuci
53        Call headColumns
54        lblJumlahHasil.Caption = rsTempSelesaiCuci.RecordCount
End Sub

Private Sub headColumns()
    With listSelesaiCuci
        If Log.rsA = 1 Then
            .Columns(0).Caption = "No"
            .Columns(0).DataField = "No"
            .Columns(0).Width = 500
            .Splits(0).Columns(0).Locked = True
            
            .Columns(1).Caption = "Input"
            .Columns(1).DataField = "input"
            .Columns(1).Width = 1000
            .Columns(1).Alignment = dbgCenter
            .Splits(0).Columns(1).Locked = False
            
            .Columns(2).Caption = "Tgl Request TL"
            .Columns(2).DataField = "Tanggal"
            .Columns(2).Width = 1200
            .Splits(0).Columns(2).Locked = True
            
            .Columns(3).Caption = "No Polisi"
            .Columns(3).DataField = "NoPolisi"
            .Columns(3).Width = 1000
            .Splits(0).Columns(3).Locked = True
            
            .Columns(4).Caption = "No Wo"
            .Columns(4).DataField = "NoWo"
            .Columns(4).Width = 1000
            .Splits(0).Columns(4).Locked = True
            
            
            .Columns(5).Caption = "Kode Model"
            .Columns(5).DataField = "KodeModel"
            .Columns(5).Width = 1500
            .Splits(0).Columns(5).Locked = True
                    
            .Columns(6).Caption = "Jenis Kendaraan"
            .Columns(6).DataField = "JenisKendaraan"
            .Columns(6).Width = 2000
            .Splits(0).Columns(6).Locked = True
            
            .Columns(7).Caption = "RSA"
            .Columns(7).DataField = "RSA"
            .Columns(7).Width = 700
            .Splits(0).Columns(7).Locked = True
            
            .Columns(8).Caption = "Perbaikan"
            .Columns(8).DataField = "Perbaikan"
            .Columns(8).Width = 3000
            .Splits(0).Columns(8).Locked = True
            
            .Columns(9).Caption = "Teknisi"
            .Columns(9).DataField = "Teknisi"
            .Columns(9).Width = 700
            .Splits(0).Columns(9).Locked = True
            
            
            .Columns(10).Caption = "F O"
            .Columns(10).DataField = "FO"
            .Columns(10).Width = 700
            .Splits(0).Columns(10).Locked = False
            
            .Columns(11).Caption = "Keterangan"
            .Columns(11).DataField = "keterangan"
            .Columns(11).Width = 1200
            .Columns(11).Visible = False
            .Splits(0).Columns(11).Locked = False
        Else
            .Columns(0).Caption = "No"
            .Columns(0).DataField = "No"
            .Columns(0).Width = 500
            .Splits(0).Columns(0).Locked = True
            
            .Columns(1).Caption = "Input"
            .Columns(1).DataField = "input"
            .Columns(1).Width = 1000
            .Columns(1).Alignment = dbgCenter
            .Splits(0).Columns(1).Locked = False
            
            .Columns(2).Caption = "No Wo"
            .Columns(2).DataField = "NoWo"
            .Columns(2).Width = 1000
            .Splits(0).Columns(2).Locked = True
            
            .Columns(3).Caption = "No Polisi"
            .Columns(3).DataField = "NoPolisi"
            .Columns(3).Width = 1000
            .Splits(0).Columns(3).Locked = True
                    
            .Columns(4).Caption = "Jenis Kendaraan"
            .Columns(4).DataField = "JenisKendaraan"
            .Columns(4).Width = 2000
            .Splits(0).Columns(4).Locked = True
            
            .Columns(5).Caption = "Tahun Kendaraan"
            .Columns(5).DataField = "TahunKendaraan"
            .Columns(5).Width = 1500
            .Splits(0).Columns(5).Locked = True
            
            .Columns(6).Caption = "Perbaikan"
            .Columns(6).DataField = "Perbaikan"
            .Columns(6).Width = 3000
            .Splits(0).Columns(6).Locked = True
            
            .Columns(7).Caption = "Teknisi"
            .Columns(7).DataField = "Teknisi"
            .Columns(7).Width = 700
            .Splits(0).Columns(7).Locked = True
            
            .Columns(8).Caption = "SA"
            .Columns(8).DataField = "SA"
            .Columns(8).Width = 700
            .Splits(0).Columns(8).Locked = True
            
            .Columns(9).Caption = "Tgl Request TL"
            .Columns(9).DataField = "Tanggal"
            .Columns(9).Width = 1200
            .Splits(0).Columns(9).Locked = True
            
            .Columns(10).Caption = "Status"
            .Columns(10).DataField = "status"
            .Columns(10).Width = 1200
            .Splits(0).Columns(10).Locked = False
            
            .Columns(11).Caption = "Keterangan"
            .Columns(11).DataField = "keterangan"
            .Columns(11).Width = 1200
            .Splits(0).Columns(11).Locked = False
        End If
    End With
End Sub

 

Private Sub asxPowerButton2_Click()
    Frame1.Visible = True
    asxPowerButton2.Enabled = False
    Tgl1 = isnow
    Tgl2 = isnow
End Sub

Private Sub btnKembali_Click()
    Unload Me
End Sub

Private Sub btnSimpanAntri_Click()
    If Not rsTempAntriCuci.EOF Then
        If MsgBox("Apakah yakin menanggapi Antrian TL ?", vbYesNo) = vbYes Then
            
            Do While Not rsTempAntriCuci.EOF
                If rsTempAntriCuci("Ditanggapi") Then
                    If Log.rsA = 1 Then
                        Call objO.FetchData("UPDATE NEEDSUPPORTTL SET tglfu = getdate(), userfu = '" & gSA & "' WHERE id = '" & CDbl(rsTempAntriCuci("id")) & "'")
                    Else
                        Call objO.FetchData("UPDATE NEEDSUPPORTTL SET tglfu = getdate(), userfu = '" & gSA & "' WHERE id = '" & CDbl(rsTempAntriCuci("id")) & "'")
                    End If
                    
                End If
                rsTempAntriCuci.MoveNext
            Loop
        End If
        Call refreshGridAntriCuci
    Else
        MsgBox "Tidak ada data untuk disimpan", vbInformation
    End If
End Sub

Private Sub cbKirimpbt_Change()
    If NulSID(cbKirimpbt.Text) = "YA" Then
        ckProblemSolve.value = vbUnchecked
        ckProblemSolve.Enabled = False
    Else
        ckProblemSolve.Enabled = True
    End If
End Sub

Private Sub cbKirimpbt_Click()
    If NulSID(cbKirimpbt.Text) = "YA" Then
        ckProblemSolve.value = vbUnchecked
        ckProblemSolve.Enabled = False
    Else
        ckProblemSolve.Enabled = True
    End If
End Sub

Private Sub cmdBatalKirim_Click()
    frameKirimTL.Visible = False
End Sub

Private Sub cmdBatalTidakKirim_Click()
    frameTidakKirim.Visible = False
End Sub

Private Sub cmdKirim_Click()
          Dim boleh   As Boolean
          Dim begin   As Boolean
          Dim pesan   As String
          
1        On Error GoTo cmdKirim_Click_Error

2         boleh = True
3         begin = False
4         If NulSID(txtPerbaikan.Text) = "" Then
5             pesan = pesan & "Perbaikan yang dilakukan harus diisi " & vbCrLf
6             boleh = False
7         End If
8         If NulSID(txtProblem.Text) = "" Then
9             pesan = pesan & "Detail problem harus diisi " & vbCrLf
10            boleh = False
11        End If
12        If NulSID(txtRootCause.Text) = "" Then
13            pesan = pesan & "Root Cause harus diisi " & vbCrLf
              
14            boleh = False
15        End If
          
16        If NulSID(cbKategoriProblem.value) = "" Then
17            pesan = pesan & "Kategori Problem harus dipilih. " & vbCrLf
18            boleh = False
19        End If
          
20        If NulSID(cbKategoriKomponen.value) = "" Then
21            pesan = pesan & "Kategori Komponen harus dipilih. " & vbCrLf
22            boleh = False
23        End If
          
24        If NulSID(cbKirimpbt.Text) = "" Then
25            pesan = pesan & "Kirim PBT harus dipilih. " & vbCrLf
26            boleh = False
27        End If
      '
      '    If NulSID(cbRTJ.Text) = "" Then
      '        pesan = pesan & "Pilihan RTJ harus dipilih. " & vbCrLf
      '        boleh = False
      '    End If
28        If boleh Then
29            If ckProblemSolve.value = vbChecked Then
30                sSQL = "UPDATE NEEDSUPPORTTL "
31                sSQL = sSQL & " SET SOLVE = 'YA',kirimNTL = 'TIDAK',solveATL = 'YA',tglInfoATL = getdate(),userInfoATL = '" & gSA & "', "
32                sSQL = sSQL & " idProblemATL = '" & cbKategoriProblem.value & "',idKomponenATL = '" & cbKategoriKomponen.value & "',"
33                sSQL = sSQL & " DetailProblemATL = '" & NulSID(txtProblem.Text) & "',PerbaikanATL = '" & NulSID(txtPerbaikan.Text) & "',RootCauseATL = '" & NulSID(txtRootCause.Text) & "',tglKirimNTL='" & Format(tglKirim.value, "yyyy-mm-dd") & "',tglResponNTL='" & Format(tglRespon.value, "yyyy-mm-dd") & "',tglSolved = getdate()"
34                sSQL = sSQL & " WHERE ID = " & idSimpan
35            Else
36                sSQL = "UPDATE NEEDSUPPORTTL "
37                sSQL = sSQL & " SET solveATL = 'TIDAK',kirimNTL = 'YA',tglInfoATL = getdate(),userInfoATL = '" & gSA & "', "
38                sSQL = sSQL & " idProblemATL = '" & cbKategoriProblem.value & "',idKomponenATL = '" & cbKategoriKomponen.value & "',"
39                sSQL = sSQL & " DetailProblemATL = '" & NulSID(txtProblem.Text) & "',PerbaikanATL = '" & NulSID(txtPerbaikan.Text) & "',RootCauseATL = '" & NulSID(txtRootCause.Text) & "',tglKirimNTL='" & Format(tglKirim.value, "yyyy-mm-dd") & "',tglResponNTL='" & Format(tglRespon.value, "yyyy-mm-dd") & "'"
40                sSQL = sSQL & " WHERE ID = " & idSimpan
41            End If
42            Call objO.InsertDataHO(sSQL)
43            Call refreshGridSelesaiCuci
44            MsgBox "Data berhasil disimpan.", vbInformation
45            frameKirimTL.Visible = False
46        Else
47            MsgBox "Silahkan lengkapi data berikut :" & vbCrLf & pesan, vbInformation
48        End If
49       On Error GoTo 0
50       Exit Sub

cmdKirim_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
51        varErlError = Erl
52        varErrNumber = err.Number
53        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdKirim_Click of Form frmHOSupportTL")
          
End Sub

Private Sub cmdProses_Click()
1         If Pilih = 0 Then
2             MsgBox "Silahkan mencentang salah satu data", vbCritical
3             Exit Sub
4         ElseIf Pilih > 1 Then
5             MsgBox "Proses simpan Evaluasi Support ATL hanya diperbolehkan 1 data", vbCritical
6             Exit Sub
7         Else
8             If Not rsTempSelesaiCuci.EOF Then
9                 rsTempSelesaiCuci.MoveFirst
                  Dim rsNeedSupport     As ADODB.Recordset
                  
10                Do While Not rsTempSelesaiCuci.EOF
11                    pesan = ""
12                    idSimpan = 0
13                    If rsTempSelesaiCuci("Input") Then
14                        idSimpan = CDbl(rsTempSelesaiCuci("id"))
15                        If Log.rsA = 1 Then
16                            Set rsNeedSupport = objO.FetchDataHO("SELECT SOLVE,tglinfo,idProblem,idKomponen,rtj,DetailProblem,Perbaikan,RootCause,tglKirimNTL,tglResponNTL FROM NEEDSUPPORTTL WITH (NOLOCK) WHERE ID = " & idSimpan)
17                            txtPerbaikan.Text = ""
18                            txtProblem.Text = ""
19                            txtRootCause.Text = ""
20                            ckProblemSolve.value = vbUnchecked
21                            tglKirim.value = "01/01/1900"
22                            tglRespon.value = "01/01/1900"
23                            Call refreshCombo
24                            If Not rsNeedSupport.EOF Then
25                                txtPerbaikan.Text = NulSID(rsNeedSupport!Perbaikan)
26                                txtProblem.Text = NulSID(rsNeedSupport!DetailProblem)
27                                txtRootCause.Text = NulSID(rsNeedSupport!RootCause)
28                                ckProblemSolve.value = IIf(rsNeedSupport!SOLVE = "YA", vbChecked, vbUnchecked)
29                                tglKirim.value = Format(rsNeedSupport!tglKirimNTL, "dd/mm/yyyy")
30                                tglRespon.value = Format(rsNeedSupport!tglResponNTL, "dd/mm/yyyy")
31                            End If
32                            frameKirimTL.Visible = True
33                            Call refreshGridSelesaiCuci
34                            Exit Sub
35                        End If
36                    End If
37                    rsTempSelesaiCuci.MoveNext
38                Loop
39            End If
40        End If
End Sub

Private Sub cmdRefresh_Click()
    Call refreshGridSelesaiCuci
End Sub

Private Sub cmdSimpanTidakKirim_Click()
    Dim boleh As Boolean
    boleh = True
    If NulSID(txtKeteranganTidakKirim.Text) = "" Then
        MsgBox "Keterangan harus diisi.", vbInformation
        Exit Sub
    End If
    sSQL = "UPDATE NEEDSUPPORTTL SET SOLVE = 'YA',kirimATL = 'TIDAK',infoTL = '" & NulSID(txtKeteranganTidakKirim.Text) & "',tglinfo = getdate(),userInfo = '" & gSA & "'   WHERE ID = " & idSimpan
    Call objO.InsertData(sSQL)
    Call refreshGridSelesaiCuci
    MsgBox "Data berhasil disimpan.", vbInformation
    frameTidakKirim.Visible = False
    Exit Sub
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
    asxPowerButton2.Enabled = True
End Sub

Private Sub Form_Activate()
'    With AdoStatus
'        .ConnectionString = adoCon
'        .CommandType = adCmdText
'        .RecordSource = "SELECT * FROM STATUSCASE WITH (NOLOCK) ORDER BY status"
'        .Refresh
'    End With
'    cbStatus.DataSource = AdoStatus
End Sub

Private Sub Form_Load()
    Call refreshGridSelesaiCuci
    'Call FetchCombo
End Sub



Private Sub Form_Resize()
    CRViewer1.Top = Shape2(2).Top + Shape2(2).Height '0
    CRViewer1.Left = 120
    If ScaleHeight > 100 Then
        CRViewer1.Height = ScaleHeight - Shape2(2).Height - 100
        CRViewer1.Width = ScaleWidth - 40
    End If
    
    'asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
End Sub

Private Sub listSelesaiCuci_AfterColUpdate(ByVal ColIndex As Integer)
    With listSelesaiCuci
        Select Case ColIndex
            Case 1
                If .Columns(ColIndex) Then
                    Pilih = Pilih + 1
                Else
                    Pilih = Pilih - 1
                End If
               
        End Select
        
    End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Call refreshGridAntriCuci
    ElseIf SSTab1.Tab = 1 Then
        Call refreshGridSelesaiCuci
    End If
End Sub

Private Sub hapusForm()
    Tgl1 = isnow
    Tgl2 = isnow
End Sub


Private Sub tblRefresh_Click()
    Tgl1 = isnow '
    Tgl2 = isnow
    Call FetchCombo
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean

    boleh = True
    If Nul(Trim(Tgl2)) = "" Then
        boleh = False
        MsgBox "Tanggal Akhir tidak boleh Kosong", vbInformation + vbOKOnly
    End If
    
    If boleh Then
        If Log.rsA = 1 Then
            Call CetakRSA
        Else
            Call Cetak
        End If
    End If
End Sub

Private Sub CetakRSA()
      Dim rsRepo As ADODB.Recordset
      Dim txt As String
      Dim keterangan As String
1        On Error GoTo Cetak_Error

2     Screen.MousePointer = vbHourglass
          
3         Set Report = New RptRepoSupportATLRSA
          
4         txt = "select b.FS_KD_DEALER,b.FS_NM_DEALER,'Maret' as bulan,C.Komponen,d.Problem,a.DetailProblem,"
5         txt = txt & " e.Kode_Kategory as model,b.no_polisi,a.tglInfo as tglPBT,tglinfoatl as tglResponATL,a.kirimNTL,a.tglKirimNTL,"
6         txt = txt & " a.tglResponNTL,a.tglsolved,a.tglpesanpart,a.tglPartTersedia,a.RootCauseatl,a.PerbaikanATL,"
7         txt = txt & " DATEDIFF(day,a.tglInfo,a.tglsolved) as LTRSA,DATEDIFF(day,a.tglKirimNTL,a.tglsolved) as LTNTL,"
8         txt = txt & " DATEDIFF(Day, a.tglpesanpart, a.tglPartTersedia) As LTPart"
9         txt = txt & " from NeedSupportTL a WITH (NOLOCK)"
10        txt = txt & " left join T_DEALER t with (nolock) on 'bengkel'+a.KodeDealer = t.DBASECBG and t.FS_LOG_DEL = ''"
11        txt = txt & " left join jobdta b WITH (NOLOCK) ON a.no_wo = b.no_wo and t.FS_KD_DEALER = b.FS_KD_DEALER"
12        txt = txt & " LEFT JOIN m_katkomponen C WITH (nolock) ON c.idkomponen = a.idkomponen"
13        txt = txt & " LEFT JOIN m_katproblem D WITH (nolock) ON D.idproblem = a.idproblem"
14        txt = txt & " LEFT JOIN SPK00..InKategory E with (nolock) on e.KodeKategory = b.MDL"
15        txt = txt & " where a.kirimATL = 'YA'"
16        txt = txt & " and a.tglInfo between '" & Format(Tgl1, "yyyy-mm-dd") & " 00:00:00' and '" & Format(Tgl2, "yyyy-mm-dd") & " 23:59:59'"
         
17        Set rsRepo = objO.FetchDataHO(txt)
          
18        Report.database.SetDataSource rsRepo
19        Report.txtRangeTanggal.SetText Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
           
20        sSQL = "select COUNT(no_wo) from NeedSupportTL WITH (NOLOCK) where kirimATL = 'YA' and tglInfo between '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00:00' and '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59:59'"
          
21        Report.txt1.SetText objO.FetchData_Double(sSQL)
           
22        sSQL = "select COUNT(no_wo) from NeedSupportTL WITH (NOLOCK) where kirimATL = 'YA' and kirimNTL = 'YA' and tglInfo between '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00:00' and '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59:59'"
23        Report.txt2.SetText objO.FetchData_Double(sSQL)
            
          Dim jumlahKasus As Double
          Dim jumlahSolve As Double
           
24        sSQL = "select COUNT(no_wo) from NeedSupportTL WITH (NOLOCK) where kirimATL = 'YA' and sumbertl <> 'PBT' and tglInfo between '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00:00' and '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59:59'" & keterangan
          
25        jumlahSolve = objO.FetchData_Double(sSQL)
26        sSQL = "select COUNT(no_wo) from NeedSupportTL WITH (NOLOCK) where kirimATL = 'YA' and tglInfo between '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00:00' and '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59:59'"
27        jumlahKasus = objO.FetchData_Double(sSQL)
28        If jumlahKasus > 0 Then
29            Report.txt3.SetText Round((jumlahSolve * 100 / jumlahKasus), 0) & "%"
30        Else
31            Report.txt3.SetText " 0%"
32        End If
          
33        sSQL = "select COUNT(no_wo) from NeedSupportTL WITH (NOLOCK) where kirimATL = 'YA' and sumbertl <> 'PBT' and tglInfo between '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00:00' and '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59:59'" & keterangan
          
34        jumlahSolve = objO.FetchData_Double(sSQL)
35        sSQL = "select COUNT(no_wo) from NeedSupportTL WITH (NOLOCK) where kirimATL = 'YA' and tglInfo between '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00:00' and '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59:59'"
36        jumlahKasus = objO.FetchData_Double(sSQL)
37        If jumlahKasus > 0 Then
38            Report.txt4.SetText Round((jumlahSolve / jumlahKasus), 0)
39        Else
40            Report.txt4.SetText " 0"
41        End If
          
42        With CRViewer1
43            .ReportSource = Report
44            .ViewReport
45            .Zoom 100
46            .Visible = True
47        End With
48        cmdTutup.Visible = True
49        cmdSettingPrinter.Visible = True
50        Screen.MousePointer = vbDefault

51       On Error GoTo 0
52       Exit Sub

Cetak_Error:

53        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure CetakRSA of Form frmRepoSupportTL")
End Sub

Private Sub Cetak()
      Dim rsRepo As ADODB.Recordset
      Dim txt As String
1        On Error GoTo Cetak_Error

2     Screen.MousePointer = vbHourglass
          
3         Set Report = New RptRepoSupportTL
          
4         txt = "select n.no_wo, '' as kategoryProblem , '' as kategoryKomponen, n.keterangan as problem, j.FMDL , j.NO_POLISI,"
5         txt = txt & " case j.FN_RETURN when 1 then 'RTJ' else 'Tidak' end as RTJ, j.DATE, n.tglEntry as tglPBT,'' as kirimPBT,''as tglKirimPBT,"
6         txt = txt & " '' as tglResponPBT, (select top 1 tanggal from tempprogress where flagp = 'F' and noorder = n.no_wo order by tanggal desc,pause desc) as tglSelesaiPerbaikan,'' as tglPesanPart,'' as tglPartTersedia,"
7         txt = txt & " n.infoTL as PenyebabMasalah,n.statusCase as PerbaikanDilakukan,"
8         txt = txt & " case when d.KUALITAS = 'P' and d.LAMA = 'P' and d.TEPAT = 'P' then 'FIR' else 'Non FIR' end"
9         txt = txt & "  as fir,DATEDIFF(DAY,n.tglEntry,c.tanggalstopfi)  as LTRSA"
10        txt = txt & " from NeedSupportTL n with (nolock) left join JOBDTA j with (nolock) on j.NO_WO = n.no_wo"
11        txt = txt & " left join ChassingForm c with (nolock) on c.NoWo = n.no_wo"
12        txt = txt & " left join DTLFU d with (nolock) on d.NO_WO = n.no_wo"
13        txt = txt & " where n.tglEntry between '" & Format(Tgl1, "MM/DD/YYYY") & " 00:00:00' and '" & Format(Tgl2, "MM/DD/YYYY") & " 23:59:59'"
          
14        Set rsRepo = objO.FetchData(txt)
          
15        Report.database.SetDataSource rsRepo
16        Report.txtPerusahaan.SetText gPerush
17        Report.txtAlamat.SetText gAlamat
18        Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
          
19        With CRViewer1
20            .ReportSource = Report
21            .ViewReport
22            .Zoom 100
23            .Visible = True
24        End With
25        cmdTutup.Visible = True
26        cmdSettingPrinter.Visible = True
27        Screen.MousePointer = vbDefault

28       On Error GoTo 0
29       Exit Sub

Cetak_Error:

30        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Cetak of Form frmRepoSupportTL")
End Sub


Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    Frame1.Visible = False
    asxPowerButton2.Enabled = True
End Sub
Private Sub Timer1_Timer()
   ' If SSTab1.Tab = 0 Then
   '     Call refreshGridAntriCuci
   ' ElseIf SSTab1.Tab = 1 Then
      ''Call refreshGridSelesaiCuci
    'End If
End Sub

Private Sub refreshCombo()
    adcProblem.ConnectionString = SQLConHO
    adcProblem.RecordSource = "select idProblem,Problem from m_katproblem with (nolock) where userDel = ''"
    cbKategoriProblem.DataSourceList = adcProblem
    cbKategoriProblem.DataFieldList = "idProblem"
    cbKategoriProblem.DataFieldToDisplay = "Problem"
    cbKategoriProblem.Columns(0).Width = 800
    cbKategoriProblem.Columns(1).Width = 3000
    cbKategoriProblem.Columns(0).Visible = False
    cbKategoriProblem.Columns(1).Caption = "Problem"
    
    adcKomponen.ConnectionString = SQLConHO
    adcKomponen.RecordSource = "select idKomponen,Komponen from m_katkomponen with (nolock) where userDel = ''"
    cbKategoriKomponen.DataSourceList = adcKomponen
    cbKategoriKomponen.DataFieldList = "idKomponen"
    cbKategoriKomponen.DataFieldToDisplay = "Komponen"
    cbKategoriKomponen.Columns(0).Width = 800
    cbKategoriKomponen.Columns(1).Width = 3000
    cbKategoriKomponen.Columns(0).Visible = False
    cbKategoriKomponen.Columns(1).Caption = "Komponen"
End Sub
