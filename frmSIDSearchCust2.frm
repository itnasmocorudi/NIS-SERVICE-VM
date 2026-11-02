VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "arprogbar.ocx"
Begin VB.Form frmSIDSearchCust 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   14550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   2760
      Left            =   60
      TabIndex        =   0
      Top             =   375
      Width           =   12510
      Begin VB.Frame Frame3 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   5670
         TabIndex        =   26
         Top             =   90
         Width           =   5460
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1785
            MaxLength       =   12
            TabIndex        =   27
            Top             =   45
            Width           =   1755
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   1785
            MaxLength       =   20
            TabIndex        =   28
            Top             =   375
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   1785
            MaxLength       =   25
            TabIndex        =   29
            Top             =   705
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Index           =   6
            Left            =   1785
            MaxLength       =   160
            TabIndex        =   30
            Top             =   1035
            Width           =   3225
         End
         Begin VB.TextBox text 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Index           =   5
            Left            =   1785
            MaxLength       =   13
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   1710
            Width           =   3630
         End
         Begin TDBDate6Ctl.TDBDate tglLahirx 
            Height          =   330
            Left            =   1785
            TabIndex        =   31
            Top             =   1365
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   582
            Calendar        =   "frmSIDSearchCust.frx":0000
            Caption         =   "frmSIDSearchCust.frx":012C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSIDSearchCust.frx":0198
            Keys            =   "frmSIDSearchCust.frx":01B6
            Spin            =   "frmSIDSearchCust.frx":0214
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
         Begin VB.Shape Shape1 
            BorderColor     =   &H000080FF&
            Height          =   2655
            Left            =   0
            Top             =   0
            Width           =   5460
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "No ID"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   44
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "No NPWP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   43
            Top             =   385
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Code Customer"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   42
            Top             =   715
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   41
            Top             =   1045
            Width           =   1620
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   40
            Top             =   1380
            Width           =   1635
         End
         Begin VB.Label Label19 
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
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   1720
            Width           =   1620
         End
         Begin VB.Label Label18 
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
            Height          =   300
            Left            =   1695
            TabIndex        =   38
            Top             =   60
            Width           =   165
         End
         Begin VB.Label Label17 
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
            Height          =   300
            Left            =   1695
            TabIndex        =   37
            Top             =   385
            Width           =   165
         End
         Begin VB.Label Label16 
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
            Height          =   300
            Left            =   1695
            TabIndex        =   36
            Top             =   715
            Width           =   165
         End
         Begin VB.Label Label15 
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
            Height          =   300
            Left            =   1695
            TabIndex        =   35
            Top             =   1050
            Width           =   165
         End
         Begin VB.Label Label14 
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
            Height          =   300
            Left            =   1695
            TabIndex        =   34
            Top             =   1380
            Width           =   165
         End
         Begin VB.Label Label13 
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
            Height          =   300
            Left            =   1695
            TabIndex        =   33
            Top             =   1720
            Width           =   165
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Height          =   2610
         Left            =   75
         TabIndex        =   1
         Top             =   105
         Width           =   5520
         Begin VB.TextBox text 
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
            Index           =   4
            Left            =   1860
            MaxLength       =   13
            TabIndex        =   12
            Top             =   1710
            Width           =   2970
         End
         Begin VB.TextBox text 
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
            Index           =   3
            Left            =   1860
            MaxLength       =   160
            TabIndex        =   11
            Top             =   1035
            Width           =   2970
         End
         Begin VB.TextBox text 
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
            Index           =   2
            Left            =   1860
            MaxLength       =   25
            TabIndex        =   10
            Top             =   705
            Width           =   2970
         End
         Begin VB.TextBox text 
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
            Index           =   1
            Left            =   1860
            MaxLength       =   20
            TabIndex        =   9
            Top             =   375
            Width           =   2970
         End
         Begin VB.TextBox text 
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
            Index           =   0
            Left            =   1860
            MaxLength       =   12
            TabIndex        =   3
            Top             =   45
            Width           =   1740
         End
         Begin TDBDate6Ctl.TDBDate tglLahir 
            Height          =   330
            Left            =   1860
            TabIndex        =   13
            Top             =   1365
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   582
            Calendar        =   "frmSIDSearchCust.frx":023C
            Caption         =   "frmSIDSearchCust.frx":0368
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmSIDSearchCust.frx":03D4
            Keys            =   "frmSIDSearchCust.frx":03F2
            Spin            =   "frmSIDSearchCust.frx":0450
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
         Begin ucXPButton.XPButton btnCari 
            Height          =   495
            Left            =   645
            TabIndex        =   21
            Top             =   2040
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
            MICON           =   "frmSIDSearchCust.frx":0478
            PICN            =   "frmSIDSearchCust.frx":0494
            PICH            =   "frmSIDSearchCust.frx":0A02
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton btnTambah 
            Height          =   495
            Left            =   3060
            TabIndex        =   22
            Top             =   2040
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
            MICON           =   "frmSIDSearchCust.frx":101F
            PICN            =   "frmSIDSearchCust.frx":103B
            PICH            =   "frmSIDSearchCust.frx":1487
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton btnKeluar 
            Height          =   495
            Left            =   4275
            TabIndex        =   23
            Top             =   2040
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
            MICON           =   "frmSIDSearchCust.frx":194F
            PICN            =   "frmSIDSearchCust.frx":196B
            PICH            =   "frmSIDSearchCust.frx":1E3F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton btnClear 
            Height          =   495
            Left            =   1845
            TabIndex        =   45
            Top             =   2040
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
            MICON           =   "frmSIDSearchCust.frx":238A
            PICN            =   "frmSIDSearchCust.frx":23A6
            PICH            =   "frmSIDSearchCust.frx":28F7
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
            Height          =   300
            Left            =   1725
            TabIndex        =   20
            Top             =   1720
            Width           =   165
         End
         Begin VB.Label Label11 
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
            Height          =   300
            Left            =   1725
            TabIndex        =   19
            Top             =   1380
            Width           =   165
         End
         Begin VB.Label Label10 
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
            Height          =   300
            Left            =   1725
            TabIndex        =   18
            Top             =   1045
            Width           =   165
         End
         Begin VB.Label Label9 
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
            Height          =   300
            Left            =   1725
            TabIndex        =   17
            Top             =   715
            Width           =   165
         End
         Begin VB.Label Label8 
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
            Height          =   300
            Left            =   1725
            TabIndex        =   16
            Top             =   385
            Width           =   165
         End
         Begin VB.Label Label7 
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
            Height          =   300
            Left            =   1725
            TabIndex        =   15
            Top             =   60
            Width           =   165
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Customer "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   8
            Top             =   1720
            Width           =   1620
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   7
            Top             =   1380
            Width           =   1635
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   1045
            Width           =   1620
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "No KTP / NPWP "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   715
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No Rangka "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   4
            Top             =   385
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Polisi"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   60
            Width           =   1305
         End
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H000080FF&
         Height          =   2655
         Left            =   0
         Top             =   90
         Width           =   5610
      End
   End
   Begin ARProgBarCtrl.ARProgressBar ARProgressBar1 
      Height          =   210
      Left            =   60
      TabIndex        =   24
      Top             =   3135
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   370
      Value           =   0
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0 %"
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   5610
      Left            =   60
      TabIndex        =   25
      Top             =   3360
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   9895
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Detail"
      Columns(0).DataField=   "Detail"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Proses"
      Columns(1).DataField=   "Proses"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode Customer"
      Columns(2).DataField=   "IDCust"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Title"
      Columns(3).DataField=   "Title"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Customer"
      Columns(4).DataField=   "NamaCust"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Tgl Lahir"
      Columns(5).DataField=   "TglLahir"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Agama"
      Columns(6).DataField=   "Agama"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Pekerjaan"
      Columns(7).DataField=   "Pekerjaan"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Propinsi"
      Columns(8).DataField=   "Propinsi"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Kecamatan"
      Columns(9).DataField=   "Kecamatan"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Kelurahan"
      Columns(10).DataField=   "Kelurahan"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Kode Pos"
      Columns(11).DataField=   "KodePos"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "No Polisi"
      Columns(12).DataField=   "Nopolisi"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Alamat"
      Columns(13).DataField=   "Alamat"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Kota"
      Columns(14).DataField=   "Kota"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "RT"
      Columns(15).DataField=   "RT"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "RW"
      Columns(16).DataField=   "RW"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Kode Telp"
      Columns(17).DataField=   "KodeTelp"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Telp"
      Columns(18).DataField=   "Telp"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "HP"
      Columns(19).DataField=   "HP1"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "HP2"
      Columns(20).DataField=   "HP2"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Email"
      Columns(21).DataField=   "Email"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "No. ID/NPWP"
      Columns(22).DataField=   "NoID"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Hobi"
      Columns(23).DataField=   "Hobi"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Olah Raga"
      Columns(24).DataField=   "Olahraga"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "Suku"
      Columns(25).DataField=   "Suku"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Preferensi"
      Columns(26).DataField=   "Preferensi"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "No Polish"
      Columns(27).DataField=   "NoPolis"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "Tgl Awal Polish"
      Columns(28).DataField=   "TglAwalPolis"
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "Tgl Akhir Polish"
      Columns(29).DataField=   "TglAkhirPolis"
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "No Rangka"
      Columns(30).DataField=   "NoRangka"
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "No Engine"
      Columns(31).DataField=   "NoEngine"
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "No Key"
      Columns(32).DataField=   "NoKey"
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "Model"
      Columns(33).DataField=   "Mdl"
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "Class"
      Columns(34).DataField=   "Class"
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).Caption=   "Full Model"
      Columns(35).DataField=   "Fmdl"
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).Caption=   "Color"
      Columns(36).DataField=   "Color"
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).Caption=   "Kode Dealer"
      Columns(37).DataField=   "KdDlrPenjual"
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).Caption=   "Nama Dealer"
      Columns(38).DataField=   "NmDlrPenjual"
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).Caption=   "DDate"
      Columns(39).DataField=   "DDate"
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).Caption=   "NONPWP"
      Columns(40).DataField=   "NONPWP"
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).Caption=   "NOKTP"
      Columns(41).DataField=   "NOKTP"
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   42
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).SizeMode=   2
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=42"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0).ButtonText=1"
      Splits(0)._ColumnProps(7)=   "Column(0).ButtonAlways=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=1773"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1667"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(12)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(1).ButtonText=1"
      Splits(0)._ColumnProps(15)=   "Column(1).ButtonAlways=1"
      Splits(0)._ColumnProps(16)=   "Column(1).ButtonHeader=1"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2752"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=8708"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=4657"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=4551"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=8708"
      Splits(0)._ColumnProps(33)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=2275"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=2170"
      Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=8708"
      Splits(0)._ColumnProps(39)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(41)=   "Column(6).Width=159"
      Splits(0)._ColumnProps(42)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._WidthInPix=53"
      Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=8708"
      Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(47)=   "Column(7).Width=3360"
      Splits(0)._ColumnProps(48)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._WidthInPix=3254"
      Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=8708"
      Splits(0)._ColumnProps(51)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(53)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(54)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(56)=   "Column(8)._ColStyle=8708"
      Splits(0)._ColumnProps(57)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(58)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(59)=   "Column(9).Width=2646"
      Splits(0)._ColumnProps(60)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(9)._WidthInPix=2540"
      Splits(0)._ColumnProps(62)=   "Column(9)._ColStyle=8708"
      Splits(0)._ColumnProps(63)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(64)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(65)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(10)._ColStyle=8708"
      Splits(0)._ColumnProps(69)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(71)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(72)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(74)=   "Column(11)._ColStyle=8708"
      Splits(0)._ColumnProps(75)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(76)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(77)=   "Column(12).Width=1746"
      Splits(0)._ColumnProps(78)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(12)._WidthInPix=1640"
      Splits(0)._ColumnProps(80)=   "Column(12)._ColStyle=8708"
      Splits(0)._ColumnProps(81)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(83)=   "Column(13).Width=6271"
      Splits(0)._ColumnProps(84)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(13)._WidthInPix=6165"
      Splits(0)._ColumnProps(86)=   "Column(13)._ColStyle=8708"
      Splits(0)._ColumnProps(87)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(88)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(89)=   "Column(14).Width=2963"
      Splits(0)._ColumnProps(90)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(14)._WidthInPix=2858"
      Splits(0)._ColumnProps(92)=   "Column(14)._ColStyle=8708"
      Splits(0)._ColumnProps(93)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(95)=   "Column(15).Width=7064"
      Splits(0)._ColumnProps(96)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(15)._WidthInPix=6959"
      Splits(0)._ColumnProps(98)=   "Column(15)._ColStyle=8708"
      Splits(0)._ColumnProps(99)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(100)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(101)=   "Column(16).Width=2461"
      Splits(0)._ColumnProps(102)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(16)._WidthInPix=2355"
      Splits(0)._ColumnProps(104)=   "Column(16)._ColStyle=8708"
      Splits(0)._ColumnProps(105)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(106)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(107)=   "Column(17).Width=2646"
      Splits(0)._ColumnProps(108)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(109)=   "Column(17)._WidthInPix=2540"
      Splits(0)._ColumnProps(110)=   "Column(17)._ColStyle=8708"
      Splits(0)._ColumnProps(111)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(112)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(113)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(114)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(116)=   "Column(18)._ColStyle=8708"
      Splits(0)._ColumnProps(117)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(118)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(119)=   "Column(19).Width=3784"
      Splits(0)._ColumnProps(120)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(121)=   "Column(19)._WidthInPix=3678"
      Splits(0)._ColumnProps(122)=   "Column(19)._ColStyle=8708"
      Splits(0)._ColumnProps(123)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(124)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(125)=   "Column(20).Width=2646"
      Splits(0)._ColumnProps(126)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(20)._WidthInPix=2540"
      Splits(0)._ColumnProps(128)=   "Column(20)._ColStyle=8708"
      Splits(0)._ColumnProps(129)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(130)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(131)=   "Column(21).Width=3281"
      Splits(0)._ColumnProps(132)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(21)._WidthInPix=3175"
      Splits(0)._ColumnProps(134)=   "Column(21)._ColStyle=8708"
      Splits(0)._ColumnProps(135)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(136)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(137)=   "Column(22).Width=4789"
      Splits(0)._ColumnProps(138)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(22)._WidthInPix=4683"
      Splits(0)._ColumnProps(140)=   "Column(22)._ColStyle=8708"
      Splits(0)._ColumnProps(141)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(142)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(143)=   "Column(23).Width=873"
      Splits(0)._ColumnProps(144)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(23)._WidthInPix=767"
      Splits(0)._ColumnProps(146)=   "Column(23)._ColStyle=8708"
      Splits(0)._ColumnProps(147)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(148)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(149)=   "Column(24).Width=1773"
      Splits(0)._ColumnProps(150)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(24)._WidthInPix=1667"
      Splits(0)._ColumnProps(152)=   "Column(24)._ColStyle=8708"
      Splits(0)._ColumnProps(153)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(154)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(155)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(156)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(157)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(158)=   "Column(25)._ColStyle=8708"
      Splits(0)._ColumnProps(159)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(160)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(161)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(162)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(164)=   "Column(26)._ColStyle=8196"
      Splits(0)._ColumnProps(165)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(166)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(167)=   "Column(27).Width=3519"
      Splits(0)._ColumnProps(168)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(169)=   "Column(27)._WidthInPix=3413"
      Splits(0)._ColumnProps(170)=   "Column(27)._ColStyle=8708"
      Splits(0)._ColumnProps(171)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(172)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(173)=   "Column(28).Width=3519"
      Splits(0)._ColumnProps(174)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(28)._WidthInPix=3413"
      Splits(0)._ColumnProps(176)=   "Column(28)._ColStyle=8708"
      Splits(0)._ColumnProps(177)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(178)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(179)=   "Column(29).Width=3519"
      Splits(0)._ColumnProps(180)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(181)=   "Column(29)._WidthInPix=3413"
      Splits(0)._ColumnProps(182)=   "Column(29)._ColStyle=8708"
      Splits(0)._ColumnProps(183)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(184)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(185)=   "Column(30).Width=4366"
      Splits(0)._ColumnProps(186)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(30)._WidthInPix=4260"
      Splits(0)._ColumnProps(188)=   "Column(30)._ColStyle=8708"
      Splits(0)._ColumnProps(189)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(190)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(191)=   "Column(31).Width=873"
      Splits(0)._ColumnProps(192)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(193)=   "Column(31)._WidthInPix=767"
      Splits(0)._ColumnProps(194)=   "Column(31)._ColStyle=8708"
      Splits(0)._ColumnProps(195)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(196)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(197)=   "Column(32).Width=3519"
      Splits(0)._ColumnProps(198)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(32)._WidthInPix=3413"
      Splits(0)._ColumnProps(200)=   "Column(32)._ColStyle=8708"
      Splits(0)._ColumnProps(201)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(202)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(203)=   "Column(33).Width=3519"
      Splits(0)._ColumnProps(204)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(205)=   "Column(33)._WidthInPix=3413"
      Splits(0)._ColumnProps(206)=   "Column(33)._ColStyle=8708"
      Splits(0)._ColumnProps(207)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(208)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(209)=   "Column(34).Width=5292"
      Splits(0)._ColumnProps(210)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(34)._WidthInPix=5186"
      Splits(0)._ColumnProps(212)=   "Column(34)._ColStyle=8708"
      Splits(0)._ColumnProps(213)=   "Column(34).Visible=0"
      Splits(0)._ColumnProps(214)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(215)=   "Column(35).Width=3281"
      Splits(0)._ColumnProps(216)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(217)=   "Column(35)._WidthInPix=3175"
      Splits(0)._ColumnProps(218)=   "Column(35)._ColStyle=8708"
      Splits(0)._ColumnProps(219)=   "Column(35).Visible=0"
      Splits(0)._ColumnProps(220)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(221)=   "Column(36).Width=3281"
      Splits(0)._ColumnProps(222)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(223)=   "Column(36)._WidthInPix=3175"
      Splits(0)._ColumnProps(224)=   "Column(36)._ColStyle=8708"
      Splits(0)._ColumnProps(225)=   "Column(36).Visible=0"
      Splits(0)._ColumnProps(226)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(227)=   "Column(37).Width=5292"
      Splits(0)._ColumnProps(228)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(229)=   "Column(37)._WidthInPix=5186"
      Splits(0)._ColumnProps(230)=   "Column(37)._ColStyle=8708"
      Splits(0)._ColumnProps(231)=   "Column(37).Visible=0"
      Splits(0)._ColumnProps(232)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(233)=   "Column(38).Width=3281"
      Splits(0)._ColumnProps(234)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(235)=   "Column(38)._WidthInPix=3175"
      Splits(0)._ColumnProps(236)=   "Column(38)._ColStyle=8708"
      Splits(0)._ColumnProps(237)=   "Column(38).Visible=0"
      Splits(0)._ColumnProps(238)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(239)=   "Column(39).Width=3281"
      Splits(0)._ColumnProps(240)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(241)=   "Column(39)._WidthInPix=3175"
      Splits(0)._ColumnProps(242)=   "Column(39)._ColStyle=8708"
      Splits(0)._ColumnProps(243)=   "Column(39).Visible=0"
      Splits(0)._ColumnProps(244)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(245)=   "Column(40).Width=3281"
      Splits(0)._ColumnProps(246)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(40)._WidthInPix=3175"
      Splits(0)._ColumnProps(248)=   "Column(40).Visible=0"
      Splits(0)._ColumnProps(249)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(250)=   "Column(41).Width=3281"
      Splits(0)._ColumnProps(251)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(252)=   "Column(41)._WidthInPix=3175"
      Splits(0)._ColumnProps(253)=   "Column(41).Visible=0"
      Splits(0)._ColumnProps(254)=   "Column(41).Order=42"
      Splits(1)._UserFlags=   0
      Splits(1).AllowSizing=   -1  'True
      Splits(1).RecordSelectorWidth=   688
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).AllowColSelect=   0   'False
      Splits(1).FetchRowStyle=   -1  'True
      Splits(1).AlternatingRowStyle=   -1  'True
      Splits(1).DividerColor=   14215660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=42"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(7)=   "Column(0).ButtonText=1"
      Splits(1)._ColumnProps(8)=   "Column(0).ButtonAlways=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=1773"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1667"
      Splits(1)._ColumnProps(12)=   "Column(1)._ColStyle=513"
      Splits(1)._ColumnProps(13)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(14)=   "Column(1).Button=1"
      Splits(1)._ColumnProps(15)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(16)=   "Column(1).ButtonText=1"
      Splits(1)._ColumnProps(17)=   "Column(1).ButtonAlways=1"
      Splits(1)._ColumnProps(18)=   "Column(1).ButtonHeader=1"
      Splits(1)._ColumnProps(19)=   "Column(2).Width=2752"
      Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(22)=   "Column(2)._ColStyle=8708"
      Splits(1)._ColumnProps(23)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(24)=   "Column(3).Width=3281"
      Splits(1)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(26)=   "Column(3)._WidthInPix=3175"
      Splits(1)._ColumnProps(27)=   "Column(3)._ColStyle=8708"
      Splits(1)._ColumnProps(28)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(30)=   "Column(4).Width=4657"
      Splits(1)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(32)=   "Column(4)._WidthInPix=4551"
      Splits(1)._ColumnProps(33)=   "Column(4)._ColStyle=8708"
      Splits(1)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(35)=   "Column(5).Width=2275"
      Splits(1)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(37)=   "Column(5)._WidthInPix=2170"
      Splits(1)._ColumnProps(38)=   "Column(5)._ColStyle=8708"
      Splits(1)._ColumnProps(39)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(40)=   "Column(6).Width=159"
      Splits(1)._ColumnProps(41)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(42)=   "Column(6)._WidthInPix=53"
      Splits(1)._ColumnProps(43)=   "Column(6)._ColStyle=8708"
      Splits(1)._ColumnProps(44)=   "Column(6).Visible=0"
      Splits(1)._ColumnProps(45)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(46)=   "Column(7).Width=3360"
      Splits(1)._ColumnProps(47)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(48)=   "Column(7)._WidthInPix=3254"
      Splits(1)._ColumnProps(49)=   "Column(7)._ColStyle=8708"
      Splits(1)._ColumnProps(50)=   "Column(7).Visible=0"
      Splits(1)._ColumnProps(51)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(52)=   "Column(8).Width=3281"
      Splits(1)._ColumnProps(53)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(54)=   "Column(8)._WidthInPix=3175"
      Splits(1)._ColumnProps(55)=   "Column(8)._ColStyle=8708"
      Splits(1)._ColumnProps(56)=   "Column(8).Visible=0"
      Splits(1)._ColumnProps(57)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(58)=   "Column(9).Width=2646"
      Splits(1)._ColumnProps(59)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(60)=   "Column(9)._WidthInPix=2540"
      Splits(1)._ColumnProps(61)=   "Column(9)._ColStyle=8708"
      Splits(1)._ColumnProps(62)=   "Column(9).Visible=0"
      Splits(1)._ColumnProps(63)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(64)=   "Column(10).Width=3281"
      Splits(1)._ColumnProps(65)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(66)=   "Column(10)._WidthInPix=3175"
      Splits(1)._ColumnProps(67)=   "Column(10)._ColStyle=8708"
      Splits(1)._ColumnProps(68)=   "Column(10).Visible=0"
      Splits(1)._ColumnProps(69)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(70)=   "Column(11).Width=3281"
      Splits(1)._ColumnProps(71)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(72)=   "Column(11)._WidthInPix=3175"
      Splits(1)._ColumnProps(73)=   "Column(11)._ColStyle=8708"
      Splits(1)._ColumnProps(74)=   "Column(11).Visible=0"
      Splits(1)._ColumnProps(75)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(76)=   "Column(12).Width=1746"
      Splits(1)._ColumnProps(77)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(78)=   "Column(12)._WidthInPix=1640"
      Splits(1)._ColumnProps(79)=   "Column(12)._ColStyle=8708"
      Splits(1)._ColumnProps(80)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(81)=   "Column(13).Width=6271"
      Splits(1)._ColumnProps(82)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(13)._WidthInPix=6165"
      Splits(1)._ColumnProps(84)=   "Column(13)._ColStyle=8708"
      Splits(1)._ColumnProps(85)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(86)=   "Column(14).Width=2963"
      Splits(1)._ColumnProps(87)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(88)=   "Column(14)._WidthInPix=2858"
      Splits(1)._ColumnProps(89)=   "Column(14)._ColStyle=8708"
      Splits(1)._ColumnProps(90)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(91)=   "Column(15).Width=7064"
      Splits(1)._ColumnProps(92)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(93)=   "Column(15)._WidthInPix=6959"
      Splits(1)._ColumnProps(94)=   "Column(15)._ColStyle=8708"
      Splits(1)._ColumnProps(95)=   "Column(15).Visible=0"
      Splits(1)._ColumnProps(96)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(97)=   "Column(16).Width=2461"
      Splits(1)._ColumnProps(98)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(99)=   "Column(16)._WidthInPix=2355"
      Splits(1)._ColumnProps(100)=   "Column(16)._ColStyle=8708"
      Splits(1)._ColumnProps(101)=   "Column(16).Visible=0"
      Splits(1)._ColumnProps(102)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(103)=   "Column(17).Width=2646"
      Splits(1)._ColumnProps(104)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(105)=   "Column(17)._WidthInPix=2540"
      Splits(1)._ColumnProps(106)=   "Column(17)._ColStyle=8708"
      Splits(1)._ColumnProps(107)=   "Column(17).Visible=0"
      Splits(1)._ColumnProps(108)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(109)=   "Column(18).Width=3281"
      Splits(1)._ColumnProps(110)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(18)._WidthInPix=3175"
      Splits(1)._ColumnProps(112)=   "Column(18)._ColStyle=8708"
      Splits(1)._ColumnProps(113)=   "Column(18).Visible=0"
      Splits(1)._ColumnProps(114)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(115)=   "Column(19).Width=3784"
      Splits(1)._ColumnProps(116)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(117)=   "Column(19)._WidthInPix=3678"
      Splits(1)._ColumnProps(118)=   "Column(19)._ColStyle=8708"
      Splits(1)._ColumnProps(119)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(120)=   "Column(20).Width=2646"
      Splits(1)._ColumnProps(121)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(122)=   "Column(20)._WidthInPix=2540"
      Splits(1)._ColumnProps(123)=   "Column(20)._ColStyle=8708"
      Splits(1)._ColumnProps(124)=   "Column(20).Visible=0"
      Splits(1)._ColumnProps(125)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(126)=   "Column(21).Width=3281"
      Splits(1)._ColumnProps(127)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(128)=   "Column(21)._WidthInPix=3175"
      Splits(1)._ColumnProps(129)=   "Column(21)._ColStyle=8708"
      Splits(1)._ColumnProps(130)=   "Column(21).Visible=0"
      Splits(1)._ColumnProps(131)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(132)=   "Column(22).Width=4789"
      Splits(1)._ColumnProps(133)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(134)=   "Column(22)._WidthInPix=4683"
      Splits(1)._ColumnProps(135)=   "Column(22)._ColStyle=8708"
      Splits(1)._ColumnProps(136)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(137)=   "Column(23).Width=873"
      Splits(1)._ColumnProps(138)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(139)=   "Column(23)._WidthInPix=767"
      Splits(1)._ColumnProps(140)=   "Column(23)._ColStyle=8708"
      Splits(1)._ColumnProps(141)=   "Column(23).Visible=0"
      Splits(1)._ColumnProps(142)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(143)=   "Column(24).Width=1773"
      Splits(1)._ColumnProps(144)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(145)=   "Column(24)._WidthInPix=1667"
      Splits(1)._ColumnProps(146)=   "Column(24)._ColStyle=8708"
      Splits(1)._ColumnProps(147)=   "Column(24).Visible=0"
      Splits(1)._ColumnProps(148)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(149)=   "Column(25).Width=3281"
      Splits(1)._ColumnProps(150)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(151)=   "Column(25)._WidthInPix=3175"
      Splits(1)._ColumnProps(152)=   "Column(25)._ColStyle=8708"
      Splits(1)._ColumnProps(153)=   "Column(25).Visible=0"
      Splits(1)._ColumnProps(154)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(155)=   "Column(26).Width=3281"
      Splits(1)._ColumnProps(156)=   "Column(26).DividerColor=0"
      Splits(1)._ColumnProps(157)=   "Column(26)._WidthInPix=3175"
      Splits(1)._ColumnProps(158)=   "Column(26)._ColStyle=8196"
      Splits(1)._ColumnProps(159)=   "Column(26).Visible=0"
      Splits(1)._ColumnProps(160)=   "Column(26).Order=27"
      Splits(1)._ColumnProps(161)=   "Column(27).Width=3519"
      Splits(1)._ColumnProps(162)=   "Column(27).DividerColor=0"
      Splits(1)._ColumnProps(163)=   "Column(27)._WidthInPix=3413"
      Splits(1)._ColumnProps(164)=   "Column(27)._ColStyle=8708"
      Splits(1)._ColumnProps(165)=   "Column(27).Visible=0"
      Splits(1)._ColumnProps(166)=   "Column(27).Order=28"
      Splits(1)._ColumnProps(167)=   "Column(28).Width=3519"
      Splits(1)._ColumnProps(168)=   "Column(28).DividerColor=0"
      Splits(1)._ColumnProps(169)=   "Column(28)._WidthInPix=3413"
      Splits(1)._ColumnProps(170)=   "Column(28)._ColStyle=8708"
      Splits(1)._ColumnProps(171)=   "Column(28).Visible=0"
      Splits(1)._ColumnProps(172)=   "Column(28).Order=29"
      Splits(1)._ColumnProps(173)=   "Column(29).Width=3519"
      Splits(1)._ColumnProps(174)=   "Column(29).DividerColor=0"
      Splits(1)._ColumnProps(175)=   "Column(29)._WidthInPix=3413"
      Splits(1)._ColumnProps(176)=   "Column(29)._ColStyle=8708"
      Splits(1)._ColumnProps(177)=   "Column(29).Visible=0"
      Splits(1)._ColumnProps(178)=   "Column(29).Order=30"
      Splits(1)._ColumnProps(179)=   "Column(30).Width=4366"
      Splits(1)._ColumnProps(180)=   "Column(30).DividerColor=0"
      Splits(1)._ColumnProps(181)=   "Column(30)._WidthInPix=4260"
      Splits(1)._ColumnProps(182)=   "Column(30)._ColStyle=8708"
      Splits(1)._ColumnProps(183)=   "Column(30).Order=31"
      Splits(1)._ColumnProps(184)=   "Column(31).Width=873"
      Splits(1)._ColumnProps(185)=   "Column(31).DividerColor=0"
      Splits(1)._ColumnProps(186)=   "Column(31)._WidthInPix=767"
      Splits(1)._ColumnProps(187)=   "Column(31)._ColStyle=8708"
      Splits(1)._ColumnProps(188)=   "Column(31).Visible=0"
      Splits(1)._ColumnProps(189)=   "Column(31).Order=32"
      Splits(1)._ColumnProps(190)=   "Column(32).Width=3519"
      Splits(1)._ColumnProps(191)=   "Column(32).DividerColor=0"
      Splits(1)._ColumnProps(192)=   "Column(32)._WidthInPix=3413"
      Splits(1)._ColumnProps(193)=   "Column(32)._ColStyle=8708"
      Splits(1)._ColumnProps(194)=   "Column(32).Visible=0"
      Splits(1)._ColumnProps(195)=   "Column(32).Order=33"
      Splits(1)._ColumnProps(196)=   "Column(33).Width=3519"
      Splits(1)._ColumnProps(197)=   "Column(33).DividerColor=0"
      Splits(1)._ColumnProps(198)=   "Column(33)._WidthInPix=3413"
      Splits(1)._ColumnProps(199)=   "Column(33)._ColStyle=8708"
      Splits(1)._ColumnProps(200)=   "Column(33).Visible=0"
      Splits(1)._ColumnProps(201)=   "Column(33).Order=34"
      Splits(1)._ColumnProps(202)=   "Column(34).Width=5292"
      Splits(1)._ColumnProps(203)=   "Column(34).DividerColor=0"
      Splits(1)._ColumnProps(204)=   "Column(34)._WidthInPix=5186"
      Splits(1)._ColumnProps(205)=   "Column(34)._ColStyle=8708"
      Splits(1)._ColumnProps(206)=   "Column(34).Visible=0"
      Splits(1)._ColumnProps(207)=   "Column(34).Order=35"
      Splits(1)._ColumnProps(208)=   "Column(35).Width=3281"
      Splits(1)._ColumnProps(209)=   "Column(35).DividerColor=0"
      Splits(1)._ColumnProps(210)=   "Column(35)._WidthInPix=3175"
      Splits(1)._ColumnProps(211)=   "Column(35)._ColStyle=8708"
      Splits(1)._ColumnProps(212)=   "Column(35).Visible=0"
      Splits(1)._ColumnProps(213)=   "Column(35).Order=36"
      Splits(1)._ColumnProps(214)=   "Column(36).Width=3281"
      Splits(1)._ColumnProps(215)=   "Column(36).DividerColor=0"
      Splits(1)._ColumnProps(216)=   "Column(36)._WidthInPix=3175"
      Splits(1)._ColumnProps(217)=   "Column(36)._ColStyle=8708"
      Splits(1)._ColumnProps(218)=   "Column(36).Visible=0"
      Splits(1)._ColumnProps(219)=   "Column(36).Order=37"
      Splits(1)._ColumnProps(220)=   "Column(37).Width=5292"
      Splits(1)._ColumnProps(221)=   "Column(37).DividerColor=0"
      Splits(1)._ColumnProps(222)=   "Column(37)._WidthInPix=5186"
      Splits(1)._ColumnProps(223)=   "Column(37)._ColStyle=8708"
      Splits(1)._ColumnProps(224)=   "Column(37).Visible=0"
      Splits(1)._ColumnProps(225)=   "Column(37).Order=38"
      Splits(1)._ColumnProps(226)=   "Column(38).Width=3281"
      Splits(1)._ColumnProps(227)=   "Column(38).DividerColor=0"
      Splits(1)._ColumnProps(228)=   "Column(38)._WidthInPix=3175"
      Splits(1)._ColumnProps(229)=   "Column(38)._ColStyle=8708"
      Splits(1)._ColumnProps(230)=   "Column(38).Visible=0"
      Splits(1)._ColumnProps(231)=   "Column(38).Order=39"
      Splits(1)._ColumnProps(232)=   "Column(39).Width=3281"
      Splits(1)._ColumnProps(233)=   "Column(39).DividerColor=0"
      Splits(1)._ColumnProps(234)=   "Column(39)._WidthInPix=3175"
      Splits(1)._ColumnProps(235)=   "Column(39)._ColStyle=8708"
      Splits(1)._ColumnProps(236)=   "Column(39).Visible=0"
      Splits(1)._ColumnProps(237)=   "Column(39).Order=40"
      Splits(1)._ColumnProps(238)=   "Column(40).Width=3281"
      Splits(1)._ColumnProps(239)=   "Column(40).DividerColor=0"
      Splits(1)._ColumnProps(240)=   "Column(40)._WidthInPix=3175"
      Splits(1)._ColumnProps(241)=   "Column(40).Visible=0"
      Splits(1)._ColumnProps(242)=   "Column(40).Order=41"
      Splits(1)._ColumnProps(243)=   "Column(41).Width=3281"
      Splits(1)._ColumnProps(244)=   "Column(41).DividerColor=0"
      Splits(1)._ColumnProps(245)=   "Column(41)._WidthInPix=3175"
      Splits(1)._ColumnProps(246)=   "Column(41).Visible=0"
      Splits(1)._ColumnProps(247)=   "Column(41).Order=42"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14215660
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=71,.parent=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(25)  =   ":id=71,.strikethrough=0,.charset=0"
      _StyleDefs(26)  =   ":id=71,.fontname=Verdana"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=136,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=72,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=73,.parent=3,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(30)  =   ":id=73,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(31)  =   ":id=73,.fontname=Verdana"
      _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=74,.parent=5"
      _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=84,.parent=6"
      _StyleDefs(34)  =   "Splits(0).EditorStyle:id=83,.parent=7"
      _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=86,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=135,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=137,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=138,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=190,.parent=71,.alignment=2,.fgcolor=&H0&"
      _StyleDefs(41)  =   ":id=190,.locked=0"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=187,.parent=72,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=188,.parent=73"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=189,.parent=83"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=194,.parent=71,.alignment=2,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=191,.parent=72,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=192,.parent=73"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=193,.parent=83"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=198,.parent=71,.locked=-1"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=195,.parent=72,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=196,.parent=73"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=197,.parent=83"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=222,.parent=71,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=219,.parent=72,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=220,.parent=73"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=221,.parent=83"
      _StyleDefs(57)  =   "Splits(0).Columns(4).Style:id=226,.parent=71,.locked=-1"
      _StyleDefs(58)  =   "Splits(0).Columns(4).HeadingStyle:id=223,.parent=72,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(4).FooterStyle:id=224,.parent=73"
      _StyleDefs(60)  =   "Splits(0).Columns(4).EditorStyle:id=225,.parent=83"
      _StyleDefs(61)  =   "Splits(0).Columns(5).Style:id=230,.parent=71,.locked=-1"
      _StyleDefs(62)  =   "Splits(0).Columns(5).HeadingStyle:id=227,.parent=72,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(5).FooterStyle:id=228,.parent=73"
      _StyleDefs(64)  =   "Splits(0).Columns(5).EditorStyle:id=229,.parent=83"
      _StyleDefs(65)  =   "Splits(0).Columns(6).Style:id=234,.parent=71,.locked=-1"
      _StyleDefs(66)  =   "Splits(0).Columns(6).HeadingStyle:id=231,.parent=72,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(6).FooterStyle:id=232,.parent=73"
      _StyleDefs(68)  =   "Splits(0).Columns(6).EditorStyle:id=233,.parent=83"
      _StyleDefs(69)  =   "Splits(0).Columns(7).Style:id=238,.parent=71,.locked=-1"
      _StyleDefs(70)  =   "Splits(0).Columns(7).HeadingStyle:id=235,.parent=72,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(7).FooterStyle:id=236,.parent=73"
      _StyleDefs(72)  =   "Splits(0).Columns(7).EditorStyle:id=237,.parent=83"
      _StyleDefs(73)  =   "Splits(0).Columns(8).Style:id=242,.parent=71,.locked=-1"
      _StyleDefs(74)  =   "Splits(0).Columns(8).HeadingStyle:id=239,.parent=72,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(8).FooterStyle:id=240,.parent=73"
      _StyleDefs(76)  =   "Splits(0).Columns(8).EditorStyle:id=241,.parent=83"
      _StyleDefs(77)  =   "Splits(0).Columns(9).Style:id=246,.parent=71,.locked=-1"
      _StyleDefs(78)  =   "Splits(0).Columns(9).HeadingStyle:id=243,.parent=72,.alignment=2"
      _StyleDefs(79)  =   "Splits(0).Columns(9).FooterStyle:id=244,.parent=73"
      _StyleDefs(80)  =   "Splits(0).Columns(9).EditorStyle:id=245,.parent=83"
      _StyleDefs(81)  =   "Splits(0).Columns(10).Style:id=250,.parent=71,.locked=-1"
      _StyleDefs(82)  =   "Splits(0).Columns(10).HeadingStyle:id=247,.parent=72,.alignment=2"
      _StyleDefs(83)  =   "Splits(0).Columns(10).FooterStyle:id=248,.parent=73"
      _StyleDefs(84)  =   "Splits(0).Columns(10).EditorStyle:id=249,.parent=83"
      _StyleDefs(85)  =   "Splits(0).Columns(11).Style:id=254,.parent=71,.locked=-1"
      _StyleDefs(86)  =   "Splits(0).Columns(11).HeadingStyle:id=251,.parent=72,.alignment=2"
      _StyleDefs(87)  =   "Splits(0).Columns(11).FooterStyle:id=252,.parent=73"
      _StyleDefs(88)  =   "Splits(0).Columns(11).EditorStyle:id=253,.parent=83"
      _StyleDefs(89)  =   "Splits(0).Columns(12).Style:id=258,.parent=71,.locked=-1"
      _StyleDefs(90)  =   "Splits(0).Columns(12).HeadingStyle:id=255,.parent=72,.alignment=2"
      _StyleDefs(91)  =   "Splits(0).Columns(12).FooterStyle:id=256,.parent=73"
      _StyleDefs(92)  =   "Splits(0).Columns(12).EditorStyle:id=257,.parent=83"
      _StyleDefs(93)  =   "Splits(0).Columns(13).Style:id=262,.parent=71,.locked=-1"
      _StyleDefs(94)  =   "Splits(0).Columns(13).HeadingStyle:id=259,.parent=72,.alignment=2"
      _StyleDefs(95)  =   "Splits(0).Columns(13).FooterStyle:id=260,.parent=73"
      _StyleDefs(96)  =   "Splits(0).Columns(13).EditorStyle:id=261,.parent=83"
      _StyleDefs(97)  =   "Splits(0).Columns(14).Style:id=266,.parent=71,.locked=-1"
      _StyleDefs(98)  =   "Splits(0).Columns(14).HeadingStyle:id=263,.parent=72,.alignment=2"
      _StyleDefs(99)  =   "Splits(0).Columns(14).FooterStyle:id=264,.parent=73"
      _StyleDefs(100) =   "Splits(0).Columns(14).EditorStyle:id=265,.parent=83"
      _StyleDefs(101) =   "Splits(0).Columns(15).Style:id=270,.parent=71,.locked=-1"
      _StyleDefs(102) =   "Splits(0).Columns(15).HeadingStyle:id=267,.parent=72,.alignment=2"
      _StyleDefs(103) =   "Splits(0).Columns(15).FooterStyle:id=268,.parent=73"
      _StyleDefs(104) =   "Splits(0).Columns(15).EditorStyle:id=269,.parent=83"
      _StyleDefs(105) =   "Splits(0).Columns(16).Style:id=274,.parent=71,.locked=-1"
      _StyleDefs(106) =   "Splits(0).Columns(16).HeadingStyle:id=271,.parent=72,.alignment=2"
      _StyleDefs(107) =   "Splits(0).Columns(16).FooterStyle:id=272,.parent=73"
      _StyleDefs(108) =   "Splits(0).Columns(16).EditorStyle:id=273,.parent=83"
      _StyleDefs(109) =   "Splits(0).Columns(17).Style:id=278,.parent=71,.locked=-1"
      _StyleDefs(110) =   "Splits(0).Columns(17).HeadingStyle:id=275,.parent=72,.alignment=2"
      _StyleDefs(111) =   "Splits(0).Columns(17).FooterStyle:id=276,.parent=73"
      _StyleDefs(112) =   "Splits(0).Columns(17).EditorStyle:id=277,.parent=83"
      _StyleDefs(113) =   "Splits(0).Columns(18).Style:id=282,.parent=71,.locked=-1"
      _StyleDefs(114) =   "Splits(0).Columns(18).HeadingStyle:id=279,.parent=72,.alignment=2"
      _StyleDefs(115) =   "Splits(0).Columns(18).FooterStyle:id=280,.parent=73"
      _StyleDefs(116) =   "Splits(0).Columns(18).EditorStyle:id=281,.parent=83"
      _StyleDefs(117) =   "Splits(0).Columns(19).Style:id=286,.parent=71,.locked=-1"
      _StyleDefs(118) =   "Splits(0).Columns(19).HeadingStyle:id=283,.parent=72,.alignment=2"
      _StyleDefs(119) =   "Splits(0).Columns(19).FooterStyle:id=284,.parent=73"
      _StyleDefs(120) =   "Splits(0).Columns(19).EditorStyle:id=285,.parent=83"
      _StyleDefs(121) =   "Splits(0).Columns(20).Style:id=290,.parent=71,.locked=-1"
      _StyleDefs(122) =   "Splits(0).Columns(20).HeadingStyle:id=287,.parent=72,.alignment=2"
      _StyleDefs(123) =   "Splits(0).Columns(20).FooterStyle:id=288,.parent=73"
      _StyleDefs(124) =   "Splits(0).Columns(20).EditorStyle:id=289,.parent=83"
      _StyleDefs(125) =   "Splits(0).Columns(21).Style:id=294,.parent=71,.locked=-1"
      _StyleDefs(126) =   "Splits(0).Columns(21).HeadingStyle:id=291,.parent=72,.alignment=2"
      _StyleDefs(127) =   "Splits(0).Columns(21).FooterStyle:id=292,.parent=73"
      _StyleDefs(128) =   "Splits(0).Columns(21).EditorStyle:id=293,.parent=83"
      _StyleDefs(129) =   "Splits(0).Columns(22).Style:id=298,.parent=71,.locked=-1"
      _StyleDefs(130) =   "Splits(0).Columns(22).HeadingStyle:id=295,.parent=72,.alignment=2"
      _StyleDefs(131) =   "Splits(0).Columns(22).FooterStyle:id=296,.parent=73"
      _StyleDefs(132) =   "Splits(0).Columns(22).EditorStyle:id=297,.parent=83"
      _StyleDefs(133) =   "Splits(0).Columns(23).Style:id=302,.parent=71,.locked=-1"
      _StyleDefs(134) =   "Splits(0).Columns(23).HeadingStyle:id=299,.parent=72,.alignment=2"
      _StyleDefs(135) =   "Splits(0).Columns(23).FooterStyle:id=300,.parent=73"
      _StyleDefs(136) =   "Splits(0).Columns(23).EditorStyle:id=301,.parent=83"
      _StyleDefs(137) =   "Splits(0).Columns(24).Style:id=306,.parent=71,.locked=-1"
      _StyleDefs(138) =   "Splits(0).Columns(24).HeadingStyle:id=303,.parent=72,.alignment=2"
      _StyleDefs(139) =   "Splits(0).Columns(24).FooterStyle:id=304,.parent=73"
      _StyleDefs(140) =   "Splits(0).Columns(24).EditorStyle:id=305,.parent=83"
      _StyleDefs(141) =   "Splits(0).Columns(25).Style:id=310,.parent=71,.locked=-1"
      _StyleDefs(142) =   "Splits(0).Columns(25).HeadingStyle:id=307,.parent=72,.alignment=2"
      _StyleDefs(143) =   "Splits(0).Columns(25).FooterStyle:id=308,.parent=73"
      _StyleDefs(144) =   "Splits(0).Columns(25).EditorStyle:id=309,.parent=83"
      _StyleDefs(145) =   "Splits(0).Columns(26).Style:id=314,.parent=71,.locked=-1"
      _StyleDefs(146) =   "Splits(0).Columns(26).HeadingStyle:id=311,.parent=72"
      _StyleDefs(147) =   "Splits(0).Columns(26).FooterStyle:id=312,.parent=73"
      _StyleDefs(148) =   "Splits(0).Columns(26).EditorStyle:id=313,.parent=83"
      _StyleDefs(149) =   "Splits(0).Columns(27).Style:id=318,.parent=71,.locked=-1"
      _StyleDefs(150) =   "Splits(0).Columns(27).HeadingStyle:id=315,.parent=72,.alignment=2"
      _StyleDefs(151) =   "Splits(0).Columns(27).FooterStyle:id=316,.parent=73"
      _StyleDefs(152) =   "Splits(0).Columns(27).EditorStyle:id=317,.parent=83"
      _StyleDefs(153) =   "Splits(0).Columns(28).Style:id=322,.parent=71,.locked=-1"
      _StyleDefs(154) =   "Splits(0).Columns(28).HeadingStyle:id=319,.parent=72,.alignment=2"
      _StyleDefs(155) =   "Splits(0).Columns(28).FooterStyle:id=320,.parent=73"
      _StyleDefs(156) =   "Splits(0).Columns(28).EditorStyle:id=321,.parent=83"
      _StyleDefs(157) =   "Splits(0).Columns(29).Style:id=326,.parent=71,.locked=-1"
      _StyleDefs(158) =   "Splits(0).Columns(29).HeadingStyle:id=323,.parent=72,.alignment=2"
      _StyleDefs(159) =   "Splits(0).Columns(29).FooterStyle:id=324,.parent=73"
      _StyleDefs(160) =   "Splits(0).Columns(29).EditorStyle:id=325,.parent=83"
      _StyleDefs(161) =   "Splits(0).Columns(30).Style:id=330,.parent=71,.locked=-1"
      _StyleDefs(162) =   "Splits(0).Columns(30).HeadingStyle:id=327,.parent=72,.alignment=2"
      _StyleDefs(163) =   "Splits(0).Columns(30).FooterStyle:id=328,.parent=73"
      _StyleDefs(164) =   "Splits(0).Columns(30).EditorStyle:id=329,.parent=83"
      _StyleDefs(165) =   "Splits(0).Columns(31).Style:id=334,.parent=71,.locked=-1"
      _StyleDefs(166) =   "Splits(0).Columns(31).HeadingStyle:id=331,.parent=72,.alignment=2"
      _StyleDefs(167) =   "Splits(0).Columns(31).FooterStyle:id=332,.parent=73"
      _StyleDefs(168) =   "Splits(0).Columns(31).EditorStyle:id=333,.parent=83"
      _StyleDefs(169) =   "Splits(0).Columns(32).Style:id=338,.parent=71,.locked=-1"
      _StyleDefs(170) =   "Splits(0).Columns(32).HeadingStyle:id=335,.parent=72,.alignment=2"
      _StyleDefs(171) =   "Splits(0).Columns(32).FooterStyle:id=336,.parent=73"
      _StyleDefs(172) =   "Splits(0).Columns(32).EditorStyle:id=337,.parent=83"
      _StyleDefs(173) =   "Splits(0).Columns(33).Style:id=342,.parent=71,.locked=-1"
      _StyleDefs(174) =   "Splits(0).Columns(33).HeadingStyle:id=339,.parent=72,.alignment=2"
      _StyleDefs(175) =   "Splits(0).Columns(33).FooterStyle:id=340,.parent=73"
      _StyleDefs(176) =   "Splits(0).Columns(33).EditorStyle:id=341,.parent=83"
      _StyleDefs(177) =   "Splits(0).Columns(34).Style:id=346,.parent=71,.locked=-1"
      _StyleDefs(178) =   "Splits(0).Columns(34).HeadingStyle:id=343,.parent=72,.alignment=2"
      _StyleDefs(179) =   "Splits(0).Columns(34).FooterStyle:id=344,.parent=73"
      _StyleDefs(180) =   "Splits(0).Columns(34).EditorStyle:id=345,.parent=83"
      _StyleDefs(181) =   "Splits(0).Columns(35).Style:id=350,.parent=71,.locked=-1"
      _StyleDefs(182) =   "Splits(0).Columns(35).HeadingStyle:id=347,.parent=72,.alignment=2"
      _StyleDefs(183) =   "Splits(0).Columns(35).FooterStyle:id=348,.parent=73"
      _StyleDefs(184) =   "Splits(0).Columns(35).EditorStyle:id=349,.parent=83"
      _StyleDefs(185) =   "Splits(0).Columns(36).Style:id=354,.parent=71,.locked=-1"
      _StyleDefs(186) =   "Splits(0).Columns(36).HeadingStyle:id=351,.parent=72,.alignment=2"
      _StyleDefs(187) =   "Splits(0).Columns(36).FooterStyle:id=352,.parent=73"
      _StyleDefs(188) =   "Splits(0).Columns(36).EditorStyle:id=353,.parent=83"
      _StyleDefs(189) =   "Splits(0).Columns(37).Style:id=358,.parent=71,.locked=-1"
      _StyleDefs(190) =   "Splits(0).Columns(37).HeadingStyle:id=355,.parent=72,.alignment=2"
      _StyleDefs(191) =   "Splits(0).Columns(37).FooterStyle:id=356,.parent=73"
      _StyleDefs(192) =   "Splits(0).Columns(37).EditorStyle:id=357,.parent=83"
      _StyleDefs(193) =   "Splits(0).Columns(38).Style:id=362,.parent=71,.locked=-1"
      _StyleDefs(194) =   "Splits(0).Columns(38).HeadingStyle:id=359,.parent=72,.alignment=2"
      _StyleDefs(195) =   "Splits(0).Columns(38).FooterStyle:id=360,.parent=73"
      _StyleDefs(196) =   "Splits(0).Columns(38).EditorStyle:id=361,.parent=83"
      _StyleDefs(197) =   "Splits(0).Columns(39).Style:id=366,.parent=71,.locked=-1"
      _StyleDefs(198) =   "Splits(0).Columns(39).HeadingStyle:id=363,.parent=72,.alignment=2"
      _StyleDefs(199) =   "Splits(0).Columns(39).FooterStyle:id=364,.parent=73"
      _StyleDefs(200) =   "Splits(0).Columns(39).EditorStyle:id=365,.parent=83"
      _StyleDefs(201) =   "Splits(0).Columns(40).Style:id=378,.parent=71"
      _StyleDefs(202) =   "Splits(0).Columns(40).HeadingStyle:id=375,.parent=72"
      _StyleDefs(203) =   "Splits(0).Columns(40).FooterStyle:id=376,.parent=73"
      _StyleDefs(204) =   "Splits(0).Columns(40).EditorStyle:id=377,.parent=83"
      _StyleDefs(205) =   "Splits(0).Columns(41).Style:id=370,.parent=71"
      _StyleDefs(206) =   "Splits(0).Columns(41).HeadingStyle:id=367,.parent=72"
      _StyleDefs(207) =   "Splits(0).Columns(41).FooterStyle:id=368,.parent=73"
      _StyleDefs(208) =   "Splits(0).Columns(41).EditorStyle:id=369,.parent=83"
      _StyleDefs(209) =   "Splits(1).Style:id=13,.parent=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(210) =   ":id=13,.strikethrough=0,.charset=0"
      _StyleDefs(211) =   ":id=13,.fontname=Verdana"
      _StyleDefs(212) =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(213) =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(214) =   "Splits(1).FooterStyle:id=15,.parent=3,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(215) =   ":id=15,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(216) =   ":id=15,.fontname=Verdana"
      _StyleDefs(217) =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(218) =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(219) =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(220) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(221) =   "Splits(1).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(222) =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(223) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(224) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(225) =   "Splits(1).Columns(0).Style:id=28,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(226) =   "Splits(1).Columns(0).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(227) =   "Splits(1).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(228) =   "Splits(1).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(229) =   "Splits(1).Columns(1).Style:id=32,.parent=13,.alignment=2,.locked=0"
      _StyleDefs(230) =   "Splits(1).Columns(1).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(231) =   "Splits(1).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(232) =   "Splits(1).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(233) =   "Splits(1).Columns(2).Style:id=186,.parent=13,.locked=-1"
      _StyleDefs(234) =   "Splits(1).Columns(2).HeadingStyle:id=183,.parent=14,.alignment=2"
      _StyleDefs(235) =   "Splits(1).Columns(2).FooterStyle:id=184,.parent=15"
      _StyleDefs(236) =   "Splits(1).Columns(2).EditorStyle:id=185,.parent=17"
      _StyleDefs(237) =   "Splits(1).Columns(3).Style:id=182,.parent=13,.locked=-1"
      _StyleDefs(238) =   "Splits(1).Columns(3).HeadingStyle:id=179,.parent=14,.alignment=2"
      _StyleDefs(239) =   "Splits(1).Columns(3).FooterStyle:id=180,.parent=15"
      _StyleDefs(240) =   "Splits(1).Columns(3).EditorStyle:id=181,.parent=17"
      _StyleDefs(241) =   "Splits(1).Columns(4).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(242) =   "Splits(1).Columns(4).HeadingStyle:id=55,.parent=14,.alignment=2"
      _StyleDefs(243) =   "Splits(1).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(244) =   "Splits(1).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(245) =   "Splits(1).Columns(5).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(246) =   "Splits(1).Columns(5).HeadingStyle:id=59,.parent=14,.alignment=2"
      _StyleDefs(247) =   "Splits(1).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(248) =   "Splits(1).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(249) =   "Splits(1).Columns(6).Style:id=66,.parent=13,.locked=-1"
      _StyleDefs(250) =   "Splits(1).Columns(6).HeadingStyle:id=63,.parent=14,.alignment=2"
      _StyleDefs(251) =   "Splits(1).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(252) =   "Splits(1).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(253) =   "Splits(1).Columns(7).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(254) =   "Splits(1).Columns(7).HeadingStyle:id=67,.parent=14,.alignment=2"
      _StyleDefs(255) =   "Splits(1).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(256) =   "Splits(1).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(257) =   "Splits(1).Columns(8).Style:id=202,.parent=13,.locked=-1"
      _StyleDefs(258) =   "Splits(1).Columns(8).HeadingStyle:id=199,.parent=14,.alignment=2"
      _StyleDefs(259) =   "Splits(1).Columns(8).FooterStyle:id=200,.parent=15"
      _StyleDefs(260) =   "Splits(1).Columns(8).EditorStyle:id=201,.parent=17"
      _StyleDefs(261) =   "Splits(1).Columns(9).Style:id=82,.parent=13,.locked=-1"
      _StyleDefs(262) =   "Splits(1).Columns(9).HeadingStyle:id=79,.parent=14,.alignment=2"
      _StyleDefs(263) =   "Splits(1).Columns(9).FooterStyle:id=80,.parent=15"
      _StyleDefs(264) =   "Splits(1).Columns(9).EditorStyle:id=81,.parent=17"
      _StyleDefs(265) =   "Splits(1).Columns(10).Style:id=178,.parent=13,.locked=-1"
      _StyleDefs(266) =   "Splits(1).Columns(10).HeadingStyle:id=175,.parent=14,.alignment=2"
      _StyleDefs(267) =   "Splits(1).Columns(10).FooterStyle:id=176,.parent=15"
      _StyleDefs(268) =   "Splits(1).Columns(10).EditorStyle:id=177,.parent=17"
      _StyleDefs(269) =   "Splits(1).Columns(11).Style:id=214,.parent=13,.locked=-1"
      _StyleDefs(270) =   "Splits(1).Columns(11).HeadingStyle:id=211,.parent=14,.alignment=2"
      _StyleDefs(271) =   "Splits(1).Columns(11).FooterStyle:id=212,.parent=15"
      _StyleDefs(272) =   "Splits(1).Columns(11).EditorStyle:id=213,.parent=17"
      _StyleDefs(273) =   "Splits(1).Columns(12).Style:id=90,.parent=13,.locked=-1"
      _StyleDefs(274) =   "Splits(1).Columns(12).HeadingStyle:id=87,.parent=14,.alignment=2"
      _StyleDefs(275) =   "Splits(1).Columns(12).FooterStyle:id=88,.parent=15"
      _StyleDefs(276) =   "Splits(1).Columns(12).EditorStyle:id=89,.parent=17"
      _StyleDefs(277) =   "Splits(1).Columns(13).Style:id=94,.parent=13,.locked=-1"
      _StyleDefs(278) =   "Splits(1).Columns(13).HeadingStyle:id=91,.parent=14,.alignment=2"
      _StyleDefs(279) =   "Splits(1).Columns(13).FooterStyle:id=92,.parent=15"
      _StyleDefs(280) =   "Splits(1).Columns(13).EditorStyle:id=93,.parent=17"
      _StyleDefs(281) =   "Splits(1).Columns(14).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(282) =   "Splits(1).Columns(14).HeadingStyle:id=75,.parent=14,.alignment=2"
      _StyleDefs(283) =   "Splits(1).Columns(14).FooterStyle:id=76,.parent=15"
      _StyleDefs(284) =   "Splits(1).Columns(14).EditorStyle:id=77,.parent=17"
      _StyleDefs(285) =   "Splits(1).Columns(15).Style:id=98,.parent=13,.locked=-1"
      _StyleDefs(286) =   "Splits(1).Columns(15).HeadingStyle:id=95,.parent=14,.alignment=2"
      _StyleDefs(287) =   "Splits(1).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(288) =   "Splits(1).Columns(15).EditorStyle:id=97,.parent=17"
      _StyleDefs(289) =   "Splits(1).Columns(16).Style:id=102,.parent=13,.locked=-1"
      _StyleDefs(290) =   "Splits(1).Columns(16).HeadingStyle:id=99,.parent=14,.alignment=2"
      _StyleDefs(291) =   "Splits(1).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(292) =   "Splits(1).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(293) =   "Splits(1).Columns(17).Style:id=106,.parent=13,.locked=-1"
      _StyleDefs(294) =   "Splits(1).Columns(17).HeadingStyle:id=103,.parent=14,.alignment=2"
      _StyleDefs(295) =   "Splits(1).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(296) =   "Splits(1).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(297) =   "Splits(1).Columns(18).Style:id=206,.parent=13,.locked=-1"
      _StyleDefs(298) =   "Splits(1).Columns(18).HeadingStyle:id=203,.parent=14,.alignment=2"
      _StyleDefs(299) =   "Splits(1).Columns(18).FooterStyle:id=204,.parent=15"
      _StyleDefs(300) =   "Splits(1).Columns(18).EditorStyle:id=205,.parent=17"
      _StyleDefs(301) =   "Splits(1).Columns(19).Style:id=110,.parent=13,.locked=-1"
      _StyleDefs(302) =   "Splits(1).Columns(19).HeadingStyle:id=107,.parent=14,.alignment=2"
      _StyleDefs(303) =   "Splits(1).Columns(19).FooterStyle:id=108,.parent=15"
      _StyleDefs(304) =   "Splits(1).Columns(19).EditorStyle:id=109,.parent=17"
      _StyleDefs(305) =   "Splits(1).Columns(20).Style:id=114,.parent=13,.locked=-1"
      _StyleDefs(306) =   "Splits(1).Columns(20).HeadingStyle:id=111,.parent=14,.alignment=2"
      _StyleDefs(307) =   "Splits(1).Columns(20).FooterStyle:id=112,.parent=15"
      _StyleDefs(308) =   "Splits(1).Columns(20).EditorStyle:id=113,.parent=17"
      _StyleDefs(309) =   "Splits(1).Columns(21).Style:id=210,.parent=13,.locked=-1"
      _StyleDefs(310) =   "Splits(1).Columns(21).HeadingStyle:id=207,.parent=14,.alignment=2"
      _StyleDefs(311) =   "Splits(1).Columns(21).FooterStyle:id=208,.parent=15"
      _StyleDefs(312) =   "Splits(1).Columns(21).EditorStyle:id=209,.parent=17"
      _StyleDefs(313) =   "Splits(1).Columns(22).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(314) =   "Splits(1).Columns(22).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(315) =   "Splits(1).Columns(22).FooterStyle:id=48,.parent=15"
      _StyleDefs(316) =   "Splits(1).Columns(22).EditorStyle:id=49,.parent=17"
      _StyleDefs(317) =   "Splits(1).Columns(23).Style:id=118,.parent=13,.locked=-1"
      _StyleDefs(318) =   "Splits(1).Columns(23).HeadingStyle:id=115,.parent=14,.alignment=2"
      _StyleDefs(319) =   "Splits(1).Columns(23).FooterStyle:id=116,.parent=15"
      _StyleDefs(320) =   "Splits(1).Columns(23).EditorStyle:id=117,.parent=17"
      _StyleDefs(321) =   "Splits(1).Columns(24).Style:id=122,.parent=13,.locked=-1"
      _StyleDefs(322) =   "Splits(1).Columns(24).HeadingStyle:id=119,.parent=14,.alignment=2"
      _StyleDefs(323) =   "Splits(1).Columns(24).FooterStyle:id=120,.parent=15"
      _StyleDefs(324) =   "Splits(1).Columns(24).EditorStyle:id=121,.parent=17"
      _StyleDefs(325) =   "Splits(1).Columns(25).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(326) =   "Splits(1).Columns(25).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(327) =   "Splits(1).Columns(25).FooterStyle:id=44,.parent=15"
      _StyleDefs(328) =   "Splits(1).Columns(25).EditorStyle:id=45,.parent=17"
      _StyleDefs(329) =   "Splits(1).Columns(26).Style:id=218,.parent=13,.locked=-1"
      _StyleDefs(330) =   "Splits(1).Columns(26).HeadingStyle:id=215,.parent=14"
      _StyleDefs(331) =   "Splits(1).Columns(26).FooterStyle:id=216,.parent=15"
      _StyleDefs(332) =   "Splits(1).Columns(26).EditorStyle:id=217,.parent=17"
      _StyleDefs(333) =   "Splits(1).Columns(27).Style:id=126,.parent=13,.locked=-1"
      _StyleDefs(334) =   "Splits(1).Columns(27).HeadingStyle:id=123,.parent=14,.alignment=2"
      _StyleDefs(335) =   "Splits(1).Columns(27).FooterStyle:id=124,.parent=15"
      _StyleDefs(336) =   "Splits(1).Columns(27).EditorStyle:id=125,.parent=17"
      _StyleDefs(337) =   "Splits(1).Columns(28).Style:id=130,.parent=13,.locked=-1"
      _StyleDefs(338) =   "Splits(1).Columns(28).HeadingStyle:id=127,.parent=14,.alignment=2"
      _StyleDefs(339) =   "Splits(1).Columns(28).FooterStyle:id=128,.parent=15"
      _StyleDefs(340) =   "Splits(1).Columns(28).EditorStyle:id=129,.parent=17"
      _StyleDefs(341) =   "Splits(1).Columns(29).Style:id=134,.parent=13,.locked=-1"
      _StyleDefs(342) =   "Splits(1).Columns(29).HeadingStyle:id=131,.parent=14,.alignment=2"
      _StyleDefs(343) =   "Splits(1).Columns(29).FooterStyle:id=132,.parent=15"
      _StyleDefs(344) =   "Splits(1).Columns(29).EditorStyle:id=133,.parent=17"
      _StyleDefs(345) =   "Splits(1).Columns(30).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(346) =   "Splits(1).Columns(30).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(347) =   "Splits(1).Columns(30).FooterStyle:id=52,.parent=15"
      _StyleDefs(348) =   "Splits(1).Columns(30).EditorStyle:id=53,.parent=17"
      _StyleDefs(349) =   "Splits(1).Columns(31).Style:id=142,.parent=13,.locked=-1"
      _StyleDefs(350) =   "Splits(1).Columns(31).HeadingStyle:id=139,.parent=14,.alignment=2"
      _StyleDefs(351) =   "Splits(1).Columns(31).FooterStyle:id=140,.parent=15"
      _StyleDefs(352) =   "Splits(1).Columns(31).EditorStyle:id=141,.parent=17"
      _StyleDefs(353) =   "Splits(1).Columns(32).Style:id=146,.parent=13,.locked=-1"
      _StyleDefs(354) =   "Splits(1).Columns(32).HeadingStyle:id=143,.parent=14,.alignment=2"
      _StyleDefs(355) =   "Splits(1).Columns(32).FooterStyle:id=144,.parent=15"
      _StyleDefs(356) =   "Splits(1).Columns(32).EditorStyle:id=145,.parent=17"
      _StyleDefs(357) =   "Splits(1).Columns(33).Style:id=150,.parent=13,.locked=-1"
      _StyleDefs(358) =   "Splits(1).Columns(33).HeadingStyle:id=147,.parent=14,.alignment=2"
      _StyleDefs(359) =   "Splits(1).Columns(33).FooterStyle:id=148,.parent=15"
      _StyleDefs(360) =   "Splits(1).Columns(33).EditorStyle:id=149,.parent=17"
      _StyleDefs(361) =   "Splits(1).Columns(34).Style:id=154,.parent=13,.locked=-1"
      _StyleDefs(362) =   "Splits(1).Columns(34).HeadingStyle:id=151,.parent=14,.alignment=2"
      _StyleDefs(363) =   "Splits(1).Columns(34).FooterStyle:id=152,.parent=15"
      _StyleDefs(364) =   "Splits(1).Columns(34).EditorStyle:id=153,.parent=17"
      _StyleDefs(365) =   "Splits(1).Columns(35).Style:id=158,.parent=13,.locked=-1"
      _StyleDefs(366) =   "Splits(1).Columns(35).HeadingStyle:id=155,.parent=14,.alignment=2"
      _StyleDefs(367) =   "Splits(1).Columns(35).FooterStyle:id=156,.parent=15"
      _StyleDefs(368) =   "Splits(1).Columns(35).EditorStyle:id=157,.parent=17"
      _StyleDefs(369) =   "Splits(1).Columns(36).Style:id=162,.parent=13,.locked=-1"
      _StyleDefs(370) =   "Splits(1).Columns(36).HeadingStyle:id=159,.parent=14,.alignment=2"
      _StyleDefs(371) =   "Splits(1).Columns(36).FooterStyle:id=160,.parent=15"
      _StyleDefs(372) =   "Splits(1).Columns(36).EditorStyle:id=161,.parent=17"
      _StyleDefs(373) =   "Splits(1).Columns(37).Style:id=166,.parent=13,.locked=-1"
      _StyleDefs(374) =   "Splits(1).Columns(37).HeadingStyle:id=163,.parent=14,.alignment=2"
      _StyleDefs(375) =   "Splits(1).Columns(37).FooterStyle:id=164,.parent=15"
      _StyleDefs(376) =   "Splits(1).Columns(37).EditorStyle:id=165,.parent=17"
      _StyleDefs(377) =   "Splits(1).Columns(38).Style:id=170,.parent=13,.locked=-1"
      _StyleDefs(378) =   "Splits(1).Columns(38).HeadingStyle:id=167,.parent=14,.alignment=2"
      _StyleDefs(379) =   "Splits(1).Columns(38).FooterStyle:id=168,.parent=15"
      _StyleDefs(380) =   "Splits(1).Columns(38).EditorStyle:id=169,.parent=17"
      _StyleDefs(381) =   "Splits(1).Columns(39).Style:id=174,.parent=13,.locked=-1"
      _StyleDefs(382) =   "Splits(1).Columns(39).HeadingStyle:id=171,.parent=14,.alignment=2"
      _StyleDefs(383) =   "Splits(1).Columns(39).FooterStyle:id=172,.parent=15"
      _StyleDefs(384) =   "Splits(1).Columns(39).EditorStyle:id=173,.parent=17"
      _StyleDefs(385) =   "Splits(1).Columns(40).Style:id=382,.parent=13"
      _StyleDefs(386) =   "Splits(1).Columns(40).HeadingStyle:id=379,.parent=14"
      _StyleDefs(387) =   "Splits(1).Columns(40).FooterStyle:id=380,.parent=15"
      _StyleDefs(388) =   "Splits(1).Columns(40).EditorStyle:id=381,.parent=17"
      _StyleDefs(389) =   "Splits(1).Columns(41).Style:id=374,.parent=13"
      _StyleDefs(390) =   "Splits(1).Columns(41).HeadingStyle:id=371,.parent=14"
      _StyleDefs(391) =   "Splits(1).Columns(41).FooterStyle:id=372,.parent=15"
      _StyleDefs(392) =   "Splits(1).Columns(41).EditorStyle:id=373,.parent=17"
      _StyleDefs(393) =   "Named:id=33:Normal"
      _StyleDefs(394) =   ":id=33,.parent=0"
      _StyleDefs(395) =   "Named:id=34:Heading"
      _StyleDefs(396) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(397) =   ":id=34,.wraptext=-1"
      _StyleDefs(398) =   "Named:id=35:Footing"
      _StyleDefs(399) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(400) =   "Named:id=36:Selected"
      _StyleDefs(401) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H0&"
      _StyleDefs(402) =   "Named:id=37:Caption"
      _StyleDefs(403) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(404) =   "Named:id=38:HighlightRow"
      _StyleDefs(405) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(406) =   "Named:id=39:EvenRow"
      _StyleDefs(407) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(408) =   "Named:id=40:OddRow"
      _StyleDefs(409) =   ":id=40,.parent=33"
      _StyleDefs(410) =   "Named:id=41:RecordSelector"
      _StyleDefs(411) =   ":id=41,.parent=34"
      _StyleDefs(412) =   "Named:id=42:FilterBar"
      _StyleDefs(413) =   ":id=42,.parent=33"
      _StyleDefs(414) =   "Named:id=0:"
      _StyleDefs(415) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(416) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(417) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(418) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(419) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(420) =   ":id=0,.fontname=MS Sans Serif"
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: SEARCH DATA CUSTOMER  ::."
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
      Left            =   4155
      TabIndex        =   14
      Top             =   75
      Width           =   2985
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   60
      X2              =   11200
      Y1              =   8970
      Y2              =   8970
   End
End
Attribute VB_Name = "frmSIDSearchCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim isTambah    As Boolean
Dim rsTemp      As New ADODB.Recordset
Dim rsHlp       As New ADODB.Recordset
Dim rsGrid      As New ADODB.Recordset
Dim rsCust      As New ADODB.Recordset
Dim rsTempCust  As New ADODB.Recordset

Dim strWhere    As String
Dim strWhere2   As String
Dim outPut()    As String
Dim strTemp     As String
Dim strTemp2    As String
Dim strTemp3    As String
Dim strTemp4    As String
Dim strTemp5    As String
'Dim SQLConBP    As String

Private sSQL As String
Private MySearch    As String
Private isPassing   As Boolean
Private IsAct       As String
Private MyRs        As ADODB.Recordset
Private rsPel       As ADODB.Recordset

Private cari        As String
Private cari2       As String
Private cari3       As String
Private cari4       As String
Private cari5       As String
Private carix       As String
Private cari2x      As String
Private cari3x      As String
Private cari4x      As String
Private cari5x      As String

Private isLoad As Boolean

Dim cols As TrueOleDBGrid70.Columns

Public Property Let SetSearch(value As String)
    MySearch = value
End Property

Public Property Get IsPass() As Boolean
    IsPass = isPassing
End Property

Public Property Get rs() As ADODB.Recordset
    Set rs = MyRs
End Property

Public Property Get Action() As String
    Action = IsAct
End Property

Private Sub btnClear_Click()
    Call clearForm
End Sub

Private Sub Form_Load()
    isLoad = False
    clearForm
    
End Sub

Private Sub clearForm()
    Dim i As Integer
    For i = 0 To 4
        text(i).text = ""
    Next
    
     text(9).text = ""
        text(8).text = ""
        text(7).text = ""
        text(6).text = ""
        tglLahir.value = "" '"__/__/____"
        text(5).text = ""
    
    btnTambah.Enabled = False
    ARProgressBar1.value = 0
    ARProgressBar1.Visible = False
    
    Call createRsGrid
End Sub

Private Sub btnCari_Click()
    Dim boleh1 As Boolean, boleh2 As Boolean
    boleh1 = True: boleh2 = True
    
    If NulSID(text(0)) = "" And NulSID(text(1)) = "" Then
        boleh1 = False
    End If
    If NulSID(text(2)) = "" And NulSID(text(3)) = "" And NulSID(text(4)) = "" And NulSID(tglLahir.value) = "" Then
        boleh2 = False
    End If
    If boleh1 = False And boleh2 = False Then
        MsgBox "Data filter kosong, tidak dapat proses data", vbInformation
        Exit Sub
    Else
        btnCari.Enabled = False
        Call PencarianData
        Grid.SetFocus
        btnCari.Enabled = True
    End If
End Sub

Private Sub PencarianData()
          Dim i As Integer
          Dim outPut() As String, param As String, paramX As String, paramBP As String, paramGR As String, paramSingle As String, paramSPK As String
          
1         cari = "": cari2 = "": cari3 = "": cari4 = "": cari5 = ""
2         carix = "": cari2x = "": cari3x = "": cari4x = "": cari5x = ""
3         param = "": paramX = "": paramBP = "": paramGR = ""
4         paramSingle = ""
5         paramSPK = ""
      '=============================================================================================
      '   cari    -   PENCARIAN DI UNIT CUSTOMER
      '=============================================================================================
6         If NulSID(text(1)) <> "" Then
7             cari = cari & " a.NoRangka Like '%" & NulSID(text(1)) & "%'"
8             carix = carix & " zNoRangka Like '%" & NulSID(text(1)) & "%'"
9         End If
10        If NulSID(text(0)) <> "" Then
11            If NulSID(cari) = "" Then
12                cari = cari & "a.NoPolisi Like '%" & NulSID(text(0)) & "%'"
13                carix = carix & "zNoPolisi Like '%" & NulSID(text(0)) & "%'"
14            Else
15                cari = cari & " OR a.NoPolisi Like '%" & NulSID(text(0)) & "%'"
16                carix = carix & " OR zNoPolisi Like '%" & NulSID(text(0)) & "%'"
17            End If
18        End If
19        If cari <> "" Then
20            cari = "(" & cari & ")"
21            carix = "(" & carix & ")"
22            paramSingle = paramSingle & " AND (NORANGKA LIKE '%" & NulSID(text(1)) & "%' OR NOPOLISI LIKE '%" & NulSID(text(0)) & "%')"
23        End If
          
      '=============================================================================================
      '   cari2   -   NO KTP / NPWP
      '=============================================================================================
24        If NulSID(text(2)) <> "" Then
25            cari2 = "((a.NoKTPSTNK Like '%" & NulSID(text(2)) & "%' OR a.NoNPWPSTNK Like '%" & NulSID(text(2)) & "%') OR (a.NoKTPSL Like '%" & NulSID(text(2)) & "%' OR a.NoNPWPSL Like '%" & NulSID(text(2)) & "%') OR (a.NoKTPSV Like '%" & NulSID(text(2)) & "%' OR a.NoNPWPSV Like '%" & NulSID(text(2)) & "%') OR (b.NoKTP Like '%" & NulSID(text(2)) & "%' OR b.NoNPWP Like '%" & NulSID(text(2)) & "%'))"
26            cari2x = "((zNoKTP Like '%" & NulSID(text(2)) & "%' OR zNoNPWP Like '%" & NulSID(text(2)) & "%'))"
27            paramSingle = paramSingle & " AND (NOKTP LIKE '%" & NulSID(text(2)) & "%' OR NONPWP LIKE '%" & NulSID(text(2)) & "%')"
28        End If
          
      '=============================================================================================
      '   cari3   -   NAMA CUSTOMER
      '=============================================================================================
29        If NulSID(text(3)) <> "" Then
      '        outPut = Split(NulSID(text(3)), " ")
              Dim txtcari As String
30            txtcari = Replace(NulSID(text(3)), " ", "%")
31            strTemp = "": strTemp2 = "": strTemp3 = "": strTemp4 = ""
      '        For i = 0 To UBound(outPut)
      '            If NulSID(outPut(i)) <> "" Then
      ''                If NulSID(strTemp) = "" Then
      ''                    cari3x = "(zNama Like '%" & NulSID(outPut(i)) & "%')"
      ''                    strTemp = "(a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp2 = "(a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp3 = "(a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp4 = "(b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(outPut(i)) & "%'"
      ''                Else
      ''                    cari3x = cari3x & " AND (" & "zNama Like '%" & NulSID(outPut(i)) & "%')"
      ''                    strTemp = strTemp & " AND (a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp2 = strTemp2 & " AND (a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp3 = strTemp3 & " AND (a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(outPut(i)) & "%'"
      ''                    strTemp4 = strTemp4 & " AND (b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(outPut(i)) & "%'"
      ''                End If
      '            End If
      '        Next
32            If NulSID(strTemp) = "" Then
33                cari3x = "(zNama Like '%" & NulSID(txtcari) & "%')"
34                strTemp = "(a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(txtcari) & "%'"
35                strTemp2 = "(a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(txtcari) & "%'"
36                strTemp3 = "(a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(txtcari) & "%'"
37                strTemp4 = "(b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(txtcari) & "%'"
38            Else
39                cari3x = cari3x & " AND (" & "zNama Like '%" & NulSID(txtcari) & "%')"
40                strTemp = strTemp & " AND (a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & NulSID(txtcari) & "%'"
41                strTemp2 = strTemp2 & " AND (a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & NulSID(txtcari) & "%'"
42                strTemp3 = strTemp3 & " AND (a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & NulSID(txtcari) & "%'"
43                strTemp4 = strTemp4 & " AND (b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & NulSID(txtcari) & "%'"
44            End If
      '        strTemp = strTemp & " AND (a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%" & nulsid(Text(3)) & "%'"
      '        strTemp2 = strTemp2 & " AND (a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%" & nulsid(Text(3)) & "%'"
      '        strTemp3 = strTemp3 & " AND (a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%" & nulsid(Text(3)) & "%'"
      '        strTemp4 = strTemp4 & " AND (b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%" & nulsid(Text(3)) & "%'"
              
45            cari3 = " ((" & strTemp & ") OR (" & strTemp2 & ") OR (" & strTemp3 & " ) OR (" & strTemp4 & " ))"
46            If cari3x <> "" Then cari3x = "(" & cari3x & ")"
47            paramSingle = paramSingle & " AND (NAMA1 + ' ' + NAMA2 + '' + NAMA3 LIKE '%" & NulSID(txtcari) & "%')"
48        End If
          
      '=============================================================================================
      '   cari4   -   TANGGAL LAHIR
      '=============================================================================================
49        If NulSID(tglLahir.value) <> "" Then
50            cari4 = "(a.tglLahirSTNK = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "' OR a.tglLahirSL = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "' OR a.tglLahirSV = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "' OR b.tglLahir = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "')"
51            cari4x = "(zTgLahir = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "')"
52            paramSingle = paramSingle & " AND (TGLLAHIR = '" & Format(NulSID(tglLahir.value), "MM/DD/YYYY") & "')"
53        End If
          
      '=============================================================================================
      '   cari5   -   KODE CUSTOMER
      '=============================================================================================
54        If NulSID(text(4)) <> "" Then
55            outPut = Split(NulSID(text(4)), " ")
56            strTemp = "": strTemp2 = "": strTemp3 = "": strTemp4 = "": strTemp5 = ""
57            For i = 0 To UBound(outPut)
58                If NulSID(outPut(i)) <> "" Then
59                    If NulSID(cari5) = "" Then
60                        cari5 = "(a.IDCustSales Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustService Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustAss Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSTNK Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSL Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSV Like '%" & NulSID(outPut(i)) & "%') OR (b.IDCust Like '%" & NulSID(outPut(i)) & "%')"
61                        cari5x = "(zKodeLg Like '%" & NulSID(outPut(i)) & "%')"
62                    Else
63                        cari5 = cari5 & " OR " & "(a.IDCustSales Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustService Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustAss Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSTNK Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSL Like '%" & NulSID(outPut(i)) & "%') OR (a.IDCustSV Like '%" & NulSID(outPut(i)) & "%') OR (b.IDCust Like '%" & NulSID(outPut(i)) & "%')"
64                        cari5x = cari5x & " OR " & "(zKodeLg Like '%" & NulSID(outPut(i)) & "%')"
65                    End If
66                End If
67            Next
68            cari5 = "(" & cari5 & ")"
69            cari5x = "(" & cari5x & ")"
70            paramSingle = paramSingle & " AND (IDCUST = '" & NulSID(text(4)) & "')"
71        End If
          
      '=============================================================================================
      '   SETTING POSISI PENCARIAN
      '=============================================================================================
72        If cari = "" Then
73            If cari2 = "" Then
74                If cari3 = "" Then
75                    If cari4 = "" Then
76                        If cari5 = "" Then
77                            Exit Sub
78                        Else
79                            param = cari5
80                            paramX = cari5x
81                        End If
82                    Else
83                        If cari5 = "" Then
84                            param = cari4
85                            paramX = cari4x
86                        Else
87                            param = "(" & cari4 & " AND " & cari5 & ")"
88                            paramX = "(" & cari4x & " AND " & cari5x & ")"
89                        End If
90                    End If
91                Else
92                    If cari4 = "" Then
93                        If cari5 = "" Then
94                            param = cari3
95                            paramX = cari3x
96                        Else
97                            param = "(" & cari3 & " AND " & cari5 & ")"
98                            paramX = "(" & cari3x & " AND " & cari5x & ")"
99                        End If
100                   Else
101                       If cari5 = "" Then
102                           param = "(" & cari3 & " AND " & cari4 & ")"
103                           paramX = "(" & cari3x & " AND " & cari4x & ")"
104                       Else
105                           param = "(" & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
106                           paramX = "(" & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
107                       End If
108                   End If
109               End If
110           Else
111               If cari3 = "" Then
112                   If cari4 = "" Then
113                       If cari5 = "" Then
114                           param = cari2
115                           paramX = cari2x
116                       Else
117                           param = "(" & cari2 & " AND " & cari5 & ")"
118                           paramX = "(" & cari2x & " AND " & cari5x & ")"
119                       End If
120                   Else
121                       If cari5 = "" Then
122                           param = "(" & cari2 & " AND " & cari4 & ")"
123                           paramX = "(" & cari2x & " AND " & cari4x & ")"
124                       Else
125                           param = "(" & cari2 & " AND " & cari4 & " AND " & cari5 & ")"
126                           paramX = "(" & cari2x & " AND " & cari4x & " AND " & cari5x & ")"
127                       End If
128                   End If
129               Else
130                   If cari4 = "" Then
131                       If cari5 = "" Then
132                           param = "(" & cari2 & " AND " & cari3 & ")"
133                           paramX = "(" & cari2x & " AND " & cari3x & ")"
134                       Else
135                           param = "(" & cari2 & " AND " & cari3 & " AND " & cari5 & ")"
136                           paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari5x & ")"
137                       End If
138                   Else
139                       If cari5 = "" Then
140                           param = "(" & cari2 & " AND " & cari3 & " AND " & cari4 & ")"
141                           paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari4x & ")"
142                       Else
143                           param = "(" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
144                           paramX = "(" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
145                       End If
146                   End If
147               End If
148           End If
149       Else
150           If cari2 = "" Then
151               If cari3 = "" Then
152                   If cari4 = "" Then
153                       If cari5 = "" Then
154                           param = cari
155                           paramX = carix
156                       Else
157                           param = "(" & cari & ") AND (" & cari5 & ")"
158                           paramX = "(" & carix & ") AND (" & cari5x & ")"
159                       End If
160                   Else
161                       If cari5 = "" Then
162                           param = "(" & cari & ") AND (" & cari4 & ")"
163                           paramX = "(" & carix & ") AND (" & cari4x & ")"
164                       Else
165                           param = "(" & cari & ") AND (" & cari4 & " AND " & cari5 & ")"
166                           paramX = "(" & carix & ") AND (" & cari4x & " AND " & cari5x & ")"
167                       End If
168                   End If
169               Else
170                   If cari4 = "" Then
171                       If cari5 = "" Then
172                           param = "(" & cari & ") AND (" & cari3 & ")"
173                           paramX = "(" & carix & ") AND (" & cari3x & ")"
174                       Else
175                           param = "(" & cari & ") AND (" & cari3 & " AND " & cari5 & ")"
176                           paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari5x & ")"
177                       End If
178                   Else
179                       If cari5 = "" Then
180                           param = "(" & cari & ") AND (" & cari3 & " AND " & cari4 & ")"
181                           paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari4x & ")"
182                       Else
183                           param = "(" & cari & ") AND (" & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
184                           paramX = "(" & carix & ") AND (" & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
185                       End If
186                   End If
187               End If
188           Else
189               If cari3 = "" Then
190                   If cari4 = "" Then
191                       If cari5 = "" Then
192                           param = "(" & cari & ") AND (" & cari2 & ")"
193                           paramX = "(" & carix & ") AND (" & cari2x & ")"
194                       Else
195                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari5 & ")"
196                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari5x & ")"
197                       End If
198                   Else
199                       If cari5 = "" Then
200                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari4 & ")"
201                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari4x & ")"
202                       Else
203                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari4 & " AND " & cari5 & ")"
204                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari4x & " AND " & cari5x & ")"
205                       End If
206                   End If
207               Else
208                   If cari4 = "" Then
209                       If cari5 = "" Then
210                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & ")"
211                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & ")"
212                       Else
213                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari5 & ")"
214                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari5x & ")"
215                       End If
216                   Else
217                       If cari5 = "" Then
218                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari4 & ")"
219                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari4x & ")"
220                       Else
221                           param = "(" & cari & ") AND (" & cari2 & " AND " & cari3 & " AND " & cari4 & " AND " & cari5 & ")"
222                           paramX = "(" & carix & ") AND (" & cari2x & " AND " & cari3x & " AND " & cari4x & " AND " & cari5x & ")"
223                       End If
224                   End If
225               End If
226           End If
227       End If
          
      '    sSQL = "Select * From ( Select a.NoRangka, a.IDCustSTNK as IDCust, a.NoKTPSTNK as NoKTP, a.NoNPWPSTNK as NoNPWP, a.TitlesSTNK as Title, a.NamaCust1STNK as Nama1, a.NamaCust2STNK as Nama2, a.NamaCust3STNK as Nama3, a.TglLahirSTNK as TglLahir, a.AgamaSTNK as Agama, " & _
      '           "a.PekerjaanSTNK as Pekerjaan, a.PropinsiSTNK as Propinsi, a.KotaSTNK as Kota, a.KecamatanSTNK as Kecamatan, a.KelurahanSTNK as Kelurahan, a.KodePosSTNK as KodePos, a.AlamatSTNK as Alamat, a.RTSTNK as RT, a.RWSTNK as RW, a.IdTelpSTNK as KodeTelp, " & _
      '           "a.TelpSTNK as Telp, a.HP1STNK as HP1, a.HP2STNK as HP2, a.EmailSTNK as Email, a.HobiSTNK as Hobi, a.OlahragaSTNK as Olahraga, a.SukuSTNK as Suku, a.PreferensiSTNK as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, " & _
      '           "a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a Left Join NewCustomer b On a.IDCustSTNK = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSL as IDCust, a.NoKTPSL as " & _
      '           "NoKTP, a.NoNPWPSL as NoNPWP, a.TitlesSL as Title, a.NamaCust1SL as Nama1, a.NamaCust2SL as Nama2, a.NamaCust3SL as Nama3, a.TglLahirSL as tglLahir, a.AgamaSL as Agama, a.PekerjaanSL as Pekerjaan, a.PropinsiSL as Propinsi, a.KotaSL as Kota, " & _
      '           "a.KecamatanSL as Kecamatan, a.KelurahanSL as Kelurahan, a.KodePosSL as KodePos, a.AlamatSL as Alamat, a.RTSL as RT, a.RWSL as RW, a.IdTelpSL as KodeTelp, a.TelpSL as Telp, a.HP1SL as HP1, a.HP2SL as HP2, a.EmailSL as Email, a.HobiSL as Hobi, " & _
      '           "a.OlahragaSL as Olahraga, a.SukuSL as Suku, a.PreferensiSL as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil " & _
      '           "From UnitCustomer a Left Join NewCustomer b On a.IDCustSL = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSV as IDCust, a.NoKTPSV as NoKTP, a.NoNPWPSV as NoNPWP, a.TitlesSV as Title, a.NamaCust1SV as Nama1, a.NamaCust2SV as Nama2, " & _
      '           "a.NamaCust3SV as Nama3, a.TglLahirSV as TglLahir, a.AgamaSV as Agama, a.PekerjaanSV as Pekerjaan, a.PropinsiSV as Propinsi, a.KotaSV as Kota, a.KecamatanSV as Kecamatan, a.KelurahanSV as Kelurahan, a.KodePosSV as KodePos, a.AlamatSV as Alamat, " & _
      '           "a.RTSV as RT, a.RWSV as RW, a.IdTelpSV as KodeTelp, a.TelpSV as Telp, a.HP1SV as HP1, a.HP2SV as HP2, a.EmailSV as Email, a.HobiSV as Hobi, a.OlahragaSV as Olahraga, a.SukuSV as Suku, a.PreferensiSV as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, " & _
      '           "a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a Left Join NewCustomer b On a.IDCustSV = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, " & _
      '           "b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, " & _
      '           "b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a INNER Join NewCustomer b On " & _
      '           "a.IDCustService = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, " & _
      '           "b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, " & _
      '           "a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a INNER Join NewCustomer b On a.IDCustSales = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, " & _
      '           "b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, " & _
      '           "a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil From UnitCustomer a INNER Join NewCustomer b On a.IDCustAss = b.IDCust Where " & param
          
      '    sSQL = "Select * From ( Select a.NoRangka, a.IDCustSTNK as IDCust, a.NoKTPSTNK as NoKTP, a.NoNPWPSTNK as NoNPWP, a.NamaCust1STNK as Nama1, a.NamaCust2STNK as Nama2, a.NamaCust3STNK as Nama3, a.TglLahirSTNK as TglLahir, " & _
      '           " a.AlamatSTNK as Alamat, " & _
      '           " a.HP1STNK as HP1, a.NoPolisi, " & _
      '           "  From NewCustomer b Left Join UnitCustomer a On a.IDCustSTNK = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSL as IDCust, a.NoKTPSL as " & _
      '           "NoKTP, a.NoNPWPSL as NoNPWP, a.NamaCust1SL as Nama1, a.NamaCust2SL as Nama2, a.NamaCust3SL as Nama3, a.TglLahirSL as tglLahir, " & _
      '           "a.KecamatanSL as Kecamatan, a.KelurahanSL as Kelurahan, a.KodePosSL as KodePos, a.AlamatSL as Alamat, a.RTSL as RT, a.RWSL as RW, a.IdTelpSL as KodeTelp, a.TelpSL as Telp, a.HP1SL as HP1, a.HP2SL as HP2, a.EmailSL as Email, a.HobiSL as Hobi, " & _
      '           "a.OlahragaSL as Olahraga, a.SukuSL as Suku, a.PreferensiSL as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn  " & _
      '           "From NewCustomer b Left Join UnitCustomer a On a.IDCustSL = b.IDCust Where " & param & " Union Select a.NoRangka, a.IDCustSV as IDCust, a.NoKTPSV as NoKTP, a.NoNPWPSV as NoNPWP, a.TitlesSV as Title, a.NamaCust1SV as Nama1, a.NamaCust2SV as Nama2, " & _
      '           "a.NamaCust3SV as Nama3, a.TglLahirSV as TglLahir, a.AgamaSV as Agama, a.PekerjaanSV as Pekerjaan, a.PropinsiSV as Propinsi, a.KotaSV as Kota, a.KecamatanSV as Kecamatan, a.KelurahanSV as Kelurahan, a.KodePosSV as KodePos, a.AlamatSV as Alamat, " & _
      '           "a.RTSV as RT, a.RWSV as RW, a.IdTelpSV as KodeTelp, a.TelpSV as Telp, a.HP1SV as HP1, a.HP2SV as HP2, a.EmailSV as Email, a.HobiSV as Hobi, a.OlahragaSV as Olahraga, a.SukuSV as Suku, a.PreferensiSV as Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, " & _
      '           "a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn  From NewCustomer b Left Join UnitCustomer a On a.IDCustService = b.IDCust AND b.IDCust <> a.idcustsv Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, " & _
      '           "b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, " & _
      '           "b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn From NewCustomer b left join UnitCustomer a On " & _
      '           "a.IDCustService = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, " & _
      '           "b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, " & _
      '           "a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn From NewCustomer b left Join UnitCustomer a On a.IDCustSales = b.IDCust Where " & param & " Union Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.Title, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, " & _
      '           "b.TglLahir, b.Agama, b.Pekerjaan, b.Propinsi, b.Kota, b.Kecamatan, b.Kelurahan, b.KodePos, b.Alamat, b.RT, b.RW, b.IdTelp as KodeTelp, b.Telp, b.HP1, b.HP2, b.Email, b.Hobi, b.Olahraga, b.Suku, b.Preferensi, a.NoPolis, a.TglAwalPolis, a.TglAkhirPolis, a.NoPolisi, " & _
      ''           "a.NoEngine, a.NoKey, a.Mdl, a.Class, a.Fmdl, a.Color, a.KdDlrPenjual, a.NmDlrPenjual, a.Ddate, a.Thmobil,'' AS KodeLgn From NewCustomer b left Join UnitCustomer a On a.IDCustAss = b.IDCust Where " & param
228               sSQL = " Select * From ("
229               sSQL = sSQL & " Select a.NoRangka, a.IDCustSTNK as IDCust, a.NoKTPSTNK as NoKTP, a.NoNPWPSTNK as NoNPWP, a.NamaCust1STNK as Nama1, a.NamaCust2STNK as Nama2, a.NamaCust3STNK as Nama3, a.TglLahirSTNK as TglLahir,  a.AlamatSTNK as Alamat,  a.HP1STNK as HP1, a.NoPolisi, a.KotaSTNK as Kota,'' AS KodeLgn"
230               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
231               sSQL = sSQL & " Left Join UnitCustomer a WITH (NOLOCK) On a.IDCustSTNK = b.IDCust"
232               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
233               sSQL = sSQL & " Union"
234               sSQL = sSQL & " Select a.NoRangka, a.IDCustSL as IDCust, a.NoKTPSL as NoKTP, a.NoNPWPSL as NoNPWP, a.NamaCust1SL as Nama1, a.NamaCust2SL as Nama2, a.NamaCust3SL as Nama3, a.TglLahirSL as tglLahir, a.AlamatSL as Alamat,a.HP1SL as HP1, a.NoPolisi, a.KotaSl as Kota,'' AS KodeLgn"
235               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
236               sSQL = sSQL & " Left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustSL = b.IDCust"
237               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
238               sSQL = sSQL & " Union"
239               sSQL = sSQL & " Select a.NoRangka, a.IDCustSV as IDCust, a.NoKTPSV as NoKTP, a.NoNPWPSV as NoNPWP, a.NamaCust1SV as Nama1, a.NamaCust2SV as Nama2, a.NamaCust3SV as Nama3, a.TglLahirSV as TglLahir,a.AlamatSV as Alamat,a.HP1SV as HP1,a.NoPolisi, a.KotaSv as Kota,'' AS KodeLgn"
240               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
241               sSQL = sSQL & " Left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustService = b.IDCust AND b.IDCust <> a.idcustsv"
242               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
243               sSQL = sSQL & " Union"
244               sSQL = sSQL & " Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir,b.Alamat,b.HP1, a.NoPolisi, b.Kota as Kota,'' AS KodeLgn"
245               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
246               sSQL = sSQL & " left join UnitCustomer a WITH (NOLOCK)  On a.IDCustService = b.IDCust"
247               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
248               sSQL = sSQL & " Union"
249               sSQL = sSQL & " Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP,b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir,b.Alamat,b.HP1,a.NoPolisi, b.kota as Kota,'' AS KodeLgn"
250               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
251               sSQL = sSQL & " left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustSales = b.IDCust"
252               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
253               sSQL = sSQL & " Union"
254               sSQL = sSQL & " Select a.Norangka, b.IDCust, b.NoKTP, b.NoNPWP, b.NamaCust1 as Nama1, b.NamaCust2 as Nama2, b.NamaCust3 as Nama3, b.TglLahir,b.Alamat, b.HP1,a.NoPolisi, b.kota as Kota,'' AS KodeLgn"
255               sSQL = sSQL & " From NewCustomer b WITH (NOLOCK) "
256               sSQL = sSQL & " left Join UnitCustomer a WITH (NOLOCK)  On a.IDCustAss = b.IDCust"
257               sSQL = sSQL & " Where " & param ' (((a.NamaCust1STNK + ' ' + a.NamaCust2STNK + ' ' + a.NamaCust3STNK) Like '%WASISTO%') OR ((a.NamaCust1SL + ' ' + a.NamaCust2SL + ' ' + a.NamaCust3SL) Like '%WASISTO%') OR ((a.NamaCust1SV + ' ' + a.NamaCust2SV + ' ' + a.NamaCust3SV) Like '%WASISTO%' ) OR ((b.NamaCust1 + ' ' + b.NamaCust2 + ' ' + b.NamaCust3) Like '%WASISTO%' ))"
258               sSQL = sSQL & " ) Z Where Nama1 + Nama2 + Nama3 <> ''" & paramSingle
          'sSQL = sSQL & " ) Z Where Nama1 + Nama2 + Nama3 <> '' " & paramSingle
259       Set rsTemp = Nothing
          
260       Set rsTemp = objSID.FetchDataSID(sSQL)
261       Call createRsGrid
          
262       GoSub InsertData
          
263       paramBP = paramX: paramGR = paramX: paramSPK = paramX
          '======================================================================================================================
          '   CARI DI CUSTDAT BP
          '======================================================================================================================
264       paramBP = Replace(paramBP, "zNoRangka", "a.NoRangka")
265       paramBP = Replace(paramBP, "zNoPolisi", "a.NoPolisi")
266       paramBP = Replace(paramBP, "zNoKTP", "a.NoKTP")
267       paramBP = Replace(paramBP, "zNoNPWP", "a.NoNPWP")
268       paramBP = Replace(paramBP, "zNama", "(a.Nama1 + a.nama2 + a.nama3)")
269       paramBP = Replace(paramBP, "zTgLahir", "a.TglLahir")
270       paramBP = Replace(paramBP, "zKodeLg", "a.IDCust")
          
271       Set rsTemp = New ADODB.Recordset
272       If cekKoneksiBP = True Then
273           Set rsTemp = objSID.FetchListAllPelanggan("Where " & paramBP)
274           GoSub InsertData
275       End If
          '======================================================================================================================
          '   CARI DI CUSTDAT GRP
          '======================================================================================================================
276       paramGR = Replace(paramGR, "zNoRangka", "c.norangka")
277       paramGR = Replace(paramGR, "zNoPolisi", "c.NoPolisi")
278       paramGR = Replace(paramGR, "zNoKTP", "c.NoKTP")
279       paramGR = Replace(paramGR, "zNoNPWP", "c.NoNPWP")
280       paramGR = Replace(paramGR, "zNama", "(c.nama1 + c.nama2 + c.nama3)")
281       paramGR = Replace(paramGR, "zTgLahir", "c.TglLahir")
282       paramGR = Replace(paramGR, "zKodeLg", "c.IDCust")
          
283       Set rsTemp = New ADODB.Recordset
284       Set rsTemp = objSID.FetchPelangganGRP("Where " & paramGR)
285       GoSub InsertData
          
          '======================================================================================================================
          '   CARI DI CUSTDAT SPK
          '======================================================================================================================
286       paramSPK = Replace(paramSPK, "zNoRangka", "c.norangka")
287       paramSPK = Replace(paramSPK, "zNoPolisi", "c.NoPolisi")
288       paramSPK = Replace(paramSPK, "zNoKTP", "c.NoKTP")
289       paramSPK = Replace(paramSPK, "zNoNPWP", "c.NoNPWP")
290       paramSPK = Replace(paramSPK, "zNama", "c.NamaFull")
291       paramSPK = Replace(paramSPK, "zTgLahir", "c.TglLahir")
292       paramSPK = Replace(paramSPK, "zKodeLg", "c.IDCust")
          
293       Set rsTemp = New ADODB.Recordset
294       Set rsTemp = objSID.fetchPelangganSPK("Where " & paramSPK)
295       GoSub InsertData
          
296       Set Grid.DataSource = rsGrid
297       btnTambah.Enabled = False
298       If rsGrid.EOF Then
299           btnTambah.Enabled = True
300       End If
301       Exit Sub
          
InsertData:
302       With rsGrid
303           If Not rsTemp.EOF Then 'If rsTemp.RecordCount <> 0 Then
304               rsTemp.MoveFirst
305               Do Until rsTemp.EOF
306                   .AddNew
307                   !Detail = "Ubah" 'False
308                   !Proses = "Proses" 'False
309                   !idcust = NulSID(rsTemp!idcust)
310                   !namaCust = NulSID(rsTemp!nama1) & " " & NulSID(rsTemp!nama2) & " " & NulSID(rsTemp!nama3)
311                   !tglLahir = Format((rsTemp!tglLahir), "DD/MM/YYYY")
312                   !NoPolisi = NulSID(rsTemp!NoPolisi)
313                   !Alamat = NulSID(rsTemp!Alamat)
314                   !Kota = NulSID(rsTemp!Kota)
315                   !HP = NulSID(rsTemp!HP1)
316                   !kodeLgn = NulSID(rsTemp!kodeLgn)
317                   If Len(NulSID(rsTemp!noKTP)) > 0 Then
318                       If Len(NulSID(rsTemp!noNPWP)) > 0 Then
319                           !NoID = NulSID(rsTemp!noKTP) & " - " & NulSID(rsTemp!noNPWP)
320                       Else
321                           !NoID = NulSID(rsTemp!noKTP)
322                       End If
323                   Else
324                       If Len(NulSID(rsTemp!noNPWP)) > 0 Then
325                           !NoID = NulSID(rsTemp!noKTP)
326                       Else
327                           !NoID = ""
328                       End If
329                   End If
330                   !noRangka = NulSID(rsTemp!noRangka)
331                   !noKTP = NulSID(rsTemp!noKTP)
332                   !noNPWP = NulSID(rsTemp!noNPWP)
                      
333                   .Update
334                   rsTemp.MoveNext
335               Loop
336           End If
337       End With
338       Return
End Sub

Private Sub pencarianDataLanjut()
    Dim bolehTambah As Boolean, kodeLgn As String, isWhere As String
    On Error GoTo pencarianDataLanjut_Error
    
    Call createCustomer
    Call createTempCustomer
    
'===============================================================================================
'   DATA BP
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(text(0)) <> "" Then strWhere = " c.NoPolisi LIKE '%" & NulSID(text(0)) & "%'"
    If NulSID(text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NoRangka LIKE '%" & NulSID(text(1)) & "%'"
        Else
            strWhere = strWhere & " c.NoRangka LIKE '%" & NulSID(text(1)) & "%'"
        End If
    End If
    
    If NulSID(text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.Nama LIKE '%" & NulSID(text(3)) & "%'"
        Else
            strWhere = strWhere & " c.Nama LIKE '%" & NulSID(text(3)) & "%'"
        End If
    End If
       
    Set rsPel = objSID.FetchListAllPelanggan("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!KodeLg) Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!KodeLg)
                    !KodeDealer = Right(gDb, 2)
                    !noRangka = NulSID(rsPel!noRangka)
                    !Tanggal = Format(NulSID(rsPel!tglentry), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
    
'===============================================================================================
'   DATA GRP
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(text(0)) <> "" Then strWhere = " c.NO_POLISI LIKE '%" & NulSID(text(0)) & "%'"
    If NulSID(text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NO_CHASSIS LIKE '%" & NulSID(text(1)) & "%'"
        Else
            strWhere = strWhere & " c.NO_CHASSIS LIKE '%" & NulSID(text(1)) & "%'"
        End If
    End If
    
    If NulSID(text(2)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NPWP LIKE '%" & NulSID(text(2)) & "%'"
        Else
            strWhere = strWhere & " c.NPWP LIKE '%" & NulSID(text(2)) & "%'"
        End If
    End If
    
    If NulSID(text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.CUSTOMER LIKE '%" & NulSID(text(3)) & "%'"
        Else
            strWhere = strWhere & " c.CUSTOMER LIKE '%" & NulSID(text(3)) & "%'"
        End If
    End If
    
    If NulSID(tglLahir.value) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.USETGLLAHIR = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        Else
            strWhere = strWhere & " AND c.USETGLLAHIR = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        End If
    End If
         
    Set rsPel = objSID.FetchPelangganGRP("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!kodeLgn) Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!kodeLgn)
                    !KodeDealer = Right(gDb, 2)
                    !noRangka = NulSID(rsPel!No_Chassis)
                    !Tanggal = Format(NulSID(isnow), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
    
'===============================================================================================
'   DATA SPK
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(text(0)) <> "" Then strWhere = " a.NoPol LIKE '%" & NulSID(text(0)) & "%'"
    If NulSID(text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NoRangka LIKE '%" & NulSID(text(1)) & "%'"
        Else
            strWhere = strWhere & " c.NoRangka LIKE '%" & NulSID(text(1)) & "%'"
        End If
    End If
    
    If NulSID(text(2)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NPWP LIKE '%" & NulSID(text(2)) & "%'"
        Else
            strWhere = strWhere & " c.NPWP LIKE '%" & NulSID(text(2)) & "%'"
        End If
    End If
    
    If NulSID(text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.Nama_Pemesan LIKE '%" & NulSID(text(3)) & "%'"
        Else
            strWhere = strWhere & " c.Nama_Pemesan LIKE '%" & NulSID(text(3)) & "%'"
        End If
    End If
    
    If NulSID(tglLahir.value) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        Else
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        End If
    End If
    
    Set rsPel = objSID.fetchPelangganSPK("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!kodeLgn) Then
                kodeLgn = NulSID(rsPel!kodeLgn)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!kodeLgn)
                    !KodeDealer = Right(gDb, 2)
                    !noRangka = NulSID(rsPel!noRangka)
                    !Tanggal = Format(NulSID(rsPel!tglentry), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
       
'===============================================================================================
'   DATA SPK HO
'===============================================================================================
    strWhere = "": isWhere = ""
    If NulSID(text(0)) <> "" Then strWhere = " a.NoPol LIKE '%" & NulSID(text(0)) & "%'"
    If NulSID(text(1)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND a.NoRangka LIKE '%" & NulSID(text(1)) & "%'"
        Else
            strWhere = strWhere & " a.NoRangka LIKE '%" & NulSID(text(1)) & "%'"
        End If
    End If
    
    If NulSID(text(2)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.NPWP LIKE '%" & NulSID(text(2)) & "%'"
        Else
            strWhere = strWhere & " c.NPWP LIKE '%" & NulSID(text(2)) & "%'"
        End If
    End If
    
    If NulSID(text(3)) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.Nama_Pemesan LIKE '%" & NulSID(text(3)) & "%'"
        Else
            strWhere = strWhere & " c.Nama_Pemesan LIKE '%" & NulSID(text(3)) & "%'"
        End If
    End If
    
    If NulSID(tglLahir.value) <> "" Then
        If NulSID(strWhere) <> "" Then
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        Else
            strWhere = strWhere & " AND c.TglLahir = '" & Format(tglLahir, "mm/dd/yyyy") & "'"
        End If
    End If
        
    Set rsPel = objSID.FetchPelangganSPK00("Where " & strWhere)
    If rsPel.RecordCount > 0 Then
        rsPel.MoveFirst
        Do Until rsPel.EOF
            bolehTambah = False
            If NulSID(kodeLgn) = "" Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            ElseIf NulSID(kodeLgn) <> NulSID(rsPel!KodeLg) Then
                kodeLgn = NulSID(rsPel!KodeLg)
                bolehTambah = True
            End If
            
            If bolehTambah Then
                With rsTempCust
                    .AddNew
                    !kodeLgn = NulSID(rsPel!KodeLg)
                    !KodeDealer = Right(rsPel!DBname, 2)
                    !noRangka = NulSID(rsPel!noRangka)
                    !Tanggal = Format(NulSID(isnow), "dd/mm/yyyy")
                    .Update
                End With
            End If
            rsPel.MoveNext
        Loop
    End If
          
'===============================================================================================
'   proses filter data dari kodelgn terbaru
'===============================================================================================
    If Not rsTempCust.EOF Then
        rsTempCust.MoveFirst
        Do While Not rsTempCust.EOF
            rsCust.filter = " KODELGN = '" & NulSID(rsTempCust!kodeLgn) & "'"
            If rsCust.EOF Then
                With rsCust
                    .AddNew
                    !kodeLgn = NulSID(rsTempCust!kodeLgn)
                    !KodeDealer = NulSID(rsTempCust!KodeDealer)
                    !noRangka = NulSID(rsTempCust!noRangka)
                    !Tanggal = Format(NulSID(rsTempCust!Tanggal), "dd/mm/yyyy")
                    .Update
                End With
                rsCust.MoveFirst
            Else
                rsCust.MoveFirst
                If DateDiff("d", Format(NulSID(rsCust!Tanggal), "dd/mm/yyyy"), Format(NulSID(rsTempCust!Tanggal), "dd/mm/yyyy")) > 0 Then
                    With rsCust
                        !kodeLgn = NulSID(rsTempCust!kodeLgn)
                        !KodeDealer = NulSID(rsTempCust!KodeDealer)
                        !noRangka = NulSID(rsTempCust!noRangka)
                        !Tanggal = Format(NulSID(rsTempCust!Tanggal), "dd/mm/yyyy")
                        .Update
                    End With
                End If
            End If
            rsCust.filter = ""
            rsTempCust.MoveNext
        Loop
    End If
       
    '===============================================================================================
    '   MENCOCOKAN DATA KODELGN DENGAN DATA BRIDGING
    '===============================================================================================
    If Not rsCust.EOF Then
        rsCust.MoveFirst
        Do While Not rsCust.EOF
            
            If rsTemp.State Then rsTemp.Close
            sSQL = "select DISTINCT b.[IDCustSales],b.[IDCustService]"
            sSQL = sSQL & " ,b.[NamaCust1STNK],b.[NamaCust2STNK],b.[NamaCust3STNK],b.[tglLahirSTNK],b.[AlamatSTNK],b.[kotaSTNK],b.[HP1STNK],b.[noKTPSTNK],b.[noNPWPSTNK],b.[NoRangka]"
            sSQL = sSQL & " ,b.[NamaCust1SL],b.[NamaCust2SL],b.[NamaCust3SL],b.[tglLahirSL],b.[AlamatSL],b.[kotaSL],b.[HP1SL],b.[noKTPSL],b.[noNPWPSL]"
            sSQL = sSQL & " ,b.[NamaCust1SV],b.[NamaCust2SV],b.[NamaCust3SV],b.[tglLahirSV],b.[AlamatSV],b.[kotaSV],b.[HP1SV],b.[noKTPSV],b.[noNPWPSV]"
            sSQL = sSQL & " ,b.[NOPOLISI]"
            sSQL = sSQL & " from Bridging a"
            sSQL = sSQL & " inner join UnitCustomer b on a.IDCust = b.IDCustSales OR a.IDCust = b.IDCustService"
            sSQL = sSQL & " where a.kodeLgn = '" & NulSID(rsCust!kodeLgn) & "' and a.kodeDealer = '" & Right(rsCust!KodeDealer, 2) & "'"
            rsTemp.Open sSQL, SQLConSID, adOpenStatic, adLockBatchOptimistic, adCmdText
            
            If Not rsTemp.EOF Then
                With rsGrid
                    .AddNew
                    !KodeCustomer = NulSID(rsTemp!idcustsales)
                    !NamaCustomer = NulSID(rsTemp!NamaCust1Stnk) & " " & NulSID(rsTemp!NamaCust2STNK) & " " & NulSID(rsTemp!NamaCust3STNK)
                    !tglLahir = TglNul(NulSID(rsTemp!tglLahirSTNK))
                    !noPol = NulSID(rsTemp!NoPolisi)
                    !Alamat = NulSID(rsTemp!AlamatSTNK)
                    !Kota = NulSID(rsTemp!Kotastnk)
                    !HP = NulSID(rsTemp!HP1Stnk)
                    !Identitas = NulSID(rsTemp!noKTPSTNK) & " / " & NulSID(rsTemp!noNPWPSTNK)
                    !noRangka = NulSID(rsTemp!noRangka)
                    !Action = "Detail"
                    !Action2 = "Proses"
                    .Update
                End With
            End If
            rsCust.MoveNext
        Loop
    End If
    On Error GoTo 0
    Exit Sub
pencarianDataLanjut_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure pencarianDataLanjut of Form frmSIDSearchCust"
End Sub

Private Sub Grid_ButtonClick(ByVal ColIndex As Integer)
    Call doClick(ColIndex)
End Sub

Private Sub doClick(Index As Integer)
          Dim kodeLgnLama As String
          Dim rsDataLama As ADODB.Recordset
          Dim Nama() As String
          Dim i As Integer
1         If Index = 0 Then
2             If NulSID(rsGrid.Fields("IdCust").value) = "" Then
3                 MsgBox "Data masih Id Customer kosong, tidak dapat melihat data", vbInformation
4                 Exit Sub
5             Else
6                 Set rsDataLama = objO.FetchDataACC("SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & NulSID(rsGrid.Fields("IdCust").value) & "' AND AFFILIASI = 1")
7                 If Not rsDataLama.EOF Then
8                   MsgBox "Tipe customer affiliasi tidak dapat diubah", vbInformation
9                   Exit Sub
10                End If
11                kodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
12                sSQL = " select KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry from Armst where KodeLgn = '" & NulSID(kodeLgnLama) & "'"
13                Set rsDataLama = objO.FetchDataACC(sSQL)
                  
                  
14                With frmSIDMasterPelanggan
15                    .setNoRangka = NulSID(rsGrid.Fields("norangka").value)
16                    .Txt(1) = NulSID(rsGrid.Fields("IdCust").value)
17                        .isEdit (True)
18                    .setKodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
19                    If Not rsDataLama.EOF Then
20                        Nama() = Split(NulSID(rsDataLama!Namalgn), " ")
21                        For i = 0 To UBound(Nama)
22                            If Nama(i) <> "" And i = 0 Then
23                                .Txt(4) = NulSID(Nama(i))
24                            ElseIf Nama(i) <> "" And i = 1 Then
25                                .Txt(5) = NulSID(Nama(i))
26                            Else
27                                .Txt(6) = .Txt(6) & NulSID(Nama(i))
28                            End If
29                        Next
30                        .Txt(3) = NulSID(rsDataLama!npwp)
                          
                          
                          ''pecah notelepon
31                        Nama() = Split(NulSID(rsDataLama!telepon), ",")
32                        For i = 0 To UBound(Nama)
33                            If Nama(i) <> "" And i = 0 Then
34                                .Txt(13) = NulSID(Nama(i))
35                            ElseIf Nama(i) <> "" And i = 1 Then
36                                .Txt(14) = NulSID(Nama(i))
37                            Else
38                                .Txt(15) = NulSID(Nama(i))
39                            End If
40                        Next
                          
41                        .Txt(8) = NulSID(rsDataLama!alamat1) & " " & NulSID(rsDataLama!alamat2) & " " & NulSID(rsDataLama!alamat3)
42                    End If
43                    .Show vbModal
44                End With
45            End If
              
46            Exit Sub
47        End If
48        If rsGrid.Fields("IdCust").value <> "" Then
              ''jika idcust belum ada
              ''cek ke data bridging apakah ada kodelgn lama
49            If gSearchOption = "frmSIDCustomerSurat" Then
50                frmSIDCustomer.text(66) = NulSID(rsGrid.Fields("IdCust").value)
51                gMainKeyCari = ""
52                gBantuan = ""
53                Unload Me
54                Exit Sub
55            End If
56            If NulSID(rsGrid.Fields("norangka").value) = "" Then
                  ''jika norangka kosong maka di arahkan ke penambahan norangka
57                If frmSIDCustomer.Visible = True Then
                      ''
58                    Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
                      '.isDariSearch = True
59                    With frmSIDCustomer
60                        .text(2) = NulSID(rsGrid.Fields("IdCust").value)
61                        .text(65) = NulSID(rsGrid.Fields("KodeLgn").value)
62                        For i = 0 To UBound(Nama)
63                            If Nama(i) <> "" And i = 0 Then
64                                .text(4) = NulSID(Nama(i))
65                            ElseIf Nama(i) <> "" And i = 1 Then
66                                .text(60) = NulSID(Nama(i))
67                            Else
68                                .text(61) = .Txt(6) & NulSID(Nama(i))
69                            End If
70                        Next
                          '.Text(0).SetFocus
71                    End With
72                    Unload Me
73                End If
74                gMainKeyCari = NulSID(rsGrid.Fields("IdCust").value)
75                gBantuan = NulSID(rsGrid.Fields("norangka").value)
76                Unload Me
77                Exit Sub
78            Else
              ''jika norangka sudah ada isinya dikembalikan ke form pemanggil
                  
79                If frmSIDCustomer.Visible = True Then
                      ''
80                    Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
                      
81                    With frmSIDCustomer
82                        .isDariSearch = True
                          '.text(0) = Right(nulsid(rsGrid.Fields("norangka").value), 8)
83                        If NulSID(Replace(Trim(.text(50)), ">", "") & Replace(Replace(.text(1), ">", ""), " ", "") & Replace(Replace(.text(0), ">", ""), " ", "")) <> NulSID(rsGrid.Fields("norangka").value) And NulSID(Replace(Trim(.text(50)), ">", "") & Replace(Replace(.text(1), ">", ""), " ", "") & Replace(Replace(.text(0), ">", ""), " ", "")) <> "" Then
84                            If MsgBox("No rangka yang dipilih berbeda dengan data customer." & vbCrLf & " Apakah ingin mengikuti perubahan?", vbYesNo) = vbYes Then
85                                .text(0) = Right(NulSID(rsGrid.Fields("norangka").value), 8)
86                                .text(1) = Mid(NulSID(rsGrid.Fields("norangka").value), 4, Len(NulSID(rsGrid.Fields("norangka").value)) - 11)
87                                .text(50) = Left(NulSID(rsGrid.Fields("norangka").value), 3)
88                            End If
89                        Else
90                            .text(0) = Right(NulSID(rsGrid.Fields("norangka").value), 8)
91                            .text(1) = Mid(NulSID(rsGrid.Fields("norangka").value), 4, Len(NulSID(rsGrid.Fields("norangka").value)) - 11)
92                            .text(50) = Left(NulSID(rsGrid.Fields("norangka").value), 3)
93                        End If
94                        .text(2) = NulSID(rsGrid.Fields("IdCust").value)
95                        .text(65) = NulSID(rsGrid.Fields("KodeLgn").value)
96                        For i = 0 To UBound(Nama)
97                            If Nama(i) <> "" And i = 0 Then
98                                .text(4) = NulSID(Nama(i))
99                            ElseIf Nama(i) <> "" And i = 1 Then
100                               .text(60) = NulSID(Nama(i))
101                           Else
102                               .text(61) = .Txt(6) & NulSID(Nama(i))
103                           End If
104                       Next
                          '.Text(0).SetFocus
105                   End With
106                   Unload Me
107               End If
108               gMainKeyCari = NulSID(rsGrid.Fields("IdCust").value)
109               gBantuan = NulSID(rsGrid.Fields("norangka").value)
110               Unload Me
111               Exit Sub
112           End If
              
113       Else
              ''jika belum ada idcust di arahkan ke tambah master
114           MsgBox "Data belum ada di Customer ID" & vbCrLf & "Silahkan untuk menambahkan data", vbInformation
              
115           kodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
116           sSQL = " select KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry from Armst where KodeLgn = '" & NulSID(kodeLgnLama) & "'"
117           Set rsDataLama = objO.FetchDataACC(sSQL)
              
118           sSQL = "select namacust1sv,namacust2sv,namacust3sv,tgllahirsv,agamasv,pekerjaansv,propinsisv,kotasv,kecamatansv,kelurahansv,kodepossv,alamatsv,rtsv,rwsv,hp1sv"
119           sSQL = sSQL & " From [SINGLEIDORI]..UnitCustomer"
120           sSQL = sSQL & " where NoRangka = '" & NulSID(rsGrid.Fields("norangka").value) & "'"
121           sSQL = sSQL & " and ltrim(rtrim(NamaCust1SV + ' ' + NamaCust2SV + ' ' + NamaCust3SV)) = '" & NulSID(rsGrid.Fields("namaCust").value) & "'"
122           sSQL = sSQL & " and HP1SV = '" & NulSID(rsGrid.Fields("hp").value) & "'"
123           sSQL = sSQL & " Union"
124           sSQL = sSQL & " select NamaCust1SL,namacust2sl,namacust3sl,tgllahirsl,agamasl,pekerjaansl,propinsisl,kotasl,kecamatansl,kelurahansl,kodepossl,alamatsl,rtsl,rwsl,hp1sl"
125           sSQL = sSQL & " From [SINGLEIDORI]..UnitCustomer"
126           sSQL = sSQL & " where NoRangka = '" & NulSID(rsGrid.Fields("norangka").value) & "'"
127           sSQL = sSQL & " and ltrim(rtrim(NamaCust1Sl + ' ' + NamaCust2Sl + ' ' + NamaCust3Sl)) = '" & NulSID(rsGrid.Fields("namaCust").value) & "'"
128           sSQL = sSQL & " and HP1Sl = '" & NulSID(rsGrid.Fields("hp").value) & "'"
129           sSQL = sSQL & " Union"
130           sSQL = sSQL & " select NamaCust1Stnk,namacust2stnk,namacust3stnk,tgllahirstnk,agamastnk,pekerjaanstnk,propinsistnk,kotastnk,kecamatanstnk,kelurahanstnk,kodeposstnk,alamatstnk,rtstnk,rwstnk,hp1stnk"
131           sSQL = sSQL & " From [SINGLEIDORI]..UnitCustomer"
132           sSQL = sSQL & " where NoRangka = '" & NulSID(rsGrid.Fields("norangka").value) & "'"
133           sSQL = sSQL & " and ltrim(rtrim(NamaCust1Stnk + ' ' + NamaCust2Stnk + ' ' + NamaCust3Stnk)) = '" & NulSID(rsGrid.Fields("namaCust").value) & "'"
134           sSQL = sSQL & " and HP1Stnk = '" & NulSID(rsGrid.Fields("hp").value) & "'"
135           Set rsDataLama = objO.FetchData(sSQL)
136           With frmSIDMasterPelanggan
137               .setKodeLgnLama = NulSID(rsGrid.Fields("KodeLgn").value)
138                 .isEdit (False)
139               If Not rsDataLama.EOF Then
140                   rsDataLama.MoveFirst
      '                Nama() = Split(nulsid(rsDataLama!Namalgn), " ")
      '                For i = 0 To UBound(Nama)
      '                    If Nama(i) <> "" And i = 0 Then
      '                        .Txt(4) = nulsid(Nama(i))
      '                    ElseIf Nama(i) <> "" And i = 1 Then
      '                        .Txt(5) = nulsid(Nama(i))
      '                    Else
      '                        .Txt(6) = .Txt(6) & nulsid(Nama(i))
      '                    End If
      '                Next
      '                .Txt(3) = nulsid(rsDataLama!npwp)
                      
141                   .Txt(4) = NulSID(rsDataLama!NamaCust1SV)
142                   .Txt(5) = NulSID(rsDataLama!NamaCust2SV)
143                   .Txt(6) = NulSID(rsDataLama!NamaCust3SV)
                      
144                   .Txt(14) = NulSID(rsDataLama!HP1SV)
                      
145                   .Txt(11) = NulSID(rsDataLama!HP1SV)
                      
146                   .Txt(8) = NulSID(rsDataLama!AlamatSV)
147                   .DtLahir.value = Format(NulSID(rsDataLama!tglLahirsv), "dd/mm/yyyy")
148                   .Txt(11) = NulSID(rsDataLama!KodePosSV)
149                   .Txt(63) = NulSID(rsDataLama!PropinsiSV)
150                   .Txt(64) = NulSID(rsDataLama!KotaSV)
151                   .Txt(65) = NulSID(rsDataLama!KecamatanSV)
152                   .Txt(66) = NulSID(rsDataLama!KelurahanSV)
                      
153                   .Txt(9) = NulSID(rsDataLama!RTSV)
154                   .Txt(10) = NulSID(rsDataLama!RWSV)
                      
155                   .Txt(20) = NulSID(rsDataLama!Agamasv)
                      
                      ''pecah notelepon
      '                Nama() = Split(nulsid(rsDataLama!telepon), ",")
      '                For i = 0 To UBound(Nama)
      '                    If Nama(i) <> "" And i = 0 Then
      '                        .Txt(13) = nulsid(Nama(i))
      '                    ElseIf Nama(i) <> "" And i = 1 Then
      '                        .Txt(14) = nulsid(Nama(i))
      '                    Else
      '                        .Txt(15) = nulsid(Nama(i))
      '                    End If
      '                Next
      '
      '                .Txt(8) = nulsid(rsDataLama!alamat1) & " " & nulsid(rsDataLama!alamat2) & " " & nulsid(rsDataLama!alamat3)
156               Else
157                   Nama() = Split(NulSID(rsGrid.Fields("namaCust").value), " ")
158                   For i = 0 To UBound(Nama)
159                       If Nama(i) <> "" And i = 0 Then
160                           .Txt(4) = NulSID(Nama(i))
161                       ElseIf Nama(i) <> "" And i = 1 Then
162                           .Txt(5) = NulSID(Nama(i))
163                       Else
164                           .Txt(6) = .Txt(6) & NulSID(Nama(i))
165                       End If
166                   Next
                      '.txt(3) = nulsid(rsGrid.Fields("namaCust").value)
167                   .Txt(8) = NulSID(rsGrid.Fields("Alamat").value)
168                   .Txt(64) = NulSID(rsGrid.Fields("kota").value)
169               End If
170               .Show vbModal
171           End With
172           Call PencarianData
173           Grid.SetFocus
174       End If
          'gMainKeyCari = ""
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Private Sub grid_FilterChange()
        On Error GoTo errHandler
    Set cols = Grid.Columns
    Dim C As Integer
    
    C = Grid.col
    Grid.HoldFields
    rsGrid.filter = getFilter()
    If rsGrid.EOF Then
        For Each col In Grid.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
    End If
    Grid.col = C
    Grid.EditActive = True
    
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    End If
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If rsGrid.RecordCount > 0 Then
        text(9).text = NulSID(rsGrid!noKTP) 'Mid(NulSID(rsGrid("NoID")), 1, InStr(NulSID(rsGrid("NoID")), " - "))
        text(8).text = NulSID(rsGrid!noNPWP) ' Mid(NulSID(rsGrid("NoID")), InStr(NulSID(rsGrid("NoID")), " - ") + 3)
        text(7).text = NulSID(rsGrid("IDCust"))
        text(6).text = NulSID(rsGrid("NamaCust"))
        tglLahirx.value = rsGrid("TglLahir")
        text(5).text = NulSID(rsGrid("Alamat"))
    Else
        text(9).text = ""
        text(8).text = ""
        text(7).text = ""
        text(6).text = ""
        tglLahirx.value = "__/__/____"
        text(5).text = ""
    End If
End Sub
 

Private Function cekKoneksiBP() As Boolean
    Dim hasil       As Boolean
    Dim ip          As String
    Dim database    As String
   On Error GoTo cekKoneksiBP_Error

    hasil = False
    ip = ""
    database = ""
    If rsTemp.State Then rsTemp.Close
    sSQL = "SELECT SiDbBP, isnull(BPServerIP,'') as IP,SiBklDual FROM profilacc..profil"
    rsTemp.Open sSQL, SQLConACC, adOpenStatic, adLockBatchOptimistic, adCmdText
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        If NulSID(rsTemp!SiBklDual) = 1 Then
            ip = NulSID(rsTemp!ip)
        Else
            ip = NulSID(Log.LogServ)
        End If
        
        database = NulSID(rsTemp!SiDbBP)

        If UCase(database) = "BP" Or NulSID(database) = "" Then
            hasil = False
        Else
            SQLConBP = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & Log.LogUser & ";password=" & Log.LogPass & ";Initial Catalog=" & database & " ;server=" & ip
            If rsTemp.State Then rsTemp.Close
            sSQL = "SELECT TOP 1 * FROM custdat"
            rsTemp.Open sSQL, SQLConBP, adOpenStatic, adLockBatchOptimistic, adCmdText
            hasil = True
            Log.logServBP = ip
            Log.logDbBP = database
            
        End If
    End If
    cekKoneksiBP = hasil

   On Error GoTo 0
   Exit Function

cekKoneksiBP_Error:
    cekKoneksiBP = False
End Function

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If Index = 3 Then
'            tglLahir.SetFocus
'        ElseIf Index = 4 Then
'            Call btnCari_Click
'        Else
'            If Text(Index).BackColor = fColor Then
'                Text(Index).Locked = True
'            Else
'                Text(Index).Locked = False
'                SendKeys ("{TAB}")
'            End If
'        End If
        Call btnCari_Click
    End If
    KeyAscii = UpCase(KeyAscii)
    
End Sub
 

Private Sub createTempCustomer()
1       On Error GoTo createTempCustomer_Error
2       If rsTempCust.State Then rsTempCust.Close
3       With rsTempCust
4           .Fields.Append "KodeLgn", adVarChar, 13
5           .Fields.Append "KodeDealer", adVarChar, 13
6           .Fields.Append "NoRangka", adVarChar, 20
7           .Fields.Append "tanggal", adDate
8           .Open
9       End With
10      On Error GoTo 0
11      Exit Sub
createTempCustomer_Error:
12      MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createTempCustomer of Form frmSIDSearchCust"
End Sub

Private Sub createCustomer()
1       On Error GoTo createCustomer_Error
2       If rsCust.State Then rsCust.Close
3       Set rsCust = Nothing
4       With rsCust
5           .Fields.Append "KodeLgn", adVarChar, 13
6           .Fields.Append "KodeDealer", adVarChar, 13
7           .Fields.Append "NoRangka", adVarChar, 20
8           .Fields.Append "tanggal", adDate
9           .Open
10      End With
11      On Error GoTo 0
12      Exit Sub
createCustomer_Error:
13    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createCustomer of Form frmSIDSearchCust"
End Sub

Private Sub createRsGrid()
  On Error GoTo createRsGrid_Error
  If rsGrid.State Then rsGrid.Close
  Set rsGrid = Nothing
  With rsGrid
      .Fields.Append "Detail", adVarChar, 10 'adBoolean
      .Fields.Append "Proses", adVarChar, 10
      .Fields.Append "IDCust", adVarChar, 13
      .Fields.Append "NamaCust", adVarChar, 210
      .Fields.Append "tglLahir", adVarChar, 10
      .Fields.Append "NoPolisi", adVarChar, 12
      .Fields.Append "Alamat", adVarChar, 150
      .Fields.Append "Kota", adVarChar, 50
      .Fields.Append "HP", adVarChar, 20
      .Fields.Append "NoID", adVarChar, 75
      .Fields.Append "NoRangka", adVarChar, 20
      .Fields.Append "kodeLgn", adVarChar, 40
      .Fields.Append "noNPWP", adVarChar, 75
      .Fields.Append "NoKTP", adVarChar, 75
'      .Fields.Append "Alamat", adVarChar, 150
'      .Fields.Append "Alamat", adVarChar, 150
'      .Fields.Append "Alamat", adVarChar, 150
'      .Fields.Append "Alamat", adVarChar, 150
      .Open
  End With
  'Set Grid.DataSource = rsGrid
  On Error GoTo 0
  Exit Sub
createRsGrid_Error:
  MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRsGrid of Form frmSIDSearchCust")
End Sub

Private Sub btnKeluar_Click()
    IsAct = "CLOSE"
    isPassing = False
    gMainKeyCari = ""
    gBantuan = ""
    Unload Me
End Sub

Private Sub btnTambah_Click()
    With frmSIDMasterPelanggan
        .setKodeLgnLama = ""
        .isEdit False
        .Show vbModal
    End With
    Call PencarianData
    Grid.SetFocus
End Sub

Private Sub tglLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        text(4).SetFocus
    End If
End Sub



'=====================
''20171027 - Param
'=====================
''    param = cari                                'NOPOL & NOKA
''
''    If cari2 <> "" Then                         'KTP
''        If cari = "" Then
''            param = "(" & cari2
''        Else
''            param = param & " AND (" & cari2
''        End If
''    End If
''
''    If cari3 <> "" Then                         'NAMA
''        If cari2 = "" Then
''            If cari = "" Then
''                param = "(" & cari3
''            Else
''                param = param & " AND (" & cari3
''            End If
''        Else
''            If cari = "" Then
''                param = param & " AND " & cari3
''            Else
''                param = param & " AND " & cari3
''            End If
''        End If
''    End If
''
''    If cari4 <> "" Then
''        If cari = "" Then
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    param = "(" & cari4
''                Else
''                    param = param & " AND " & cari4
''                End If
''            Else
''                param = param & " AND " & cari4
''            End If
''        Else
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    param = param & " AND (" & cari4
''                Else
''                    param = param & " AND " & cari4
''                End If
''            Else
''                If cari3 = "" Then
''                    param = param & " AND " & cari4
''                Else
''                    param = param & " AND " & cari4
''                End If
''            End If
''        End If
''    End If
''
''    If cari5 <> "" Then
''        If cari = "" Then
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = "(" & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            Else
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            End If
''        Else
''            If cari2 = "" Then
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = param & " AND (" & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            Else
''                If cari3 = "" Then
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                Else
''                    If cari4 = "" Then
''                        param = param & " AND " & cari5
''                    Else
''                        param = param & " AND " & cari5
''                    End If
''                End If
''            End If
''        End If
''    End If
