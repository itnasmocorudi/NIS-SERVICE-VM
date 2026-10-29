VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOPerbaikanNamaMekanik 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8280
      TabIndex        =   37
      Top             =   4920
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdGanti 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6000
      TabIndex        =   40
      Top             =   4920
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Ganti Nama Mekanik"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   5205
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   13455
      Begin MSAdodcLib.Adodc AdcKode 
         Height          =   330
         Left            =   8640
         Top             =   1080
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCKodeGroup 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   720
         Width           =   3195
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
         _ExtentX        =   5644
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Nama Dealer -"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   0
         TabIndex        =   4
         Top             =   1680
         Width           =   13320
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   4
            Left            =   930
            TabIndex        =   5
            Top             =   510
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   5
            Left            =   930
            TabIndex        =   6
            Top             =   800
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   6
            Left            =   930
            TabIndex        =   7
            Top             =   1100
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   7
            Left            =   930
            TabIndex        =   8
            Top             =   1400
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   10
            Left            =   6390
            TabIndex        =   9
            Top             =   450
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Mask            =   "9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   12
            Left            =   6390
            TabIndex        =   10
            Top             =   1110
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   15
            Left            =   6390
            TabIndex        =   11
            Top             =   1710
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   14
            Left            =   6390
            TabIndex        =   12
            Top             =   1410
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   13
            Left            =   8730
            TabIndex        =   13
            Top             =   1110
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   16
            Left            =   930
            TabIndex        =   14
            Top             =   1980
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   17
            Left            =   930
            TabIndex        =   15
            Top             =   2280
            Width           =   3210
            _ExtentX        =   5662
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   11
            Left            =   6390
            TabIndex        =   16
            Top             =   795
            Width           =   1215
            _ExtentX        =   2143
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
            Mask            =   "99-99-9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   8
            Left            =   930
            TabIndex        =   17
            Top             =   1680
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   9
            Left            =   6390
            TabIndex        =   18
            Top             =   2010
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   18
            Left            =   6390
            TabIndex        =   19
            Top             =   2310
            Width           =   3210
            _ExtentX        =   5662
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   2
            Left            =   10800
            TabIndex        =   41
            Top             =   1035
            Width           =   1245
            _ExtentX        =   2196
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
            Mask            =   ">99-99-9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   20
            Left            =   10800
            TabIndex        =   42
            Top             =   1440
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   1
            Left            =   10800
            TabIndex        =   45
            Top             =   1875
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mekanik"
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
            Left            =   9765
            TabIndex        =   46
            Top             =   1920
            Width           =   570
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KM Akhir"
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
            Left            =   9720
            TabIndex        =   44
            Top             =   1365
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Service"
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
            Left            =   9570
            TabIndex        =   43
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   3840
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   3840
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Customer && Kendaraan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4200
            TabIndex        =   34
            Top             =   0
            Width           =   3675
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   13320
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   13320
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   13335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GSM"
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
            TabIndex        =   33
            Top             =   1725
            Width           =   315
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            Left            =   375
            TabIndex        =   32
            Top             =   555
            Width           =   405
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            TabIndex        =   31
            Top             =   855
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
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
            TabIndex        =   30
            Top             =   1155
            Width           =   330
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PSTN"
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
            TabIndex        =   29
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Th. Mobil"
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
            Left            =   5595
            TabIndex        =   28
            Top             =   495
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
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
            Left            =   5820
            TabIndex        =   27
            Top             =   1125
            Width           =   420
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Warna"
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
            Left            =   5760
            TabIndex        =   26
            Top             =   1725
            Width           =   480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Polisi"
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
            Left            =   240
            TabIndex        =   25
            Top             =   2040
            Width           =   645
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frame."
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
            Left            =   330
            TabIndex        =   24
            Top             =   2280
            Width           =   510
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO INVOICE"
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
            Left            =   5445
            TabIndex        =   23
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CDMA"
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
            TabIndex        =   22
            Top             =   2085
            Width           =   435
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Engine."
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
            Left            =   5820
            TabIndex        =   21
            Top             =   2400
            Width           =   540
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   3960
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   3960
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Customer && Kendaraan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4320
            TabIndex        =   20
            Top             =   45
            Width           =   3675
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   8640
            X2              =   12720
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   8640
            X2              =   12720
            Y1              =   240
            Y2              =   240
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCNamaMekanik 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   4560
         Visible         =   0   'False
         Width           =   3195
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
         _ExtentX        =   5644
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Nama Mekanik -"
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
      Begin MSAdodcLib.Adodc AdcMek 
         Height          =   330
         Left            =   11160
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
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERBAIKAN NAMA MEKANIK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5310
         TabIndex        =   47
         Top             =   120
         Width           =   3990
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BorderColor     =   &H80000002&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   10
         Width           =   13455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Mekanik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   39
         Top             =   4560
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NO WO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1575
         TabIndex        =   36
         Top             =   1250
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dealer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   35
         Top             =   800
         Width           =   1215
      End
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHOPerbaikanNamaMekanik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dealer As String
Private dBase As String
Private koneksi As String
Private usernameDealer As String
Private passwordDealer As String
Private usernameHO As String
Private passwordHO As String
Private usernameGlobal As String
Private passwordGlobal As String
Private isDoubleCustDat As Boolean
Private isDoubleCustDatHO As Boolean
Private isGR As Boolean
Private FS_KD_DEALER As String

Private isDealerNewEnkripsi As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGanti_Click()
    If Len(Trim(text(0))) = 0 Then
        MsgBox "No WO Harus diisi", vbCritical, "Informasi"
        Exit Sub
    ElseIf Len(Trim(text(1))) = 0 Then
        MsgBox "KM Service yang benar harus diisi!", vbCritical, "Informasi"
        Exit Sub
    End If
    
    If MsgBox("Apakah yakin mengganti Mekanik " & Trim(text(1)) & " dengan " & Trim(DCNamaMekanik.value) & " ?", vbYesNo, "Konfirmasi Penggantian Nama Mekanik") = vbYes Then
        'lakukan pergantian data
        On Error GoTo vbError
        koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer
        erTranAll = False
        'perbaikan kecabang
        Call objHT.OpenTransactionPerbaikanDataHO(koneksi)
        
        'ganti no rangka
        'UPDATE CABANG
        Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDTA SET MECH = '" & DCNamaMekanik.value & "' WHERE NO_WO = '" & Trim(text(0)) & "'")
        Call objHT.UpdatePerbaikanDataHO("UPDATE T_INVOICE SET MECH = '" & DCNamaMekanik.value & "' WHERE NO_WO = '" & Trim(text(0)) & "'")
        Call objHT.UpdatePerbaikanDataHO("UPDATE HISTORY SET MECH = '" & DCNamaMekanik.value & "' WHERE NO_WO = '" & Trim(text(0)) & "'")
        
        Call objHT.UpdatePerbaikanDataHO("UPDATE TEMPPROGRESS SET MECH = '" & DCNamaMekanik.value & "' WHERE NOORDER = '" & Trim(text(0)) & "' AND GRP <> ''")
        Call objHT.UpdatePerbaikanDataHO("UPDATE STALL SET MECH = '" & DCNamaMekanik.value & "' WHERE KETERANGAN LIKE '%" & Trim(text(0)) & "%'")
                
        'UPDATE HO
        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE [BENGKEL99].dbo.JOBDTA SET MECH = '" & DCNamaMekanik.value & "' WHERE NO_WO = '" & Trim(text(0)) & "' AND FS_KD_DEALER = '" & Trim(FS_KD_DEALER) & "'")
        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE [BENGKEL99].dbo.T_INVOICENEW SET MECH = '" & DCNamaMekanik.value & "' WHERE NO_WO = '" & Trim(text(0)) & "' AND CABANG = '" & Trim(dBase) & "'")
        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE [BENGKEL99].dbo.HISTORYNEW SET MECH = '" & DCNamaMekanik.value & "' WHERE NO_WO = '" & Trim(text(0)) & "' AND CABANG = '" & Trim(dBase) & "'")
        'lakukan penyimpanan log d profilbengkel cabang
        
        Call objHT.UpdatePerbaikanDataHO("INSERT INTO PROFILBENGKEL..LOGUPDATE VALUES ('UpdateMech','" & Trim(text(0)) & "','" & Trim(text(17)) & "','" & gSA & "',GETDATE(),'" & Trim(text(1)) & "','" & DCNamaMekanik.value & "')")
        
        If erTranAll = False Then
            Call objHT.CloseTransactionPerbaikanDataHO
            Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanNamaMekanik", "gantiNamaMekanik", 1)
            MsgBox "Nama Mekanik berhasil diubah!", vbInformation, "Informasi Berhasil Nama Mekanik"
            Call hapusForm
            text(0).Mask = ""
            Call FetchCombo
            Exit Sub
        Else
            Call objHT.CloseTransactionPerbaikanDataHO_err
            Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanNamaMekanik", "gantiNamaMekanik", 2)
            MsgBox "Terjadi kesalahan saat mengganti Nama Mekanik!", vbInformation, "Gagal Simpan"
            Exit Sub
        End If
vbError:
        Call objHT.CloseTransactionPerbaikanDataHO_err
        Call saveLogMenu(gSA, ComptName, gVersi, "CloseTransactionPerbaikanDataHO_err", "gantiNoRangka", 2)
        MsgBox "Terjadi kesalahan saat mengganti Nama Mekanik!", vbInformation, "Gagal Simpan"
        Exit Sub
    End If
End Sub

Private Sub DCKodeGroup_Change()
    'dealer = ip dealer
    'dbase  = database dealer
    
    isDealerNewEnkripsi = False
    
    FS_KD_DEALER = DCKodeGroup.Columns(0).value
    dealer = DCKodeGroup.Columns(2).value
    dBase = DCKodeGroup.Columns(3).value
    
    If DCKodeGroup.Columns(6).value = "1" Then isDealerNewEnkripsi = True
    
    If isDealerNewEnkripsi Then
        usernameDealer = Dekripsi(DCKodeGroup.Columns(4).value, KataKunci(3))
        passwordDealer = Dekripsi(DCKodeGroup.Columns(5).value, KataKunci(5))
    Else
        usernameDealer = DekripsiGamaLama(DCKodeGroup.Columns(4).value, KataKunciGamaLama(3))
        passwordDealer = DekripsiGamaLama(DCKodeGroup.Columns(5).value, KataKunciGamaLama(5))
    End If
    
End Sub

Private Sub DCKodeGroup_Click()
    Call DCKodeGroup_Change
End Sub

Private Sub DCNamaMekanik_Click()
    'Call DCNamaMekanik_Change
End Sub


Private Sub Form_Load()
    Call FetchCombo
End Sub

Private Sub FetchCombo()
    AdcKode.ConnectionString = SQLConHO
    AdcKode.RecordSource = "select FS_KD_DEALER, FS_NM_DEALER,VPNCABANG,DBASECBG,USERCBG,PASSCBG,NEWENKRIP from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
    DCKodeGroup.DataSourceList = AdcKode
    DCKodeGroup.DataFieldList = "FS_NM_DEALER"
    DCKodeGroup.Columns(0).Width = 1200
    DCKodeGroup.Columns(1).Width = 4000
    DCKodeGroup.Columns(0).Caption = "Kode Dealer"
    DCKodeGroup.Columns(1).Caption = "Nama Dealer"
    DCKodeGroup.Columns(2).Visible = False
    DCKodeGroup.Columns(3).Visible = False
    DCKodeGroup.Columns(4).Visible = False
    DCKodeGroup.Columns(5).Visible = False
    DCKodeGroup.Columns(6).Visible = False
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Text_Validate(Index, True)
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If text(Index) <> "" Then
                Call hapusForm
                Call PencarianNamaMekanik
            Else
                MsgBox "NO WO harus diisi!", vbCritical, "Informasi"
            End If
        
    End Select
End Sub

Private Sub PencarianNamaMekanik()
    Dim noWO As String
    Dim sqlSql As String
    On Error GoTo errHandler
    
    noWO = Trim(text(0))
    isGR = False
    
    Dim rsJobDta As New ADODB.Recordset
    Dim rsHistory As New ADODB.Recordset
    
'    dealer = "10.10.26.171"
'    dbase = "bengkel05"
              
'    usernameGlobal = gUserKoneksi
'    passwordGlobal = gPasswordKonesi

    Dim rs As New ADODB.Recordset
'    Dim rsHO As New ADODB.Recordset
    
'    rsHO.Open "Select * from Koneksi", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameGlobal & ";password=" & passwordGlobal & ";Initial Catalog=koneksi;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
'    rs.Open "Select * from Koneksi", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameGlobal & ";password=" & passwordGlobal & ";Initial Catalog=koneksi;server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
'
'    usernameDealer = Dekripsi(rs.Fields(0).value, KataKunci(3))
'    passwordDealer = Dekripsi(rs.Fields(1).value, KataKunci(5))
    
'    Dim oCr As New Crypt
'    usernameDealer = oCr.Crypt(rs.Fields(0).value, "M4d462", False)
'    passwordDealer = oCr.Crypt(rs.Fields(1).value, "G4j4H", False)
 
    'mengambil data dealer
    rsJobDta.Open "SELECT J.*,C.GSM,C.CDMA FROM JOBDTA J WITH (NOLOCK) INNER JOIN CUSTDAT C WITH (NOLOCK) ON C.NO_CHASSIS = J.NO_CHASSIS WHERE NO_WO = '" & Trim(noWO) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If Not rsJobDta.EOF Then
        'jika ada data cusdat no rangka benar maka tampilkan
        text(4) = rsJobDta("CUSTOMER")
        text(5) = Left(rsJobDta("CUST_ADRES"), 60)
        text(6) = rsJobDta("KOTA")
        text(7) = rsJobDta("TELPON")
        text(8) = rsJobDta("GSM")
        text(9) = rsJobDta("CDMA")
        text(10) = rsJobDta("FS_TH_MOBIL")
        text(11) = rsJobDta("NOINV")
        text(12) = rsJobDta("MDL")
        text(13) = rsJobDta("MODEL")
        text(14) = rsJobDta("FMDL")
        text(15) = rsJobDta("COLOR")
        text(16) = rsJobDta("NO_POLISI")
        text(17) = rsJobDta("NO_CHASSIS")
        text(18) = rsJobDta("NO_ENGINE")
        text(1) = rsJobDta("MECH")
        
        If rsJobDta("FS_RC") = "" Then
            isGR = True
        End If
        
'        strSQL = "SELECT a.*,b.gsm,b.cdma,b.tgl_rep,b.rcdate,b.fs_km_akhir, b.Ddate, b.FS_kd_dealer, b.Fs_nm_dealer " _
'                    & "FROM jobdta a INNER JOIN custdat b ON a.no_chassis = b.no_chassis " _
'                    & "WHERE a.no_chassis = '" & Trim(text(17)) & "' AND a.NO_WO = '" & noWO & "'"
'
'        rsHistory.Open strSQL, ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dbase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
'        If Not rsHistory.EOF Then
        text(2).Mask = ""
        text(2) = TglDDMMYYYY(rsJobDta("date"))
        text(20) = Format(CDbl(rsJobDta("km")), "#,#")
'        End If
        Label8.Visible = True
        DCNamaMekanik.Visible = True
        DCNamaMekanik.SetFocus
        cmdGanti.Visible = True
        Call FetchComboMekanik
    Else
        MsgBox "Tidak ada NO WO yang dicari", vbOKOnly, "Informasi"
    End If
    Exit Sub
errHandler:
    MsgBox "Terjadi kesalahan dalam mengambil data!", vbCritical, "Informasi" 'err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub FetchComboMekanik()
    koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer
    AdcMek.ConnectionString = koneksi
    AdcMek.RecordSource = "select MECH, NAME from MEKANIK WITH (NOLOCK) ORDER BY MECH"
    DCNamaMekanik.DataSourceList = AdcMek
    DCNamaMekanik.DataFieldList = "MECH"
    DCNamaMekanik.Columns(0).Width = 1200
    DCNamaMekanik.Columns(1).Width = 4000
    DCNamaMekanik.Columns(0).Caption = "Kode Mekanik"
    DCNamaMekanik.Columns(1).Caption = "Nama Mekanik"
End Sub

Private Sub hapusForm()
    For i = 1 To 20
        If i = 11 Or i = 2 Then
        text(i).Mask = ""
        ElseIf i = 3 Or i = 19 Then
        'text tersebut tidak ada index
        Else
        text(i) = ""
        End If
    Next
    Label8.Visible = False
    'text(1).Visible = False
    DCNamaMekanik.Visible = False
    cmdGanti.Visible = False
    
End Sub
