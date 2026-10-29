VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOPerbaikanNoRangka 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PERBAIKAN NO RANGKA"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7800
      TabIndex        =   36
      Top             =   4800
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   8438015
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
      Left            =   5880
      TabIndex        =   40
      Top             =   4800
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Ganti No Rangka"
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
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11775
      Begin MSAdodcLib.Adodc AdcKode 
         Height          =   330
         Left            =   8760
         Top             =   720
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
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   11400
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   4
            Left            =   930
            TabIndex        =   4
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
            TabIndex        =   5
            Top             =   795
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
            TabIndex        =   6
            Top             =   1110
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
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   33
            Top             =   0
            Width           =   3675
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   12600
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   12600
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
            Width           =   12855
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   2280
            Width           =   510
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DO TAM"
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
            TabIndex        =   22
            Top             =   840
            Width           =   585
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
         TabIndex        =   1
         Top             =   1200
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   1
         Left            =   2520
         TabIndex        =   38
         Top             =   4440
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERBAIKAN NOMOR RANGKA"
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
         Left            =   3510
         TabIndex        =   41
         Top             =   120
         Width           =   4125
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BorderColor     =   &H80000002&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   15
         Width           =   11760
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rangka Benar"
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
         Left            =   405
         TabIndex        =   39
         Top             =   4440
         Width           =   2025
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rangka Salah"
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
         Left            =   480
         TabIndex        =   35
         Top             =   1200
         Width           =   1980
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
         Left            =   1200
         TabIndex        =   34
         Top             =   780
         Width           =   1215
      End
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   37
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
Attribute VB_Name = "frmHOPerbaikanNoRangka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dealer               As String
Private dBase               As String
Private koneksi             As String
Private usernameDealer      As String
Private passwordDealer      As String
Private usernameHO          As String
Private passwordHO          As String
Private usernameGlobal      As String
Private passwordGlobal      As String
Private isDoubleCustDat     As Boolean
Private isDoubleCustDatHO   As Boolean
Private isDoubleUC          As Boolean

Private isDealerNewEnkripsi As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGanti_Click()
1         If Len(Trim(Text(0))) = 0 Then
2             MsgBox "No Rangka yang salah harus diisi!", vbCritical, "Informasi"
3             Exit Sub
4         ElseIf Len(Trim(Text(1))) = 0 Then
5             MsgBox "No Rangka yang salah harus diisi!", vbCritical, "Informasi"
6             Exit Sub
7         End If
          'Call hapusForm
8         If PencarianNoRangka = False Then Exit Sub
          
9         If MsgBox("Apakah yakin mengganti No Rangka " & Trim(Text(0)) & " dengan " & Trim(Text(1)) & " ?", vbYesNo, "Konfirmasi Penggantian No Rangka") = vbYes Then
              'lakukan pergantian data
10            On Error GoTo vbError
11            koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer
12            erTranAll = False
              'perbaikan kecabang
13            Call objHT.OpenTransactionPerbaikanDataHO(koneksi)
14            If isDoubleCustDat Then
                  'lakukan hapus data custdat no rangka salah
15                Call objHT.UpdatePerbaikanDataHO("DELETE CUSTDAT WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
16            End If
              '========:: 2015-07-01 ::============
              ''Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDTL SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_WO IN (SELECT NO_WO FROM JOBDTA WHERE NO_POLISI = '" & Trim(Text(16)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "')")
              
              ''20220907 perbaikan noka tanpa menggunakan nopolisi
              ''-------------------------------------------------------
17            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDTL SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_WO IN (SELECT NO_WO FROM JOBDTA WHERE NO_CHASSIS = '" & Trim(Text(0)) & "')")
              
              '====================================
              'ganti no rangka
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE CUSTDAT SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_POlISI = '" & Trim(Text(16)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE JOBSUBLET SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_WO IN (SELECT NO_WO FROM JOBDTA WHERE NO_POLISI = '" & Trim(Text(16)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "')")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDTA SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_POLISI = '" & Trim(Text(16)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE T_INVOICE SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_POLISI = '" & Trim(Text(16)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE HISTORY SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_POLISI = '" & Trim(Text(16)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE BOOK SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_POLISI = '" & Trim(Text(16)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE DTLFU SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_POLISI = '" & Trim(Text(16)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE ESTIMASI SET NO_RANGKA = '" & Trim(Text(1)) & "' WHERE NO_RANGKA = '" & Trim(Text(0)) & "' AND NO_POLISI = '" & Trim(Text(16)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE HIST_PEMILIK SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE HIST_SUBLET SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE HIST_JOBADD SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE JOBKOORD SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE JOBPART SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE M_ESTIMASI SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "' AND NOPOLISI = '" & Trim(Text(16)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE M_LISTREMINDER SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE M_LISTREMINDERCALL SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE PROGRESS SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHO("UPDATE VALIDORDER SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
      '
              ''20220907 perbaikan noka tanpa menggunakan nopolisi
              ''-------------------------------------------------------
18            Call objHT.UpdatePerbaikanDataHO("UPDATE CUSTDAT SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
19            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBSUBLET SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_WO IN (SELECT NO_WO FROM JOBDTA WHERE NO_CHASSIS = '" & Trim(Text(0)) & "')")
20            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDTA SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
21            Call objHT.UpdatePerbaikanDataHO("UPDATE T_INVOICE SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
22            Call objHT.UpdatePerbaikanDataHO("UPDATE HISTORY SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
23            Call objHT.UpdatePerbaikanDataHO("UPDATE BOOK SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' ")
24            Call objHT.UpdatePerbaikanDataHO("UPDATE DTLFU SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' ")
25            Call objHT.UpdatePerbaikanDataHO("UPDATE ESTIMASI SET NO_RANGKA = '" & Trim(Text(1)) & "' WHERE NO_RANGKA = '" & Trim(Text(0)) & "' ")
26            Call objHT.UpdatePerbaikanDataHO("UPDATE HIST_PEMILIK SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
27            Call objHT.UpdatePerbaikanDataHO("UPDATE HIST_SUBLET SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
28            Call objHT.UpdatePerbaikanDataHO("UPDATE HIST_JOBADD SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
29            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBKOORD SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
30            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBPART SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
31            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ESTIMASI SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "' ")
32            Call objHT.UpdatePerbaikanDataHO("UPDATE M_LISTREMINDER SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
33            Call objHT.UpdatePerbaikanDataHO("UPDATE M_LISTREMINDERCALL SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
34            Call objHT.UpdatePerbaikanDataHO("UPDATE PROGRESS SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
35            Call objHT.UpdatePerbaikanDataHO("UPDATE VALIDORDER SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
              ''2023-01-10 penambahan update ganti noka
              
36            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDETAIL SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
37            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBADD2 SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
              
38            Call objHT.UpdatePerbaikanDataHO("UPDATE DIAGADD2 SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
39            Call objHT.UpdatePerbaikanDataHO("UPDATE DIAGDTA SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
40            Call objHT.UpdatePerbaikanDataHO("UPDATE DIAGDTL SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
41            Call objHT.UpdatePerbaikanDataHO("UPDATE DIAGKOORD SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
42            Call objHT.UpdatePerbaikanDataHO("UPDATE DIAGPART SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
43            Call objHT.UpdatePerbaikanDataHO("UPDATE DIAGSUBLET SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
              
44            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ListReminderCR SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
45            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ListReminderCallCR SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
              
46            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ListReminderLub SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
47            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ListReminderCallLub SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
              
48            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ListReminderSSC SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
49            Call objHT.UpdatePerbaikanDataHO("UPDATE M_ListReminderCallMapping SET NORANGKA = '" & Trim(Text(1)) & "' WHERE NORANGKA = '" & Trim(Text(0)) & "'")
              
      '        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE JOBDTA SET NO_CHASSIS = '" & Trim(text(1)) & "' WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE T_INVOICENEW SET NO_CHASSIS = '" & Trim(text(1)) & "' WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE HISTORYNEW SET NO_CHASSIS = '" & Trim(text(1)) & "' WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
      '        Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE CUSTDAT SET NO_CHASSIS = '" & Trim(text(1)) & "' WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
              
              ''2018-08-31
              ''=========================================================
              
50            Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.JOBDTA SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
51            Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.T_INVOICENEW SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
52            Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.HISTORYNEW SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
              ''Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.CUSTDAT SET NO_CHASSIS = '" & Trim(text(1)) & "' WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
              
              ''20181217 | update singleidori untuk unitcustomer
              '' ------------------------------------------------------------
53            If isDoubleUC Then
54                Call objHT.UpdateData_Ho("DELETE FROM [SingleIDOri].dbo.UNITCUSTOMER WHERE NoRangka = '" & Trim(Text(0)) & "' ")
55            End If
56            Call objHT.UpdateData_Ho("UPDATE [SingleIDOri].dbo.UNITCUSTOMER SET NoRangka = '" & Trim(Text(1)) & "' WHERE NoRangka = '" & Trim(Text(0)) & "' ")
              ''=========================================================
              
              
57            Call objHT.UpdateData_Ho("UPDATE JOBDTA SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
58            Call objHT.UpdateData_Ho("UPDATE T_INVOICENEW SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
59            Call objHT.UpdateData_Ho("UPDATE HISTORYNEW SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
60            Call objHT.UpdateData_Ho("UPDATE CUSTDAT SET NO_CHASSIS = '" & Trim(Text(1)) & "' WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
              
61            If isDoubleCustDatHO Then
                  'lakukan hapus data custdat no rangka salah HO
                  'Call objHT.UpdatePerbaikanDataHoConnectionHO("DELETE CUSTDAT WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
62                Call objHT.UpdateData_Ho("DELETE CUSTDAT WHERE NO_CHASSIS = '" & Trim(Text(0)) & "'")
63            End If
              'lakukan penyimpanan log d profilbengkel cabang
              
64            Call objHT.UpdatePerbaikanDataHO("INSERT INTO PROFILBENGKEL..LOGUPDATE VALUES ('UpdateNoka','','" & Trim(Text(0)) & "','" & gSA & "',GETDATE(),'" & Trim(Text(1)) & "','" & Trim(Text(0)) & "')")
              
65            If erTranAll = False Then
66                Call objHT.CloseTransactionPerbaikanDataHO
67                Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanNoRangka", "gantiNoRangka", 1)
68                MsgBox "No Rangka berhasil diubah!", vbInformation, "Informasi Berhasil Ganti No Rangka"
69                Call hapusForm
70                Text(0) = ""
71                Call FetchCombo
72                Exit Sub
73            Else
74                Call objHT.CloseTransactionPerbaikanDataHO_err
75                Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanNoRangka", "gantiNoRangka", 2)
76                MsgBox "Terjadi kesalahan saat mengganti No Rangka!", vbInformation, "Gagal Simpan"
77                Exit Sub
78            End If
vbError:
79            Call objHT.CloseTransactionPerbaikanDataHO_err
80            Call saveLogMenu(gSA, ComptName, gVersi, "CloseTransactionPerbaikanDataHO_err", "gantiNoRangka", 2)
81            MsgBox "Terjadi kesalahan saat mengganti No Rangka!", vbInformation, "Gagal Simpan"
82            Exit Sub
83        End If
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
        usernameDealer = "sa" 'DekripsiGamaLama(DCKodeGroup.Columns(4).value, KataKunciGamaLama(3))
        passwordDealer = "nasmoco~123" 'DekripsiGamaLama(DCKodeGroup.Columns(5).value, KataKunciGamaLama(5))
    End If
End Sub

Private Sub DCKodeGroup_Click()
    Call DCKodeGroup_Change
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
        Call text_Validate(Index, True)
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If dealer <> "" Then
                If Text(Index) <> "" Then
                    Call hapusForm
                    Call PencarianNoRangkaSalah
                Else
                    MsgBox "No Rangka harus diisi!", vbCritical, "Informasi"
                End If
            Else
                MsgBox "Dealer harus dipilih!", vbCritical, "Informasi"
            End If
    End Select
End Sub

Private Sub PencarianNoRangkaSalah()
    Dim noRangkaSalah As String
    Dim noRangkaBenar As String
    
    On Error GoTo errHandler
    
    noRangkaSalah = Trim(Text(0))
    'noRangkaBenar = Trim(text(1))
    
    Dim rsCustdatSalah As New ADODB.Recordset
    Dim rsCustdatBenar As New ADODB.Recordset
    Dim rsCustdatSalahHO As New ADODB.Recordset
    Dim rsCustdatBenarHO As New ADODB.Recordset
    
'    dealer = "10.10.26.171"
'    dbase = "bengkel11"
        
    isDoubleCustDat = False
    isDoubleCustDatHO = False
    isDoubleUC = False
              
'    usernameGlobal = gUserKoneksi
'    passwordGlobal = gPasswordKonesi

    Dim rs As New ADODB.Recordset
    Dim rsHO As New ADODB.Recordset
    
'    rsHO.Open "Select * from Koneksi", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameGlobal & ";password=" & passwordGlobal & ";Initial Catalog=koneksi;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
'    rs.Open "Select * from Koneksi", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameGlobal & ";password=" & passwordGlobal & ";Initial Catalog=koneksi;server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
'
'    usernameDealer = Dekripsi(rs.Fields(0).value, KataKunci(3))
'    passwordDealer = Dekripsi(rs.Fields(1).value, KataKunci(5))
'
'    usernameHO = Dekripsi(rsHO.Fields(0).value, KataKunci(3))
'    passwordHO = Dekripsi(rsHO.Fields(1).value, KataKunci(5))
    
'    Dim oCr As New Crypt
'    usernameDealer = oCr.Crypt(rs.Fields(0).value, "M4d462", False)
'    passwordDealer = oCr.Crypt(rs.Fields(1).value, "G4j4H", False)
'
'    usernameHO = oCr.Crypt(rsHO.Fields(0).value, "M4d462", False)
'    passwordHO = oCr.Crypt(rsHO.Fields(1).value, "G4j4H", False)
    
    'mengambil data dealer
    rsCustdatSalah.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    'rsCustdatBenar.Open "SELECT * FROM CUSTDAT WHERE NO_CHASSIS = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dbase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    
    'mengambil data HO
    Set rsCustdatSalahHO = objO.FetchDataHO("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'")
    'rsCustdatSalahHO.Open "SELECT * FROM CUSTDAT WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameHO & ";password=" & passwordHO & ";Initial Catalog=BENGKEL99;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
    'rsCustdatBenarHO.Open "SELECT * FROM CUSTDAT WHERE NO_CHASSIS = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameHO & ";password=" & passwordHO & ";Initial Catalog=BENGKEL99;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
    
'    If Not rsCustdatBenarHO.EOF Then
'        If Not rsCustdatSalahHO.EOF Then
'            isDoubleCustDatHO = True
'        End If
'    End If
    
   If Not rsCustdatSalah.EOF Then
    
        Text(4) = rsCustdatSalah("CUSTOMER")
        Text(5) = Left(rsCustdatSalah("CUST_ADRES"), 60)
        Text(6) = rsCustdatSalah("KOTA")
        Text(7) = rsCustdatSalah("TELPON")
        Text(8) = rsCustdatSalah("GSM")
        Text(9) = rsCustdatSalah("CDMA")
        Text(10) = rsCustdatSalah("FS_TH_MOBIL")
        Text(11).Mask = ""
        Text(11) = Nul(IIf(Replace(rsCustdatSalah("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCustdatSalah("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsCustdatSalah("DOTAM"), "/", "-"))))
        Text(12) = rsCustdatSalah("MDL")
        Text(13) = rsCustdatSalah("MODEL")
        Text(14) = rsCustdatSalah("FMDL")
        Text(15) = rsCustdatSalah("COLOR")
        Text(16) = rsCustdatSalah("NO_POLISI")
        Text(17) = rsCustdatSalah("NO_CHASSIS")
        Text(18) = rsCustdatSalah("NO_ENGINE")
        Label8.Visible = True
        Text(1).Visible = True
        cmdGanti.Visible = True
        Text(1).SetFocus
        
    Else
        MsgBox "Tidak ada No Rangka yang dicari", vbOKOnly, "Informasi"
    End If
    Exit Sub
errHandler:
    MsgBox "Terjadi kesalahan dalam mengambil data!", vbCritical, "Informasi" 'err.Source & ":" & vbCrLf & err.Description
End Sub

Private Function PencarianNoRangka() As Boolean
    Dim noRangkaSalah As String
    Dim noRangkaBenar As String
    
    On Error GoTo errHandler
    PencarianNoRangka = True
    noRangkaSalah = Trim(Text(0))
    noRangkaBenar = Trim(Text(1))
    
    Dim rsCustdatSalah      As New ADODB.Recordset
    Dim rsCustdatBenar      As New ADODB.Recordset
    Dim rsUCSalah           As New ADODB.Recordset
    Dim rsUCBenar           As New ADODB.Recordset
    Dim rsCustdatSalahHO    As New ADODB.Recordset
    Dim rsCustdatBenarHO    As New ADODB.Recordset
    
    
'    dealer = "10.10.26.171"
'    dbase = "bengkel11"
        
    isDoubleCustDat = False
    isDoubleCustDatHO = False
              
'    usernameGlobal = gUserKoneksi
'    passwordGlobal = gPasswordKonesi

    Dim rs As New ADODB.Recordset
    Dim rsHO As New ADODB.Recordset
    
'    rsHO.Open "Select * from Koneksi", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameGlobal & ";password=" & passwordGlobal & ";Initial Catalog=koneksi;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
'    rs.Open "Select * from Koneksi", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameGlobal & ";password=" & passwordGlobal & ";Initial Catalog=koneksi;server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
'
'    usernameDealer = Dekripsi(rs.Fields(0).value, KataKunci(3))
'    passwordDealer = Dekripsi(rs.Fields(1).value, KataKunci(5))
'
'    usernameHO = Dekripsi(rsHO.Fields(0).value, KataKunci(3))
'    passwordHO = Dekripsi(rsHO.Fields(1).value, KataKunci(5))
    
'    Dim oCr As New Crypt
'    usernameDealer = oCr.Crypt(rs.Fields(0).value, "M4d462", False)
'    passwordDealer = oCr.Crypt(rs.Fields(1).value, "G4j4H", False)
'
    usernameHO = Log.LogHOUser 'oCr.Crypt(rsHO.Fields(0).value, "M4d462", False)
    passwordHO = Log.LogHOPass 'oCr.Crypt(rsHO.Fields(1).value, "G4j4H", False)
    
    'mengambil data dealer
    rsCustdatSalah.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    rsCustdatBenar.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    
    ''singlidori cabang
    rsUCSalah.Open "SELECT * FROM UNITCUSTOMER WITH (NOLOCK) WHERE NoRangka = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=SINGLEIDORI;server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    rsUCBenar.Open "SELECT * FROM UNITCUSTOMER WITH (NOLOCK) WHERE NoRangka = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=SINGLEIDORI;server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
    
    'mengambil data HO
'    rsCustdatSalahHO.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameHO & ";password=" & passwordHO & ";Initial Catalog=BENGKEL99;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
'    rsCustdatBenarHO.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameHO & ";password=" & passwordHO & ";Initial Catalog=BENGKEL99;server=10.10.11.191"), adOpenStatic, adLockBatchOptimistic, adCmdText
    ''20181022
    rsCustdatSalahHO.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameHO & ";password=" & passwordHO & ";Initial Catalog=BENGKEL99;server=" & gServerDb), adOpenStatic, adLockBatchOptimistic, adCmdText
    rsCustdatBenarHO.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameHO & ";password=" & passwordHO & ";Initial Catalog=BENGKEL99;server=" & gServerDb), adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If Not rsCustdatBenarHO.EOF Then
        If Not rsCustdatSalahHO.EOF Then
            isDoubleCustDatHO = True
        End If
    End If
    
    If Not rsCustdatBenar.EOF Then
        If Not rsCustdatSalah.EOF Then
            isDoubleCustDat = True
        End If
    ElseIf Not rsCustdatSalah.EOF Then
        
    Else
        MsgBox "Tidak ada No Rangka yang dicari", vbOKOnly, "Informasi"
    End If
    
    If Not rsUCBenar.EOF Then
        If Not rsUCSalah.EOF Then
            isDoubleUC = True
        End If
    End If
    
    Exit Function
errHandler:
    MsgBox "Terjadi kesalahan dalam mengambil data!", vbCritical, "Informasi" 'err.Source & ":" & vbCrLf & err.Description
    PencarianNoRangka = False
End Function

Private Sub hapusForm()
    For i = 4 To 18
        If i = 11 Then
        Text(i).Mask = ""
        Else
        Text(i) = ""
        End If
    Next
    Text(1) = ""
    Label8.Visible = False
    Text(1).Visible = False
    cmdGanti.Visible = False
End Sub
