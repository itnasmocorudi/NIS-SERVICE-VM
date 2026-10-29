VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form FrmOrderListUnitCome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PERBAIKAN PANGGILAN SELESAI SERVICE"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   12600
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9480
      TabIndex        =   2
      Top             =   6000
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Kembali Menu"
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7320
      TabIndex        =   1
      Top             =   6000
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Follow"
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
      Width           =   11200
      _ExtentX        =   19764
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   -600
      TabIndex        =   0
      Top             =   480
      Width           =   11835
      Begin VB.Frame frameLegend 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LEGEND WARNA"
         Height          =   1455
         Left            =   4440
         TabIndex        =   34
         Top             =   3960
         Visible         =   0   'False
         Width           =   3495
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "BELUM PROSES"
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
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label10 
            BackColor       =   &H0000FFFF&
            Caption         =   "TELAH FOLLOW"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C000&
            Caption         =   "TELAH PROSES"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   3375
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "FILTER DATA"
         Height          =   855
         Left            =   600
         TabIndex        =   25
         Top             =   600
         Width           =   11175
         Begin VB.CheckBox ckTelahProses 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Telah Proses"
            Height          =   255
            Left            =   7800
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox ckTelahFollow 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Telah Follow"
            Height          =   255
            Left            =   6480
            TabIndex        =   31
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox ckBelumProses 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Belum Proses"
            Height          =   255
            Left            =   5160
            TabIndex        =   30
            Top             =   360
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin TDBDate6Ctl.TDBDate Tgl1 
            Height          =   330
            Left            =   1680
            TabIndex        =   27
            Top             =   300
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   582
            Calendar        =   "frmOrderListUnitCome.frx":0000
            Caption         =   "frmOrderListUnitCome.frx":012C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmOrderListUnitCome.frx":0198
            Keys            =   "frmOrderListUnitCome.frx":01B6
            Spin            =   "frmOrderListUnitCome.frx":0214
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
            Left            =   3600
            TabIndex        =   28
            Top             =   300
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   582
            Calendar        =   "frmOrderListUnitCome.frx":023C
            Caption         =   "frmOrderListUnitCome.frx":0368
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmOrderListUnitCome.frx":03D4
            Keys            =   "frmOrderListUnitCome.frx":03F2
            Spin            =   "frmOrderListUnitCome.frx":0450
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
         Begin AIFCmp1.asxPowerButton asxPowerButton5 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   9240
            TabIndex        =   33
            Top             =   300
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Refresh Grid"
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
         Begin VB.Label Label7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "S.D"
            Height          =   255
            Left            =   3240
            TabIndex        =   29
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "TANGGAL MASUK"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FOLLOW UNIT"
         Height          =   3255
         Left            =   3240
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   5775
         Begin VB.ComboBox cbFollow 
            Height          =   315
            ItemData        =   "frmOrderListUnitCome.frx":0478
            Left            =   1680
            List            =   "frmOrderListUnitCome.frx":0482
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1700
            Width           =   2175
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   3
            Left            =   1680
            TabIndex        =   9
            Top             =   480
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   9
            Left            =   1680
            TabIndex        =   10
            Top             =   1065
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   10
            Left            =   1680
            TabIndex        =   11
            Top             =   780
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
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
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   1380
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   35
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "999999999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   17
            Top             =   180
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   21
            Top             =   2040
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton3 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   3240
            TabIndex        =   23
            Top             =   2520
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Simpan"
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
         Begin AIFCmp1.asxPowerButton asxPowerButton4 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   4320
            TabIndex        =   24
            Top             =   2520
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Batal"
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
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal FU"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   705
            TabIndex        =   22
            Top             =   2040
            Width           =   930
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hasil FU"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   870
            TabIndex        =   20
            Top             =   1700
            Width           =   660
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No Polisi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jenis Kendaraan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   780
            Width           =   1395
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   735
            TabIndex        =   15
            Top             =   1065
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   495
            Width           =   1425
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Telepon"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   585
            TabIndex        =   13
            Top             =   1380
            Width           =   930
         End
      End
      Begin AIFCmp1.asxPowerButton asxPowerButton2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   9000
         TabIndex        =   5
         Top             =   5520
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Proses"
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
      Begin TrueOleDBGrid70.TDBGrid unitGrid 
         Height          =   3885
         Left            =   600
         TabIndex        =   6
         Top             =   1560
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "FormatText Event"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "FormatText Event"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "FormatText Event"
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
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).DataField=   ""
         Columns(11).NumberFormat=   "FormatText Event"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(23)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(27)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(31)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(35)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(39)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(43)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(47)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "Unit Come"
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=82,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=74,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=60,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=57,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=58,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=59,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=56,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=53,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=54,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=55,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=64,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=61,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=62,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=63,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=85,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=70,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=83,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=84,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=69,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=66,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=67,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=68,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13,.locked=0"
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
         _StyleDefs(105) =   "Named:id=51:merah"
         _StyleDefs(106) =   ":id=51,.parent=33,.bgcolor=&HFFFFFF&,.fgcolor=&HFF&,.bold=-1,.fontsize=825"
         _StyleDefs(107) =   ":id=51,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(108) =   ":id=51,.fontname=MS Sans Serif"
         _StyleDefs(109) =   "Named:id=52:kuning"
         _StyleDefs(110) =   ":id=52,.parent=51,.bgcolor=&HFFFF&,.fgcolor=&H0&"
         _StyleDefs(111) =   "Named:id=65:hijau"
         _StyleDefs(112) =   ":id=65,.parent=52,.bgcolor=&HFF00&"
      End
      Begin ucXPButton.XPButton btnLegend 
         Height          =   495
         Left            =   4440
         TabIndex        =   38
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
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
         MICON           =   "frmOrderListUnitCome.frx":04A2
         PICN            =   "frmOrderListUnitCome.frx":04BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblJumlah 
         BackColor       =   &H00FFFFFF&
         Caption         =   "JUMLAH DATA = "
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
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   5520
         Width           =   4335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIST UNIT COME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4560
         TabIndex        =   4
         Top             =   120
         Width           =   2340
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BorderColor     =   &H80000002&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11205
      End
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
Attribute VB_Name = "FrmOrderListUnitCome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private noPol As String
Private isBoleh As Boolean
Private rsCustomer As ADODB.Recordset
Private pesanError As String
Dim cols As TrueOleDBGrid70.Columns
Private tanggalInput As String

Private isDealerNewEnkrip As Boolean


Private Sub refreshGrid()
1        On Error GoTo refreshGrid_Error

2         unitGrid.DataSource = Nothing
          Dim filter As String
          Dim filter2 As String
3         filter = ""
4         filter2 = ""
          
5         If Tgl1 <> "" And Tgl2 <> "" Then
6             If filter = "" Then filter = "WHERE "
7             filter = filter & " tanggalin BETWEEN '" & Format(Tgl1, "YYYY-MM-DD") & " 00:00" & "' AND '" & Format(Tgl2, "YYYY-MM-DD") & " 23:59" & "' "
8         End If
          
9         If ckBelumProses.value = vbChecked Then
10            If filter2 = "" Then
11                filter2 = filter2 & " (STATUSFU = '0' and STATUSPROSES = '0') "
12            Else
13                filter2 = filter2 & " or (STATUSFU = '0' and STATUSPROSES = '0') "
14            End If
15        End If
16        If ckTelahFollow.value = vbChecked Then
17            If filter2 = "" Then
18                filter2 = filter2 & " (STATUSFU = '1') "
19            Else
20                filter2 = filter2 & " or (STATUSFU = '1' " & IIf(ckTelahProses.value = vbChecked, " and STATUSPROSES = '1'", " and STATUSPROSES = '0'") & ") "
21            End If
22        End If
23        If ckTelahProses.value = vbChecked Then
24            If filter2 = "" Then
25                filter2 = filter2 & " (STATUSPROSES = '1') "
26            Else
27                filter2 = filter2 & " OR (STATUSPROSES = '1') "
28            End If
29        End If
          
30            If filter2 <> "" Then
31                filter = filter & " and ( " & filter2 & " )"
32            End If
          'Set rsCustomer = objO.FetchData("SELECT *,notelp1 + ' | ' + notelp2 as notelp,convert(varchar(100),tanggalin,20) tanggalincon FROM UNITCOME WHERE STATUSPROSES = '0' AND STATUSFU = '0'")
33        Set rsCustomer = objO.FetchData("SELECT distinct u.*,u.notelp1 + ' | ' + u.notelp2 as notelp,convert(varchar(100),u.tanggalin,20) tanggalincon,isnull(c.color,'') color FROM UNITCOME u left outer join custdat c on c.no_polisi = u.nopolisi " & filter)
34        If Not rsCustomer.EOF Then
35            Set unitGrid.DataSource = rsCustomer
36            unitGrid.FetchRowStyle = True
37            Call headGrid
38            lblJumlah.Caption = "JUMLAH DATA = " & rsCustomer.RecordCount
39        End If

40       On Error GoTo 0
41       Exit Sub

refreshGrid_Error:

42        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form FrmOrderListUnitCome")
End Sub

Private Sub headGrid()
1        On Error GoTo headGrid_Error

2         With unitGrid
3             .Columns(5).Caption = "Nama Customer"
4             .Columns(5).DataField = "NamaCustomer"
5             .Columns(5).Width = 3000
6             .Columns(5).Locked = True
              
7             .Columns(0).Caption = "Tanggal + Jam In"
8             .Columns(0).DataField = "tanggalincon"
9             .Columns(0).Width = 2500
10            .Columns(0).Locked = True
              
11            .Columns(1).Caption = "No Polisi"
12            .Columns(1).DataField = "NoPolisi"
13            .Columns(1).Width = 2000
14            .Columns(1).Locked = True
              
15            .Columns(2).Caption = "Jenis"
16            .Columns(2).DataField = "JenisKendaraan"
17            .Columns(2).Width = 2000
18            .Columns(2).Locked = True
              
19            .Columns(3).Caption = "Tahun"
20            .Columns(3).DataField = "TahunKendaraan"
21            .Columns(3).Width = 800
22            .Columns(3).Locked = True
              
23            .Columns(4).Caption = "Warna"
24            .Columns(4).DataField = "Color"
25            .Columns(4).Width = 800
26            .Columns(4).Locked = True
              
27            .Columns(6).Caption = "No Telepon"
28            .Columns(6).DataField = "notelp"
29            .Columns(6).Width = 2000
30            .Columns(6).Locked = True
              
31            .Columns(7).Caption = "Keterangan"
32            .Columns(7).DataField = "keterangan"
33            .Columns(7).Width = 4000
34            .Columns(7).Locked = True
              
35            .Columns(8).Caption = "StatusFU"
36            .Columns(8).DataField = "StatusFU"
37            .Columns(8).Width = 400
38            .Columns(8).Visible = False
              
39            .Columns(9).Caption = "StatusProses"
40            .Columns(9).DataField = "StatusProses"
41            .Columns(9).Width = 400
42            .Columns(9).Visible = False
              
43            .Columns(10).Caption = "Tanggal Follow"
44            .Columns(10).DataField = "tanggalfu"
45            .Columns(10).Width = 4000
46            .Columns(10).Locked = True
              
47            .Columns(11).Caption = "Tanggal Proses"
48            .Columns(11).DataField = "TanggalProses"
49            .Columns(11).Width = 4000
50            .Columns(11).Locked = True
51        End With

52       On Error GoTo 0
53       Exit Sub

headGrid_Error:

54        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headGrid of Form FrmOrderListUnitCome")
End Sub

Private Sub asxPowerButton2_Click()
          Dim noRangka As String
          Dim rsFollow As ADODB.Recordset
          Dim rsOrder As ADODB.Recordset
          Dim selisih As Integer
1        On Error GoTo asxPowerButton2_Click_Error

2         noRangka = objO.FetchData_String("SELECT NO_CHASSIS FROM CUSTDAT WHERE NO_POLISI = '" & Trim(text(0)) & "'")
3         If unitGrid.Columns(8).value = "0" Then
4             MsgBox "Lakukan proses Follow terlebih dahulu", vbInformation
5             Exit Sub
6         End If
7         If noRangka <> "" Then
8             If gFiskal = Format(Month(isnow), "00") Then
                  'nUnit = isiCustFU(gSA)
9                 If Trim(gTipeUser) = "SA" And Log.followUp = 1 Then
                     ''2016-07-11 || pengecekan h+1 follow up berdasarkan kode sa tersebut
                     ''jumlah order untuk follow up +1 hari
10                   Set rsOrder = objO.FetchData("select count(l.no_wo) total from JOBDTL l with(nolock)  inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "'")
                     
                     ''jumlah data follow untuk tglfu +1 hari
11                   Set rsFollow = objO.FetchData("select COUNT(NO_WO) total from DTLFU where NO_WO in (select l.no_wo from JOBDTL l with(nolock)  inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "')")
12                   selisih = 0
                     
13                   If Not rsOrder.EOF Then
14                      selisih = CInt(rsOrder!total) ''CInt(rsOrder.RecordCount) - CInt(rsFollow.RecordCount)
15                   End If
                     
16                   If Not rsFollow.EOF Then
17                      selisih = selisih - CInt(rsFollow!total)
18                   End If
                     
19                   If selisih > 0 Then
                        ''munculkan peringatan follow up
20                      MsgBox "Menu order tidak dapat dibuka." & vbCrLf & "Masih ada " & selisih & " no wo yang belum difollow up", vbInformation
                        ''Unload Me
21                      Exit Sub
22                   End If
23                   frmOrder.noRangkaUnitCome = ""
24                   frmOrder.noRangkaUnitCome = noRangka
                     
25                   frmOrder.noPolisiUnitCome = ""
26                   frmOrder.noPolisiUnitCome = Trim(text(0))
                     
27                   frmOrder.tanggalInUnitCome = ""
28                   frmOrder.tanggalInUnitCome = tanggalInput
                     
29                   frmOrder.Show
30                   frmOrder.WindowState = vbMaximized
                     
31                Else
32                   If gLimitFUDealer - nUnit <= 10 And gLimitFUDealer - nUnit > 0 Then
33                    MsgBox "Hanya bisa menerima " & CDbl(gLimitFUDealer - nUnit) & " Order lagi. " & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
34                   ElseIf gLimitFUDealer <= nUnit Then
35                       MsgBox "Tidak bisa menerima Order lagi. " & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
36                   End If
37                   If nUnit < gLimitFUDealer Then
38                       frmOrder.noRangkaUnitCome = ""
39                       frmOrder.noRangkaUnitCome = noRangka
                         
40                       frmOrder.noPolisiUnitCome = ""
41                       frmOrder.noPolisiUnitCome = Trim(text(0))
                         
42                       frmOrder.tanggalInUnitCome = ""
43                       frmOrder.tanggalInUnitCome = tanggalInput
                         
44                       frmOrder.Show
45                       frmOrder.WindowState = vbMaximized
46                   ElseIf nUnit > gLimitFqUDealer Then
47                       MsgBox "Customer yang belum di-Follow telah mencapai batas " & gLimitFUDealer & " unit" & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
48                   End If
49                End If
                  
50            Else
51                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
52            End If
53        Else
54            If gFiskal = Format(Month(gTglTampung), "00") Then
                  'nUnit = isiCustFU(gSA)
55                If Trim(gTipeUser) = "SA" And Log.followUp = 1 Then
                       ''2016-07-11 || pengecekan h+1 follow up berdasarkan kode sa tersebut
                       ''jumlah order untuk follow up +1 hari
56                     Set rsOrder = objO.FetchData("select count(l.no_wo) total from JOBDTL l with(nolock)  inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "'")
                       
                       ''jumlah data follow untuk tglfu +1 hari
57                     Set rsFollow = objO.FetchData("select COUNT(NO_WO) total from DTLFU where NO_WO in (select l.no_wo from JOBDTL l with(nolock)  inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "')")
58                     selisih = 0
                       
59                     If Not rsOrder.EOF Then
60                        selisih = CInt(rsOrder!total) ''CInt(rsOrder.RecordCount) - CInt(rsFollow.RecordCount)
61                     End If
                       
62                     If Not rsFollow.EOF Then
63                        selisih = selisih - CInt(rsFollow!total)
64                     End If
                       
65                     If selisih > 0 Then
                          ''munculkan peringatan follow up
66                        MsgBox "Menu order tidak dapat dibuka." & vbCrLf & "Masih ada " & selisih & " no wo yang belum difollow up", vbInformation
                          ''Unload Me
67                        Exit Sub
68                     End If
69                      frmOrder.noRangkaUnitCome = ""
70                      frmOrder.noRangkaUnitCome = noRangka
                        
71                      frmOrder.noPolisiUnitCome = ""
72                      frmOrder.noPolisiUnitCome = Trim(text(0))
                        
73                      frmOrder.tanggalInUnitCome = ""
74                      frmOrder.tanggalInUnitCome = tanggalInput
                        
75                      frmOrder.Show
76                      frmOrder.WindowState = vbMaximized
77                Else
78                      If gLimitFUDealer - nUnit <= 10 And gLimitFUDealer - nUnit > 0 Then
79                       MsgBox "Hanya bisa menerima " & CDbl(gLimitFUDealer - nUnit) & " Order lagi. " & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
80                      ElseIf gLimitFUDealer <= nUnit Then
81                          MsgBox "Tidak bisa menerima Order lagi. " & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
82                      End If
83                      If nUnit < gLimitFUDealer Then
                               
                                 
84                          frmOrder.noRangkaUnitCome = ""
85                          frmOrder.noRangkaUnitCome = noRangka
                            
86                          frmOrder.noPolisiUnitCome = ""
87                          frmOrder.noPolisiUnitCome = Trim(text(0))
                            
88                          frmOrder.tanggalInUnitCome = ""
89                          frmOrder.tanggalInUnitCome = tanggalInput
                            
90                          frmOrder.Show
91                          frmOrder.WindowState = vbMaximized
92                      ElseIf nUnit > gLimitFUDealer Then
93                          MsgBox "Customer yang belum di-Follow telah mencapai batas " & gLimitFUDealer & " unit" & vbCr & "Lakukan Follow Up Terlebih dahulu !", vbInformation, "Limit Follow Up"
94                      End If
95                End If
                  
96            Else
97                MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
98            End If
99        End If

100      On Error GoTo 0
101      Exit Sub

asxPowerButton2_Click_Error:

102       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure asxPowerButton2_Click of Form FrmOrderListUnitCome")
End Sub

Private Sub asxPowerButton3_Click()
1        On Error GoTo asxPowerButton3_Click_Error

2         If cbFollow <> "" Then
3             On Error GoTo vb_error
              Dim statusFU As Byte
4             statusFU = 2
5             If (cbFollow) = "Terhubung" Then
6                 statusFU = 1
7             End If
8             objO.InsertData ("UPDATE UNITCOME SET statusfu = '" & statusFU & "',tanggalfu = getdate(),userfu = '" & gSA & "' WHERE nopolisi = '" & Trim(text(0)) & "' and convert(VarChar(100), tanggalin,20) = '" & tanggalInput & "' ")
9             MsgBox "Data berhasil disimpan", vbInformation
10            Frame2.Visible = False
11            refreshGrid
12            Exit Sub
vb_error:
13            MsgBox "Terjadi Kesalahan menyimpan data", vbCritical
14            Exit Sub
15        Else
16            MsgBox "Silahkan lengkapi data hasil FU", vbInformation
17        End If

18       On Error GoTo 0
19       Exit Sub

asxPowerButton3_Click_Error:

20        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure asxPowerButton3_Click of Form FrmOrderListUnitCome")
End Sub

Private Sub asxPowerButton4_Click()
    Frame2.Visible = False
End Sub

Private Sub asxPowerButton5_Click()
    Call refreshGrid
End Sub

Private Sub btnLegend_Click()
 If frameLegend.Visible Then
    frameLegend.Visible = False
 Else
    frameLegend.Visible = True
 End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    
    Frame2.Visible = True
    
End Sub

Private Sub Form_Load()
    Tgl1 = Format(isnow, "DD/MM/YYYY")
    Tgl2 = Format(isnow, "DD/MM/YYYY")
    Call refreshGrid
End Sub

Private Function getFilter() As String
On Error GoTo vb_error
   Dim tmp As String
   Dim n As Integer
   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         If col.ExternalEditor = "tgltime" Then
            If Len(col.FilterText) <> 10 Then Exit Function
            tmp = tmp & col.DataField & " > '" & col.FilterText & "' and " & col.DataField & " < '" & CDate(CLng(CDate(col.FilterText)) + 1) & "'"
         ElseIf col.ExternalEditor = "tgl" Then
            If Len(col.FilterText) <> 10 Then Exit Function
            tmp = tmp & col.DataField & " = '" & col.FilterText & "'"
         ElseIf col.ExternalEditor = "Bit" Then
            If Len(col.FilterText) <> 1 Then Exit Function
            tmp = tmp & col.DataField & " = " & col.FilterText & ""
         Else
            tmp = tmp & col.DataField & " LIKE '%" & col.FilterText & "%'"
         End If
      End If
   Next col
   getFilter = tmp
   Exit Function
vb_error:
   Resume Next
End Function

Private Sub unitGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsCustomer.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    
    If Trim(rs.Fields(11).value) = "0" And Trim(rs.Fields(12).value) = "0" Then
        RowStyle = "merah"
    ElseIf Trim(rs.Fields(11).value) = "1" And Trim(rs.Fields(12).value) = "0" Then
        RowStyle = "kuning"
    ElseIf Trim(rs.Fields(12).value) = "1" Then
        RowStyle = "hijau"
    End If
End Sub

Private Sub unitGrid_FilterChange()
On Error GoTo vb_error
   Set cols = unitGrid.Columns
   Dim C As Integer
   C = unitGrid.col
   unitGrid.HoldFields
   rsCustomer.filter = getFilter()
   If rsCustomer.EOF Then
      MsgBox "Data yang dimaksud tidak dapat ditemukan !", vbInformation
      For Each col In unitGrid.Columns
         col.FilterText = ""
      Next col
      refreshGrid
      unitGrid.SetFocus
      unitGrid.col = C
      Exit Sub
   End If
   lblJumlah.Caption = "JUMLAH DATA = " & rsCustomer.RecordCount
   unitGrid.col = C
   unitGrid.EditActive = True
   Exit Sub

vb_error:
    If err.Number = 3001 Then MsgBox "Ada Kesalahan Inputan", vbInformation
    For Each col In unitGrid.Columns
        col.FilterText = ""
    Next col
End Sub

Private Sub unitGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    text(0) = unitGrid.Columns(1).value
    text(3) = unitGrid.Columns(0).value
    text(10) = unitGrid.Columns(3).value
    text(9) = unitGrid.Columns(4).value
    text(1) = unitGrid.Columns(5).value
    text(4) = Format(isnow, "DD/MM/YYYY HH:MM")
    tanggalInput = unitGrid.Columns(0).value 'Format(unitGrid.Columns(1).value, "YYYY-MM-DD HH:MM:SS")
    
End Sub
