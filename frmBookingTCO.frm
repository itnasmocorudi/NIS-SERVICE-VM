VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBookingTCO 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Order TCO"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16245
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   16245
   Begin VB.Frame frameTCO 
      Caption         =   "DETAIL DATA TCO"
      Height          =   9135
      Left            =   48000
      TabIndex        =   29
      Top             =   120
      Width           =   15495
      Begin VB.Frame frameTambahPart 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Tambah Part TCO"
         Height          =   3495
         Left            =   720
         TabIndex        =   117
         Top             =   4320
         Visible         =   0   'False
         Width           =   9735
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmBookingTCO.frx":0000
            Left            =   1680
            List            =   "frmBookingTCO.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   2040
            Width           =   2535
         End
         Begin VB.ComboBox cbSistem 
            Height          =   315
            ItemData        =   "frmBookingTCO.frx":0028
            Left            =   1680
            List            =   "frmBookingTCO.frx":0032
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   240
            Width           =   2535
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   26
            Left            =   1680
            TabIndex        =   120
            Top             =   600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   270
            Left            =   4630
            TabIndex        =   122
            Top             =   600
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   29
            Left            =   1680
            TabIndex        =   124
            Top             =   960
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   30
            Left            =   1680
            TabIndex        =   126
            Top             =   1320
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   31
            Left            =   6360
            TabIndex        =   128
            Top             =   960
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   32
            Left            =   6360
            TabIndex        =   130
            Top             =   1320
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   33
            Left            =   1680
            TabIndex        =   132
            Top             =   1680
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   315
            Index           =   34
            Left            =   6360
            TabIndex        =   133
            Top             =   1680
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
            Height          =   315
            Index           =   35
            Left            =   6360
            TabIndex        =   137
            Top             =   2040
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
            Left            =   1680
            TabIndex        =   140
            Top             =   2400
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
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
            Mask            =   "99"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerButton btnTutupPart 
            Height          =   495
            Left            =   3960
            TabIndex        =   141
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BackColor       =   33023
            Caption         =   "TUTUP"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   16777215
         End
         Begin AIFCmp1.asxPowerButton btnSimpanPart 
            Height          =   495
            Left            =   1680
            TabIndex        =   142
            Top             =   2760
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            BackColor       =   49152
            Caption         =   "SIMPAN"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   16777215
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "QTY "
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
            Index           =   44
            Left            =   480
            TabIndex        =   139
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Est Tersedia"
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
            Index           =   43
            Left            =   4560
            TabIndex        =   138
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Available"
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
            Index           =   42
            Left            =   480
            TabIndex        =   136
            Top             =   2040
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Proses"
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
            Index           =   41
            Left            =   4560
            TabIndex        =   134
            Top             =   1680
            Width           =   1785
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Item Status"
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
            Index           =   40
            Left            =   120
            TabIndex        =   131
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Est Harga"
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
            Index           =   39
            Left            =   4800
            TabIndex        =   129
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TCO Name"
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
            Index           =   38
            Left            =   4800
            TabIndex        =   127
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TCO Number"
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
            Index           =   37
            Left            =   120
            TabIndex        =   125
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Package Number"
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
            Index           =   36
            Left            =   120
            TabIndex        =   123
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Item"
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
            Index           =   34
            Left            =   480
            TabIndex        =   121
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sistem"
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
            Index           =   35
            Left            =   480
            TabIndex        =   119
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "TCO ORDER"
         Height          =   4335
         Left            =   120
         TabIndex        =   87
         Top             =   3960
         Width           =   15135
         Begin TrueOleDBGrid70.TDBGrid gridPartTCO 
            Height          =   2055
            Left            =   240
            TabIndex        =   114
            Top             =   2160
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   3625
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "PackageID"
            Columns(0).DataField=   "PackageID"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "TCONumber"
            Columns(1).DataField=   "TCONumber"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   " TCOName"
            Columns(2).DataField=   "TCOName"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "KodeItem"
            Columns(3).DataField=   "KodeItem"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Qty"
            Columns(4).DataField=   "Qty"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Harga"
            Columns(5).DataField=   "Harga"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Sistem"
            Columns(6).DataField=   "Sistem"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Available"
            Columns(7).DataField=   "Available"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "EstTersedia"
            Columns(8).DataField=   "EstTersedia"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "StatusItem"
            Columns(9).DataField=   "StatusItem"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "TglProses"
            Columns(10).DataField=   "TglProses"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   953
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   15790320
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
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
            DeadAreaBackColor=   15790320
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
         End
         Begin ucXPButton.XPButton btnTambah 
            Height          =   375
            Left            =   1560
            TabIndex        =   113
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "Add Item"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmBookingTCO.frx":0043
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame Frame4 
            Caption         =   "ALAMAT PENGIRIMAN"
            Height          =   1335
            Left            =   4560
            TabIndex        =   97
            Top             =   120
            Width           =   10455
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Index           =   1
               Left            =   3960
               MaxLength       =   1000
               MultiLine       =   -1  'True
               TabIndex        =   110
               Top             =   120
               Width           =   6375
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Index           =   0
               Left            =   1320
               MaxLength       =   6
               TabIndex        =   98
               Top             =   240
               Width           =   1350
            End
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   270
               Index           =   0
               Left            =   2640
               TabIndex        =   99
               Top             =   240
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   22
               Left            =   1320
               TabIndex        =   102
               Top             =   600
               Width           =   2340
               _ExtentX        =   4128
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   23
               Left            =   1320
               TabIndex        =   104
               Top             =   960
               Width           =   2340
               _ExtentX        =   4128
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   24
               Left            =   4800
               TabIndex        =   106
               Top             =   960
               Width           =   2340
               _ExtentX        =   4128
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   25
               Left            =   8040
               TabIndex        =   108
               Top             =   960
               Width           =   2340
               _ExtentX        =   4128
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin VB.Label Label11 
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
               Index           =   33
               Left            =   3000
               TabIndex        =   112
               Top             =   285
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
               Index           =   0
               Left            =   3900
               TabIndex        =   111
               Top             =   330
               Width           =   60
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kelurahan"
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
               Index           =   32
               Left            =   6840
               TabIndex        =   109
               Top             =   1005
               Width           =   1185
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kecamatan"
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
               Index           =   31
               Left            =   3600
               TabIndex        =   107
               Top             =   1005
               Width           =   1185
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kota / Kab"
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
               Index           =   30
               Left            =   120
               TabIndex        =   105
               Top             =   1005
               Width           =   1185
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Provinsi"
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
               Index           =   29
               Left            =   120
               TabIndex        =   103
               Top             =   645
               Width           =   1185
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Pos"
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
               Index           =   28
               Left            =   360
               TabIndex        =   101
               Top             =   285
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
               Index           =   18
               Left            =   1260
               TabIndex        =   100
               Top             =   330
               Width           =   60
            End
         End
         Begin VB.ComboBox cbOrderStatus 
            Height          =   315
            ItemData        =   "frmBookingTCO.frx":005F
            Left            =   1560
            List            =   "frmBookingTCO.frx":006F
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1320
            Width           =   2535
         End
         Begin VB.ComboBox cbMetode 
            Height          =   315
            ItemData        =   "frmBookingTCO.frx":00A1
            Left            =   1560
            List            =   "frmBookingTCO.frx":00AE
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   960
            Width           =   2535
         End
         Begin VB.ComboBox cbKategoriTCO 
            Height          =   315
            ItemData        =   "frmBookingTCO.frx":00CD
            Left            =   1560
            List            =   "frmBookingTCO.frx":00DA
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   600
            Width           =   2535
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   315
            Index           =   28
            Left            =   1920
            TabIndex        =   89
            Top             =   240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Order Status"
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
            Index           =   27
            Left            =   360
            TabIndex        =   95
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Metode"
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
            Index           =   26
            Left            =   360
            TabIndex        =   93
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori TCO"
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
            Index           =   25
            Left            =   360
            TabIndex        =   91
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label Label78 
            BackStyle       =   0  'Transparent
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
            Left            =   3315
            TabIndex        =   90
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Pesan"
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
            Index           =   24
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   1785
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "KENDARAAN"
         Height          =   3135
         Left            =   7800
         TabIndex        =   37
         Top             =   720
         Width           =   7500
         Begin VB.ComboBox cbVehicleSource 
            Height          =   315
            ItemData        =   "frmBookingTCO.frx":0101
            Left            =   1200
            List            =   "frmBookingTCO.frx":0111
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   1680
            Width           =   2535
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   1200
            TabIndex        =   67
            Top             =   240
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   5040
            TabIndex        =   69
            Top             =   240
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   15
            Left            =   1200
            TabIndex        =   71
            Top             =   600
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   16
            Left            =   5040
            TabIndex        =   73
            Top             =   600
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   17
            Left            =   1200
            TabIndex        =   75
            Top             =   960
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   18
            Left            =   5040
            TabIndex        =   77
            Top             =   960
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   19
            Left            =   1200
            TabIndex        =   79
            Top             =   1320
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   20
            Left            =   1200
            TabIndex        =   83
            Top             =   2040
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   21
            Left            =   5040
            TabIndex        =   85
            Top             =   2040
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No SO"
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
            Index           =   23
            Left            =   3840
            TabIndex        =   86
            Top             =   2085
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No SPK"
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
            Index           =   22
            Left            =   0
            TabIndex        =   84
            Top             =   2085
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Source"
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
            Index           =   21
            Left            =   0
            TabIndex        =   81
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori Unit"
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
            Index           =   20
            Left            =   0
            TabIndex        =   80
            Top             =   1365
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Index           =   19
            Left            =   3840
            TabIndex        =   78
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Katashiki"
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
            Index           =   18
            Left            =   0
            TabIndex        =   76
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Variant"
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
            Index           =   17
            Left            =   3840
            TabIndex        =   74
            Top             =   645
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   72
            Top             =   645
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "VIN"
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
            Index           =   15
            Left            =   3840
            TabIndex        =   70
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No Polisi"
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
            Index           =   14
            Left            =   0
            TabIndex        =   68
            Top             =   285
            Width           =   1185
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "PELANGGAN"
         Height          =   3135
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   7620
         Begin VB.CheckBox ckContact 
            Caption         =   "PHONE_CALL"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   66
            Top             =   2760
            Width           =   2175
         End
         Begin VB.CheckBox ckContact 
            Caption         =   "EMAIL"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   65
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CheckBox ckContact 
            Caption         =   "WHATSAPP_OR_SMS"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   64
            Top             =   2280
            Width           =   2175
         End
         Begin VB.CheckBox ckContact 
            Caption         =   "MTOYOTA"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   62
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   1200
            TabIndex        =   58
            Top             =   1680
            Width           =   2415
         End
         Begin VB.CommandButton cmdPlgn 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Pelanggan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   255
            Width           =   1050
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   39
            Top             =   270
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox14 
            Height          =   270
            Left            =   2790
            TabIndex        =   40
            Top             =   240
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   5160
            TabIndex        =   42
            Top             =   240
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   44
            Top             =   600
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   6
            Left            =   5160
            TabIndex        =   46
            Top             =   600
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   1200
            TabIndex        =   48
            Top             =   960
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   5160
            TabIndex        =   50
            Top             =   960
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   1200
            TabIndex        =   52
            Top             =   1320
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   5160
            TabIndex        =   54
            Top             =   1320
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   315
            Index           =   11
            Left            =   5880
            TabIndex        =   59
            Top             =   1680
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            Height          =   315
            Index           =   12
            Left            =   5880
            TabIndex        =   61
            Top             =   2040
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99:99"
            PromptChar      =   ">"
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Preferred"
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
            Index           =   13
            Left            =   0
            TabIndex        =   63
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Time End"
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
            Index           =   11
            Left            =   3960
            TabIndex        =   60
            Top             =   2040
            Width           =   1785
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Time Start"
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
            Index           =   10
            Left            =   3960
            TabIndex        =   57
            Top             =   1680
            Width           =   1785
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
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
            Index           =   9
            Left            =   0
            TabIndex        =   56
            Top             =   1725
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori Cust"
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
            Index           =   7
            Left            =   3960
            TabIndex        =   55
            Top             =   1365
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Telp"
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
            Index           =   6
            Left            =   0
            TabIndex        =   53
            Top             =   1365
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Belakang"
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
            Index           =   5
            Left            =   3960
            TabIndex        =   51
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Depan"
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
            Index           =   4
            Left            =   0
            TabIndex        =   49
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
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
            Index           =   3
            Left            =   3960
            TabIndex        =   47
            Top             =   645
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "NIK"
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
            Index           =   2
            Left            =   0
            TabIndex        =   45
            Top             =   645
            Width           =   1185
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "One Account ID"
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
            Index           =   12
            Left            =   3960
            TabIndex        =   43
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lgn"
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
            TabIndex        =   41
            Top             =   285
            Width           =   1065
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   30
         Top             =   240
         Width           =   3540
         _ExtentX        =   6244
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   6000
         TabIndex        =   32
         Top             =   240
         Width           =   2820
         _ExtentX        =   4974
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   10200
         TabIndex        =   34
         Top             =   240
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin AIFCmp1.asxPowerButton asxPowerButton2 
         Height          =   495
         Left            =   2400
         TabIndex        =   115
         Top             =   8400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "TUTUP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin AIFCmp1.asxPowerButton btnSimpanPartTCO 
         Height          =   495
         Left            =   120
         TabIndex        =   116
         Top             =   8400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BackColor       =   49152
         Caption         =   "SIMPAN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Booking ID"
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
         Left            =   9120
         TabIndex        =   35
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Source Order"
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
         Left            =   4920
         TabIndex        =   33
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order ID"
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
         Index           =   8
         Left            =   360
         TabIndex        =   31
         Top             =   255
         Width           =   945
      End
   End
   Begin VB.ComboBox cbType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmBookingTCO.frx":0151
      Left            =   5880
      List            =   "frmBookingTCO.frx":0164
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   120
      Width           =   3255
   End
   Begin TDBDate6Ctl.TDBDate tglAwal 
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Top             =   120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calendar        =   "frmBookingTCO.frx":01A1
      Caption         =   "frmBookingTCO.frx":02CD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBookingTCO.frx":0339
      Keys            =   "frmBookingTCO.frx":0357
      Spin            =   "frmBookingTCO.frx":03B5
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
      Text            =   "07/08/2024"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   45511
      CenturyMode     =   0
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JADI BOOKING"
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
      Left            =   13800
      TabIndex        =   22
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BATAL BOOKING"
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
      Left            =   11880
      TabIndex        =   21
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BELUM PROSES"
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
      Left            =   10080
      TabIndex        =   20
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame frameLegend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LEGEND WARNA"
      Height          =   1455
      Left            =   480
      TabIndex        =   16
      Top             =   6360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "JADI BOOKING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BATAL BOOKING"
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
         TabIndex        =   18
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3375
      End
   End
   Begin ucXPButton.XPButton btnLegend 
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   7800
      Visible         =   0   'False
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
      MICON           =   "frmBookingTCO.frx":03DD
      PICN            =   "frmBookingTCO.frx":03F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frameBatalBooking 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6000
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox cbCustomerMenolak 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   600
         Width           =   3135
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3960
         Top             =   1080
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
      Begin AIFCmp1.asxPowerButton btnSimpanBatal 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "SIMPAN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin AIFCmp1.asxPowerButton btnBatalPembatalan 
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "BATAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "ALASAN CANCEL TCO"
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
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3375
      End
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   15960
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid gridTagih 
      Height          =   6495
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   11456
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
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
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14933984
      RowDividerColor =   14933984
      RowSubDividerColor=   14933984
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=96,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=93,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=94,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=95,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=100,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=97,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=98,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=99,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=104,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=101,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=102,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=103,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=108,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=105,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=106,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=107,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=112,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=109,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=110,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=111,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=116,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=113,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=114,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=115,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=120,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=117,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=118,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=119,.parent=17"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
      _StyleDefs(85)  =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(86)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(87)  =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(88)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(89)  =   "Named:id=87:SudahDitagih"
      _StyleDefs(90)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(91)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(92)  =   ":id=87,.fontname=Tahoma"
      _StyleDefs(93)  =   "Named:id=83:Booking"
      _StyleDefs(94)  =   ":id=83,.parent=87,.fgcolor=&HFF00&"
   End
   Begin AIFCmp1.asxPowerButton tblImport 
      Height          =   495
      Left            =   11520
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "DATA ONLINE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblBatal 
      Height          =   495
      Left            =   12720
      TabIndex        =   7
      Top             =   7920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BackColor       =   33023
      Caption         =   "TUTUP"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblSimpan 
      Height          =   495
      Left            =   10440
      TabIndex        =   8
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "TAMBAH TCO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   7920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "CANCEL TCO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton3 
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "REFRESH "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin TDBDate6Ctl.TDBDate tglAkhir 
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calendar        =   "frmBookingTCO.frx":4914
      Caption         =   "frmBookingTCO.frx":4A40
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmBookingTCO.frx":4AAC
      Keys            =   "frmBookingTCO.frx":4ACA
      Spin            =   "frmBookingTCO.frx":4B28
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
      Text            =   "07/08/2024"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   45511
      CenturyMode     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdProsesTCO 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   8160
      TabIndex        =   28
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BackColor       =   16711680
      Caption         =   "PROSES TCO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S / D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   25
      Top             =   120
      Width           =   615
   End
   Begin VB.Label jmlTagih 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   12720
      TabIndex        =   6
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Yang Dipilih :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   9480
      TabIndex        =   5
      Top             =   120
      Width           =   3060
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR ORDER TCO ::."
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
      Left            =   5895
      TabIndex        =   2
      Top             =   795
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   15255
   End
End
Attribute VB_Name = "frmBookingTCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRecord As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Private jumlahBookingOnline As Integer
Dim nTotal As Integer
Private rsPartTCO As New ADODB.Recordset
Private rsSearch As New ADODB.Recordset
Private isTambahPart As Boolean
Private isTambahTCO As Boolean

Private Sub createTableTCO()
    On Error GoTo vb_error
'  Set rsPartTCO = New ADODB.Recordset
    If rsPartTCO.State Then rsPartTCO.Close
    Set rsPartTCO = Nothing
    With rsPartTCO
        .Fields.Append "PackageID", adVarChar, 100
        .Fields.Append "TCONumber", adVarChar, 100
        .Fields.Append "KodeItem", adVarChar, 100
        .Fields.Append "TCOName", adVarChar, 100
        .Fields.Append "Qty", adInteger
        .Fields.Append "Harga", adDouble
        .Fields.Append "Sistem", adVarChar, 100
        .Fields.Append "Available", adVarChar, 100
        .Fields.Append "EstTersedia", adDate
        .Fields.Append "StatusItem", adVarChar, 100
        .Fields.Append "TglProses", adDate
        .Open
    End With
    Set gridPartTCO.DataSource = rsPartTCO

    Exit Sub
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Sub

Private Sub FetchCombo()
    ''pada menu ini tambahkan fitur untuk print stok kayu bisa all paket maupun paket yang dipilih
    ''usahakan agar dapat 1 lembar a4 bisa dibagi 2 kolom, isian nama paket detail kayu total qty per diameter dan volume
    ''diameter digabung 10-11,12-14,15-18,19-24,25-29,30-60 per masing2 paket, dan total qty batang dan qty volume total serta dibawah ada tanda tangan kepala log, kepala produksi dan kepala bahan jadi
    With cbCustomerMenolak
        .Clear
        .AddItem "CONFLICT"
        .AddItem "FINANCIAL"
        .AddItem "BETTER_OFFER"
        .AddItem "CHANGE_PREFERENCE"
        .AddItem "DOUBLE_ORDER"
        .AddItem "CHANGE_CHOICE"
        .AddItem "OTHERS"
        .ListIndex = 0
    End With
End Sub

Private Sub asxPowerButton1_Click()
    If nTotal > 0 Then
        If nTotal > 1 Then
            MsgBox "Membatalkan booking hanya diperbolehkan 1 data", vbCritical
            Exit Sub
        Else
            If Trim(rsRecord!statusorder) <> "BOOKING_REQUEST" Then
                MsgBox "Status order tidak diperbolehkan melanjutkan proses.", vbCritical
                Exit Sub
            
            End If
            If MsgBox("Apakan yakin membatalkan " & nTotal & " Pesan Booking?", vbYesNo) = vbYes Then
                frameBatalBooking.Visible = True
                    Call FetchCombo
                    btnSimpanBatal.Top = 1800
                    btnBatalPembatalan.Top = 1800
                    frameBatalBooking.Height = 2415
            End If
        End If
    End If
End Sub

Private Sub asxPowerButton2_Click()
    
    frameTCO.Left = 48000
    
    Call hapusFrameTCO
End Sub

Public Sub asxPowerButton3_Click()
    Call hapusForm
    Call refreshGrid
    
End Sub

Private Sub asxPowerButton4_Click()
    
End Sub

Private Sub btnBatalPembatalan_Click()
    'txtAlasan.Text = ""
    frameBatalBooking.Visible = False
    Call refreshGrid
    
End Sub

Private Sub btnLegend_Click()
    If frameLegend.Visible Then
        frameLegend.Visible = False
    Else
        frameLegend.Visible = True
    End If
End Sub

Private Sub btnSimpanBatal_Click()
 Dim keteranganUpdate As String
 Dim evenID As String
 Dim cancelreason As String
 Dim isBegin  As Boolean
  
   On Error GoTo btnSimpanBatal_Click_Error
    
    isBegin = False
    If cbCustomerMenolak.Text = "" Then
        MsgBox "Silahkan pilih alasan customer menolak", vbInformation
        Exit Sub
    End If
 
    If Get_ConnectionProfil_Bengkel99 = False Then
       MsgBox "Koneksi HO terputus , proses tidak dapat dilanjutkan.", vbInformation
       Exit Sub
    End If
    
    If Not rsRecord.EOF Then
           rsRecord.MoveFirst
           Do While Not rsRecord.EOF
               If rsRecord("pilih") = True Then
                   evenID = ""
                   cancelreason = ""
                   objHT.OpenTransactionHO
                   isBegin = True
                   sSQL = "update DIDX..p2_tco_order set tco_order_status='CANCELLED',tco_order_cancellation_reason='" & cbCustomerMenolak.Text & "' where tco_order_id='" & Trim(rsRecord!kodePesanan) & "'"
                   Call objHT.InsertData_Ho(sSQL)
                   sSQL = "insert into DIDX..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry,asal_cabang,id_uniq)"
                   sSQL = sSQL & " values ('BOOKTCO','" & Trim(rsRecord!kodePesanan) & "','" & Trim(rsRecord!kodePesanan) & "','','" & Right(gDb, 2) & "','CANCELLED',0,'',GETDATE(),'" & Right(gDb, 2) & "',NEWID())"
                   Call objHT.InsertData_Ho(sSQL)
                   objHT.CloseTransactionHO
               End If
               
               rsRecord.MoveNext
           Loop
           
           MsgBox "Data Berhasil Diubah", vbInformation
           frameBatalBooking.Visible = False
           Call refreshGrid
       End If

   On Error GoTo 0
   Exit Sub

btnSimpanBatal_Click_Error:
    If isBegin Then
        objHT.CloseTransactionHO_err
    End If
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure btnSimpanBatal_Click of Form frmBookingTCO"
End Sub

Private Sub btnSimpanPart_Click()
    If isTambahPart Then
        rsPartTCO.AddNew
        rsPartTCO!Sistem = cbSistem
        rsPartTCO!KodeItem = Text(26) 'Replace(Mid(Text(26), 3, 15), "-", "")
        rsPartTCO!PackageID = Text(29)
        rsPartTCO!TCONumber = Text(30)
        rsPartTCO!TCOName = Text(31)
        rsPartTCO!Harga = Text(32)
        rsPartTCO!qty = Text(27)
        rsPartTCO!StatusItem = Text(33)
        rsPartTCO!Available = Combo1
        rsPartTCO!TglProses = Text(34)
        rsPartTCO!EstTersedia = Text(35)
        rsPartTCO.Update
        frameTambahPart.Visible = False
        gridPartTCO.Refresh
    Else
        rsPartTCO!Sistem = cbSistem
        rsPartTCO!KodeItem = Text(26) 'Replace(Mid(Text(26), 3, 15), "-", "")
        rsPartTCO!PackageID = Text(29)
        rsPartTCO!TCONumber = Text(30)
        rsPartTCO!TCOName = Text(31)
        rsPartTCO!Harga = Text(32)
        rsPartTCO!qty = Text(27)
        rsPartTCO!StatusItem = Text(33)
        rsPartTCO!Available = Combo1
        rsPartTCO!TglProses = Text(34)
        rsPartTCO!EstTersedia = Text(35)
        frameTambahPart.Visible = False
        gridPartTCO.Refresh
    End If
End Sub

Private Sub btnSimpanPartTCO_Click()
          
1         If Get_ConnectionProfil_Bengkel99 = False Then
2            MsgBox "Koneksi HO terputus , proses tidak dapat dilanjutkan.", vbInformation
3            Exit Sub
4         End If
          
5         sSQL = "select * from DIDX..p2_tco_order where tco_order_ID='" & Trim(Text(0)) & "'"
6         If objO.FetchDataHO_Count(sSQL) > 0 Then
7             sSQL = "update DIDX..p2_tco_order set dealer_customer_ID='" & Trim(Text(3)) & "',spk_number='" & Trim(Text(20)) & "',so_number='" & Trim(Text(21)) & "',booking_ID='" & Trim(Text(2)) & "',tco_order_status='" & cbOrderStatus & "',"
8             sSQL = sSQL & " delivery_address='" & txt(1) & "',province='" & Trim(Text(22)) & "',city='" & Trim(Text(23)) & "',district='" & Trim(Text(24)) & "',subdistrict='" & Trim(Text(25)) & "',postal_code='" & txt(0) & "'"
9             sSQL = sSQL & " where tco_order_id='" & Trim(Text(0)) & "'"
10            Call objO.InsertDataHO(sSQL)
11            rsPartTCO.MoveFirst
12            If cbKategoriTCO = "PACKAGE" Then
13                sSQL = "select * from DIDX..p2_tco_order_items where tco_order_ID='" & Trim(Text(0)) & "' and Package_ID='" & Trim(rsPartTCO!PackageID) & "'"
14                If objO.FetchDataHO_Count(sSQL) > 0 Then
15                    sSQL = "update DIDX..p2_tco_order_package set tco_quantity='" & Trim(rsPartTCO!qty) & "',tco_source='" & Trim(rsPartTCO!Sistem) & "', tco_availability_status='" & Trim(rsPartTCO!Available) & "',tco_item_status='" & Trim(rsPartTCO!StatusItem) & "',tco_installation_datetime='" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "' where tco_order_id='" & Trim(Text(0)) & "' and Package_ID='" & Trim(rsPartTCO!PackageID) & "'"
16                    Call objO.InsertDataHO(sSQL)
17                Else
18                    sSQL = "INSERT INTO DIDX.. p2_tco_order_package"
19                    sSQL = sSQL & " (tco_order_ID,package_ID,package_name,tco_quantity,tco_est_price,tco_installation_est_price,flag_tco_required_down_payment,tco_availability_status,tco_availability_est_datetime,tco_item_status,tco_source,tco_pickup_datetime,tco_delivery_datetime,tco_installation_datetime)"
20                    sSQL = sSQL & " VALUES ('" & Trim(Text(0)) & "','" & Trim(rsPartTCO!PackageID) & "','" & Trim(rsPartTCO!TCOName) & "','" & Trim(rsPartTCO!qty) & "','" & Trim(rsPartTCO!Harga) & "',0.00,'','" & Trim(rsPartTCO!Available) & "','','" & Trim(rsPartTCO!StatusItem) & "','" & Trim(rsPartTCO!Sistem) & "','','','" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "')"
21                    Call objO.InsertDataHO(sSQL)
22                    sSQL = " INSERT INTO DIDX..p2_tco_order_package_parts"
23                    sSQL = sSQL & " (tco_order_ID,package_ID,tco_number,tco_name,kodeitem)"
24                    sSQL = sSQL & " VALUES ('" & Trim(Text(0)) & "','" & Trim(rsPartTCO!PackageID) & "','" & Trim(rsPartTCO!TCONumber) & "','" & Trim(rsPartTCO!TCOName) & "','" & Trim(rsPartTCO!KodeItem) & "')"
25                    Call objO.InsertDataHO(sSQL)
26                End If
27            Else
28                Do While Not rsPartTCO.EOF
29                    sSQL = "select * from DIDX..p2_tco_order_items where tco_order_ID='" & Trim(Text(0)) & "' and tco_number='" & Trim(rsPartTCO!TCONumber) & "'"
30                    If objO.FetchDataHO_Count(sSQL) > 0 Then
31                        sSQL = "update DIDX.. p2_tco_order_items set kodeitem='" & Trim(rsPartTCO!KodeItem) & "',"
32                        sSQL = sSQL & " tco_quantity='" & Trim(rsPartTCO!qty) & "',tco_source='" & Trim(rsPartTCO!Sistem) & "',"
33                        sSQL = sSQL & " tco_availability_status='" & Trim(rsPartTCO!Available) & "',tco_availability_est_datetime='" & Format(Trim(rsPartTCO!EstTersedia), "yyyy-mm-dd") & "',"
34                        sSQL = sSQL & " tco_item_status='" & Trim(rsPartTCO!StatusItem) & "',tco_delivery_datetime='" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "',tco_pickup_datetime='" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "'"
35                        sSQL = sSQL & " where tco_order_id='" & Trim(Text(0)) & "' and tco_number='" & Trim(rsPartTCO!TCONumber) & "'"
36                        Call objO.InsertDataHO(sSQL)
37                    Else
38                        sSQL = "INSERT INTO DIDX..p2_tco_order_items"
39                        sSQL = sSQL & " (tco_order_ID,kodeitem,tco_number,tco_name,tco_quantity,tco_est_price"
40                        sSQL = sSQL & " ,tco_installation_est_price,flag_tco_required_down_payment,tco_availability_status,tco_availability_est_datetime,tco_item_status,tco_source,tco_pickup_datetime,tco_delivery_datetime,tco_installation_datetime)"
41                        sSQL = sSQL & " VALUES"
42                        sSQL = sSQL & " ('" & Trim(Text(0)) & "','" & Trim(rsPartTCO!KodeItem) & "','" & Trim(rsPartTCO!TCONumber) & "','" & Trim(rsPartTCO!TCOName) & "','" & Trim(rsPartTCO!qty) & "','" & Trim(rsPartTCO!Harga) & "',0.00,'','" & Trim(rsPartTCO!Available) & "',"
43                        sSQL = sSQL & " '','" & Trim(rsPartTCO!StatusItem) & "','" & Trim(rsPartTCO!Sistem) & "','" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "','" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "','')"
44                        Call objO.InsertDataHO(sSQL)
45                    End If
46                    rsPartTCO.MoveNext
47                Loop
                  
48            End If
49        Else
50            sSQL = "INSERT INTO DIDX..p2_tco_order"
51            sSQL = sSQL & " (tco_order_ID,tco_order_number,event_ID,timestamp,one_account_ID,dealer_customer_ID,first_name,last_name,phone_number,email,preferred_contact_time_start,preferred_contact_time_end,customer_category,ktp_number,npwp_number,vehicle_source,vin,police_number,katashiki_suffix,color_code,model,variant,color,vehicle_category,outlet_ID,outlet_name,created_datetime,spk_number,so_number,order_source,additional_notes,tco_item_type,tco_order_status,tco_order_cancellation_reason,tco_order_preference,booking_ID,booking_number,address_label,delivery_address,province,city,district,subdistrict,postal_code,detail_address,tgl_entry)"
52            sSQL = sSQL & " VALUES ('" & Trim(Text(0)) & "',NEWID(),NEWID(),GETDATE(),'" & Trim(Text(4)) & "','" & Trim(Text(3)) & "','" & Trim(Text(7)) & "','" & Trim(Text(8)) & "','" & Trim(Text(9)) & "','" & Trim(txtEmail.Text) & "','" & Trim(Text(11)) & "','" & Trim(Text(12)) & "','" & Trim(Text(10)) & "','" & Trim(Text(5)) & "','" & Trim(Text(6)) & "','" & cbVehicleSource & "','" & Trim(Text(14)) & "','" & Trim(Text(13)) & "','" & Trim(Text(17)) & "','" & Trim(Text(18)) & "','" & Trim(Text(15)) & "','" & Trim(Text(16)) & "','" & Trim(Text(18)) & "','" & Trim(Text(19)) & "','" & Log.kdDealerDidx & "','" & gNmCbg & "',GETDATE(),'" & Trim(Text(20)) & "','" & Trim(Text(21)) & "','" & Trim(Text(1)) & "','','" & cbKategoriTCO & "','" & cbOrderStatus & "','','" & cbMetode & "','" & Trim(Text(2)) & "','','','" & Trim(txt(1)) & "','" & Trim(Text(22)) & "','" & Trim(Text(23)) & "','" & Trim(Text(24)) & "','" & Trim(Text(25)) & "','" & Trim(txt(0)) & "','',GETDATE())"
53            Call objO.InsertDataHO(sSQL)
54            If ckContact(0).value = vbChecked Then
55                sSQL = "INSERT INTO didx..p2_tco_order_pref_contact_channel (tco_order_ID,preferred_contact_channel) VALUES ('" & Trim(Text(0)) & "','MTOYOTA')"
56                Call objO.InsertDataHO(sSQL)
57            End If
58            If ckContact(1).value = vbChecked Then
59                sSQL = "INSERT INTO didx..p2_tco_order_pref_contact_channel (tco_order_ID,preferred_contact_channel) VALUES ('" & Trim(Text(0)) & "','WHATSAPP_OR_SMS')"
60                Call objO.InsertDataHO(sSQL)
61            End If
62            If ckContact(2).value = vbChecked Then
63                sSQL = "INSERT INTO didx..p2_tco_order_pref_contact_channel (tco_order_ID,preferred_contact_channel) VALUES ('" & Trim(Text(0)) & "','EMAIL')"
64                Call objO.InsertDataHO(sSQL)
65            End If
66            If ckContact(3).value = vbChecked Then
67                sSQL = "INSERT INTO didx..p2_tco_order_pref_contact_channel (tco_order_ID,preferred_contact_channel) VALUES ('" & Trim(Text(0)) & "','PHONE_CALL')"
68                Call objO.InsertDataHO(sSQL)
69            End If
70            rsPartTCO.MoveFirst
71            If cbKategoriTCO = "PACKAGE" Then
                  
72                sSQL = "INSERT INTO DIDX.. p2_tco_order_package"
73                sSQL = sSQL & " (tco_order_ID,package_ID,package_name,tco_quantity,tco_est_price,tco_installation_est_price,flag_tco_required_down_payment,tco_availability_status,tco_availability_est_datetime,tco_item_status,tco_source,tco_pickup_datetime,tco_delivery_datetime,tco_installation_datetime)"
74                sSQL = sSQL & " VALUES ('" & Trim(Text(0)) & "','" & Trim(rsPartTCO!PackageID) & "','" & Trim(rsPartTCO!TCOName) & "','" & Trim(rsPartTCO!qty) & "','" & Trim(rsPartTCO!Harga) & "',0.00,'','" & Trim(rsPartTCO!Available) & "','','" & Trim(rsPartTCO!StatusItem) & "','" & Trim(rsPartTCO!Sistem) & "','','','" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "')"
75                Call objO.InsertDataHO(sSQL)
76                sSQL = " INSERT INTO DIDX..p2_tco_order_package_parts"
77                sSQL = sSQL & " (tco_order_ID,package_ID,tco_number,tco_name,kodeitem)"
78                sSQL = sSQL & " VALUES ('" & Trim(Text(0)) & "','" & Trim(rsPartTCO!PackageID) & "','" & Trim(rsPartTCO!TCONumber) & "','" & Trim(rsPartTCO!TCOName) & "','" & Trim(rsPartTCO!KodeItem) & "')"
79                Call objO.InsertDataHO(sSQL)
80            Else
                  
81                Do While Not rsPartTCO.EOF
82                    sSQL = "INSERT INTO DIDX..p2_tco_order_items"
83                    sSQL = sSQL & " (tco_order_ID,kodeitem,tco_number,tco_name,tco_quantity,tco_est_price"
84                    sSQL = sSQL & " ,tco_installation_est_price,flag_tco_required_down_payment,tco_availability_status,tco_availability_est_datetime,tco_item_status,tco_source,tco_pickup_datetime,tco_delivery_datetime,tco_installation_datetime)"
85                    sSQL = sSQL & " VALUES"
86                    sSQL = sSQL & " ('" & Trim(Text(0)) & "','" & Trim(rsPartTCO!KodeItem) & "','" & Trim(rsPartTCO!TCONumber) & "','" & Trim(rsPartTCO!TCOName) & "','" & Trim(rsPartTCO!qty) & "','" & Trim(rsPartTCO!Harga) & "',0.00,'','" & Trim(rsPartTCO!Available) & "',"
87                    sSQL = sSQL & " '" & Format(Trim(rsPartTCO!EstTersedia), "yyyy-mm-dd") & "','" & Trim(rsPartTCO!StatusItem) & "','" & Trim(rsPartTCO!Sistem) & "','" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "','" & Format(Trim(rsPartTCO!TglProses), "yyyy-mm-dd") & "','')"
88                    Call objO.InsertDataHO(sSQL)
89                    rsPartTCO.MoveNext
90                Loop
91            End If
92        End If
          'sSQL = "insert into DIDX..kirimDataDidx"
93        sSQL = "insert into " & dbReplikasiHO & "..kirimDataDidx"
94        sSQL = sSQL & " (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry,asal_cabang,id_uniq)"
95        sSQL = sSQL & " values ('BOOKTCO','" & Trim(Text(0)) & "','" & Trim(Text(0)) & "','','" & Right(gDb, 2) & "','" & cbOrderStatus & "',0,'',GETDATE(),'" & Right(gDb, 2) & "',NEWID())"
96        Call objO.InsertData(sSQL) ''Call objO.InsertDataHO(sSQL)
97        MsgBox "Data berhasil disimpan", vbInformation
          
98        frameTCO.Left = 48000
          
99        Call hapusFrameTCO
          
End Sub

Private Sub btnTambah_Click()

    Dim itemStatus As String
    isTambahPart = True
    Text(26) = ""
    Text(27) = 0
    cbSistem.ListIndex = 0
    Combo1.ListIndex = 0
    Text(34) = Format(isnow, "dd-mm-yyyy") ' "01-01-1900"
    Text(35) = "01-01-1900"
    If cbMetode = "PICKUP" Then
        If cbOrderStatus = "CONFIRMED" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "READY" Then
            itemStatus = "READY_TAKEN"
        ElseIf cbOrderStatus = "COMPLETED" Then
            itemStatus = "COMPLETED"
        End If
    ElseIf cbMetode = NulSID("Delivery") Then
        If cbOrderStatus = "CONFIRMED" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "READY" Then
            itemStatus = "READY_DELIVER"
        ElseIf cbOrderStatus = "COMPLETED" Then
            itemStatus = "COMPLETED"
        End If
    
    ElseIf cbMetode = NulSID("Install") Then
        If cbOrderStatus = "CONFIRMED" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "READY" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "COMPLETED" Then
            itemStatus = "PROCESSED"
        End If
    End If
    Text(33) = itemStatus
    frameTambahPart.Visible = True
End Sub

Private Sub btnTutupPart_Click()
    frameTambahPart.Visible = False
End Sub

Private Sub ambilDataPelanggan(idcust As String)
    Dim rsSingleId As New ADODB.Recordset
    sSQL = "select NamaCust1,NamaCust2,NoKTP,NoNPWP,HP1,Email,TypeCustomer,* from SingleIDOri..NewCustomer where IDCust= '" & idcust & "'"
    Set rsSingleId = objO.FetchData(sSQL)
    If Not rsSingleId.EOF Then
        Text(5) = Trim(rsSingleId!noKTP)
        Text(6) = Trim(rsSingleId!noNPWP)
        Text(7) = Trim(rsSingleId!namacust1)
        Text(8) = Trim(rsSingleId!namacust2)
        Text(9) = Trim(rsSingleId!hp1)
        If UCase(Nul(rsSingleId!TypeCustomer)) = "PERORANGAN" Then
            Text(10) = NulSID("INDIVIDUAL")
        Else
            Text(10) = NulSID("COMPANY")
        End If
        
        Text(4) = objO.FetchData_String("select one_account_ID from ReplBengkel99..p2_OneAccess where phone_number='" & Nul(Text(9)) & "'")
        txtEmail.Text = Trim(rsSingleId!email)
    End If
End Sub

Private Sub ambilDataKendaraan(noRangka As String)
    Dim rsSingleId As New ADODB.Recordset
    sSQL = "select Fmdl as model,Fmdl as variant,Mdl as katashiki,Color,NoPolisi from SingleIDOri..unitcustomer where NoRangka='" & noRangka & "'"
    Set rsSingleId = objO.FetchData(sSQL)
    If Not rsSingleId.EOF Then
        Text(13) = Trim(rsSingleId!NoPolisi)
        Text(15) = Trim(rsSingleId!model)
        Text(16) = Trim(rsSingleId!Variant)
        Text(17) = Trim(rsSingleId!katashiki)
        Text(18) = Trim(rsSingleId!Color)
        
        Text(19) = objO.FetchData_String("select vehicle_category from ReplBengkel99..p2_ToyotaID where vin = '" & Nul(noRangka) & "'")
        If NulSID(Text(19)) = "" Then
            Text(19) = "RETAIL"
        End If
    End If

End Sub

Private Sub cmdPlgn_Click()
    
    Call panggilCari
    If gMainKeyCari <> "" Then
        Text(3) = gMainKeyCari
        Text(14) = gBantuan
        Call ambilDataPelanggan(Trim(Text(3)))
        
        Call ambilDataKendaraan(Trim(Text(14)))
        gMainKeyCari = ""
    End If
End Sub

Private Sub panggilCari()
    gMainKeyCari = ""
    gBantuan = ""
    With frmSIDSearchCust
        .Text(1) = gMainKeyCari
        .Text(4) = ""
        gSearchOption = "frmSIDCustomer"
        .Left = frameTCO.Left '+ 100
        .Top = Me.Top + 450
        .show vbModal
    End With
End Sub

Private Sub cmdProsesTCO_Click()
    
    If nTotal > 0 Then
        If nTotal > 1 Then
            MsgBox "Membatalkan booking hanya diperbolehkan 1 data", vbCritical
            Exit Sub
        Else
            If Trim(rsRecord!statusorder) <> "BOOKING_REQUEST" Then
                MsgBox "Status order tidak diperbolehkan melanjutkan proses.", vbCritical
                Exit Sub
            
            End If
            If MsgBox("Apakan yakin memproses data " & rsRecord!kodePesanan & " ?", vbYesNo) = vbYes Then
                isTambahTCO = False
                Call prosesTCO(rsRecord!kodePesanan)
                
            End If
        End If
    End If
End Sub

Private Sub hapusFrameTCO()
    For i = 0 To 28
        Select Case i
            Case 27
                Text(i) = 0
            Case 28
                Text(i) = ">>->>->>>>"
                Text(i).Mask = "99-99-9999"
            Case 11, 12
                Text(i) = ">>:>>"
                Text(i).Mask = "99:99"
            Case Else
                Text(i) = ""
        End Select
    Next
    txt(0) = ""
    txt(1) = ""
    ckContact(0).value = vbUnchecked
    ckContact(1).value = vbUnchecked
    ckContact(2).value = vbUnchecked
    ckContact(3).value = vbUnchecked
End Sub


Private Sub prosesTCO(kodePesanan As String)
    frameTCO.Left = 480
    Call hapusFrameTCO
    Dim rsTemp As New ADODB.Recordset
    Text(0) = kodePesanan
    Text(1) = "DEALER_SYSTEM"
    cbOrderStatus = "CONFIRMED"
    cbOrderStatus.Enabled = False
        Text(28) = Format(isnow, "dd-mm-yyyy") '
        Call createTableTCO
    sSQL = "select tco_order_ID as [OrderID],order_source as [SourceOrder],booking_id,"
    sSQL = sSQL & " dealer_customer_ID as [IDCust],one_account_ID as [AccountID],ktp_number as NIK,"
    sSQL = sSQL & " npwp_number as NPWP,first_name as [NamaDepan],last_name as [NamaBelakang],"
    sSQL = sSQL & " phone_number as Telp,customer_category as [KategoriCust],email,preferred_contact_time_start,preferred_contact_time_end,"
    sSQL = sSQL & " police_number as [NoPolisi],vin,model,variant,katashiki_suffix,color as Warna,"
    sSQL = sSQL & " vehicle_category as [KategoriUnit],vehicle_source,spk_number as [NoSPK],so_number as [SONumber],"
    sSQL = sSQL & " created_datetime as [TglPesan],tco_item_type as [KategoriTCO],tco_order_preference as Metode,tco_order_status as [OrderStatus],"
    sSQL = sSQL & " postal_code As KodePos, province As provinsi, city As Kota, district As Kecamatan, subdistrict As Kelurahan, delivery_address As Alamat"
    sSQL = sSQL & " from DIDX..p2_tco_order"
    sSQL = sSQL & " where tco_order_ID='" & Trim(kodePesanan) & "' and outlet_ID = '" & Log.kdDealerDidx & "'"
    Set rsTemp = objO.FetchDataHO(sSQL)
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Text(0) = Trim(rsTemp!OrderID)
        Text(1) = Trim(rsTemp!SourceOrder)
        Text(2) = Trim(rsTemp!booking_ID)
        Text(3) = Trim(rsTemp!idcust)
        Text(4) = Trim(rsTemp!AccountID)
        Text(5) = Trim(rsTemp!nik)
        Text(6) = Trim(rsTemp!npwp)
        Text(7) = Trim(rsTemp!NamaDepan)
        Text(8) = Trim(rsTemp!NamaBelakang)
        Text(9) = Trim(rsTemp!telp)
        Text(10) = Trim(rsTemp!KategoriCust)
        txtEmail.Text = Trim(rsTemp!email)
        Text(11) = Trim(rsTemp!preferred_contact_time_start)
        Text(12) = Trim(rsTemp!preferred_contact_time_end)
        
        Text(13) = Trim(rsTemp!NoPolisi)
        Text(14) = Trim(rsTemp!vin)
        Text(15) = Trim(rsTemp!model)
        Text(16) = Trim(rsTemp!Variant)
        Text(17) = Trim(rsTemp!katashiki_suffix)
        Text(18) = Trim(rsTemp!warna)
        Text(19) = Trim(rsTemp!KategoriUnit)
        
        If Trim(rsTemp!vehicle_source) <> "" Then
            cbVehicleSource = Trim(rsTemp!vehicle_source)
        End If
        Text(20) = Trim(rsTemp!NoSPK)
        Text(21) = Trim(rsTemp!SONumber)
        Text(28) = Format(Trim(rsTemp!TglPesan), "dd-mm-yyyy")
        
        If Trim(rsTemp!KategoriTCO) <> "" Then
            cbKategoriTCO = Trim(rsTemp!KategoriTCO)
        End If
        
        If Trim(rsTemp!Metode) <> "" Then
            cbMetode = Trim(rsTemp!Metode)
        End If
        
'        If Trim(rsTemp!OrderStatus) <> "" Then
'            cbOrderStatus = Trim(rsTemp!OrderStatus)
'        End If
'        cbOrderStatus = "CONFIRMED"
'        cbOrderStatus.Enabled = False
        
        txt(0) = Trim(rsTemp!KodePos)
        txt(1) = Trim(rsTemp!Alamat)
        Text(22) = Trim(rsTemp!provinsi)
        Text(23) = Trim(rsTemp!Kota)
        Text(24) = Trim(rsTemp!kecamatan)
        Text(25) = Trim(rsTemp!kelurahan)
        
        
        If Trim(rsTemp!KategoriTCO) <> "" Then
            If Trim(rsTemp!KategoriTCO) = "PACKAGE" Then
                sSQL = "select package_id PackageID, '' as [TCONumber],package_name as [TCOName],tco_quantity as Qty,tco_est_price as Harga,tco_source as Sistem,"
                sSQL = sSQL & " tco_availability_status As Available, tco_availability_est_datetime As EstTersedia, tco_item_status As StatusItem, tco_installation_datetime As TglProses"
                sSQL = sSQL & " from DIDX..p2_tco_order_package"
                sSQL = sSQL & " where tco_order_id='" & Trim(kodePesanan) & "'"
            Else
                sSQL = "SELECT '' AS PackageID,b.tco_number as [TCONumber],b.tco_name as [TCOName],b.tco_quantity AS [Qty],b.tco_est_price AS [Harga],"
                sSQL = sSQL & " b.tco_source as Sistem,b.tco_availability_status AS Available,b.tco_availability_est_datetime AS [EstTersedia],b.tco_item_status AS [StatusItem],"
                sSQL = sSQL & " case when a.tco_item_type='PICKUP' then b.tco_pickup_datetime"
                sSQL = sSQL & " when a.tco_item_type='DELIVERY' then b.tco_delivery_datetime"
                sSQL = sSQL & " Else B.tco_installation_datetime"
                sSQL = sSQL & " end As TglProses"
                sSQL = sSQL & " FROM DIDX..p2_tco_order a"
                sSQL = sSQL & " LEFT JOIN DIDX..p2_tco_order_items b ON b.tco_order_ID = a.tco_order_ID"
                sSQL = sSQL & " where a.tco_order_id='" & Trim(kodePesanan) & "'"
            End If
            Set rsTemp = objO.FetchDataHO(sSQL)
            If Not rsTemp.EOF Then
                rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                    With rsPartTCO
                        .AddNew
                        !PackageID = Trim(rsTemp!PackageID)
                        !TCONumber = Trim(rsTemp!TCONumber)
                        !TCOName = Trim(rsTemp!TCOName)
                        !qty = Trim(rsTemp!qty)
                        !Harga = Trim(rsTemp!Harga)
                        !Sistem = Trim(rsTemp!Sistem)
                        !Available = Trim(rsTemp!Available)
                        !EstTersedia = Nul(Trim(rsTemp!EstTersedia), "1900-01-01")
                        !StatusItem = Trim(rsTemp!StatusItem)
                        !TglProses = Nul(Trim(rsTemp!TglProses), "1900-01-01")
                        .Update
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        End If
        
    End If
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Load()
    tglAwal = "01/" & Format(isnow, "mm/yyyy")
    tglAkhir = DateAdd("D", -1, DateAdd("M", 1, tglAwal))
    
    Call hapusForm
    Call refreshGrid
    
End Sub

Private Sub hapusForm()
    nTagih = 0
    totTagih = 0
End Sub

Private Sub createRS()
    If rsRecord.State Then rsRecord.Close
    Set rsRecord = Nothing
    With rsRecord
        'If .State <> adStateClosed Then .Close
        .Fields.Append "pilih", adBoolean
        .Fields.Append "NamaPelanggan", adVarChar, 250
        .Fields.Append "KodePesanan", adVarChar, 150
        .Fields.Append "TanggalPesanan", adVarChar, 50
        .Fields.Append "Kategori", adVarChar, 100
        .Fields.Append "Metode", adVarChar, 100
        .Fields.Append "StatusOrder", adVarChar, 100
         
        .Open
    End With
    Set gridTagih.DataSource = rsRecord
    Call headColumns
End Sub

Private Sub headColumns()
    
    With gridTagih
        .Columns(0).Caption = "Check"
        .Columns(0).DataField = "pilih"
        .Columns(0).Width = 600
        .Splits(0).Columns(0).Locked = False
        
        .Columns(1).Caption = "Nama Pelanggan"
        .Columns(1).DataField = "NamaPelanggan"
        .Columns(1).Width = 2500
      '  .Splits(0).Columns(1).Locked = True
        
        .Columns(3).Caption = "Tanggal Pesanan"
        .Columns(3).DataField = "TanggalPesanan"
        .Columns(3).Width = 2500
        .Splits(0).Columns(3).Locked = True
        
        .Columns(2).Caption = "Kode Pesanan"
        .Columns(2).DataField = "KodePesanan"
        .Columns(2).Width = 3000
        .Splits(0).Columns(2).Locked = True
        
        
        .Columns(4).Caption = "Kategori"
        .Columns(4).DataField = "Kategori"
        .Columns(4).Width = 2000
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Caption = "Metode"
        .Columns(5).DataField = "Metode"
        .Columns(5).Width = 2000
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Caption = "Status Order"
        .Columns(6).DataField = "StatusOrder"
        .Columns(6).Width = 2500
        .Splits(0).Columns(6).Locked = True
         
    End With
End Sub

Private Sub refreshGrid()
 Dim rsTagih As ADODB.Recordset
Dim varTagih As String
Dim kriteria As String
kriteria = "("
        
    sSQL = "select a.tco_order_number,a.first_name + ' ' + a.last_name as Pelanggan,a.phone_number as Telp,a.police_number as [NoPolisi],"
    sSQL = sSQL & " a.model as Unit,a.vin as [NoRangka],a.tco_order_number as [OrderNumber],a.created_datetime as [TanggalPesan],"
    sSQL = sSQL & " b.tco_item_type as [Kategori],b.tco_order_preference as Metode,b.tco_order_status as [STATUS],a.created_datetime"
    sSQL = sSQL & " from DIDX..p2_tco_order a"
    sSQL = sSQL & " left join DIDX..p2_tco_order_orders b on b.tco_order_ID=a.tco_order_number"
    sSQL = sSQL & " where a.created_datetime between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "' and b.tco_order_status='" & cbType & "'"
    
    sSQL = "select first_name + ' ' + last_name as [NamaPelanggan],tco_order_ID as [KodePesanan],created_datetime as [TanggalPesanan],tco_item_type as Kategori,tco_order_preference as Metode,tco_order_status as [StatusOrder] "
    sSQL = sSQL & " from DIDX..p2_tco_order"
    sSQL = sSQL & " where outlet_ID = '" & Log.kdDealerDidx & "' aND created_datetime between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "'"
    If cbType <> "" Then
        sSQL = sSQL & "  and tco_order_status ='" & cbType & "'"
        
    End If

    If Get_ConnectionProfil_Bengkel99 = False Then
        MsgBox "Ada Kesalahan Setting ke HO," & Chr(13) & "Cek VPN dan konektivitas ke HO" & Chr(13) & "(Lakukan Routing IP)", vbCritical + vbOKOnly
        Exit Sub
    End If
    Set rsTagih = objO.FetchDataHO(sSQL)
    
    If rsTagih.RecordCount > 0 Then
        Call createRS
        Dim hasilProsesBooking As String
        i = 1
        nTotal = 0
        rsTagih.MoveFirst
        Do While Not rsTagih.EOF
            With rsRecord
                .AddNew
                !Pilih = False
                !NamaPelanggan = NulSID(rsTagih("NamaPelanggan"))
                !kodePesanan = Trim(rsTagih("KodePesanan"))
                !TanggalPesanan = Format(Trim(rsTagih("TanggalPesanan")), "dd-mm-yyyy hh:mm")
                !kategori = Nul(rsTagih("Kategori"))
                !Metode = Nul(rsTagih("Metode"))
                !statusorder = Nul(rsTagih("StatusOrder"))
                .Update
            End With
            i = i + 1
            rsTagih.MoveNext
        Loop
        gridTagih.DataSource = rsRecord
        jmlTagih.Caption = nTotal
        gridTagih.FetchRowStyle = True
        Call headColumns
        rsRecord.MoveFirst
    Else
        Call createRS
    End If
    
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

Private Sub gridPartTCO_DblClick()
    Text(26) = ""
    Text(27) = 0
    cbSistem.ListIndex = 0
    Combo1.ListIndex = 0
    Text(34) = ">>->>->>>>"
    Text(35) = ">>->>->>>>"
    Text(29) = rsPartTCO!PackageID
    Text(30) = rsPartTCO!TCONumber
    Text(31) = rsPartTCO!TCOName
    Text(32) = rsPartTCO!Harga
    Dim itemStatus As String
    isTambahPart = False
    If cbMetode = "PICKUP" Then
        If cbOrderStatus = "CONFIRMED" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "READY" Then
            itemStatus = "READY_TAKEN"
        ElseIf cbOrderStatus = "COMPLETED" Then
            itemStatus = "COMPLETED"
        End If
    ElseIf cbMetode = NulSID("Delivery") Then
        If cbOrderStatus = "CONFIRMED" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "READY" Then
            itemStatus = "READY_DELIVER"
        ElseIf cbOrderStatus = "COMPLETED" Then
            itemStatus = "COMPLETED"
        End If
    
    ElseIf cbMetode = NulSID("Install") Then
        If cbOrderStatus = "CONFIRMED" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "READY" Then
            itemStatus = "PROCESSED"
        ElseIf cbOrderStatus = "COMPLETED" Then
            itemStatus = "PROCESSED"
        End If
    End If
    Text(33) = itemStatus
    Text(27) = rsPartTCO!qty
    frameTambahPart.Visible = True
End Sub

Private Sub gridTagih_AfterColUpdate(ByVal ColIndex As Integer)
    With gridTagih
        Select Case ColIndex
            Case 0
                If .Columns(ColIndex) Then
                    nTotal = nTotal + 1
                Else
                    nTotal = nTotal - 1
                End If
               
        End Select
        
        jmlTagih.Caption = nTotal
    End With
End Sub

Private Sub gridTagih_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
'    Set rs = rsRecord.Clone
'    If Bookmark > 1 Then
'        rs.Bookmark = Bookmark
'    End If
'
'    If Trim(rs.Fields(11).value) = "Batal Booking" Then
'        RowStyle = "SudahDitagih"
'    ElseIf Trim(rs.Fields(11).value) = "Booking" Then
'        RowStyle = "Booking"
'    End If
End Sub

Private Sub gridTagih_FilterChange()
 On Error GoTo errHandler
    Set cols = gridTagih.Columns
    Dim C As Integer
    
    C = gridTagih.col
    gridTagih.HoldFields
    rsRecord.filter = getFilter()
    If rsRecord.EOF Then
        For Each col In gridTagih.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""
    End If
    gridTagih.col = C
    gridTagih.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridTagih_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 12 Then rsRecord.Sort = gridTagih.Columns(ColIndex).DataField & " " & Urut
4        gridTagih.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub tblBatal_Click()
    Unload Me
End Sub

Private Sub tblImport_Click()
    Dim f As New frmBookingImportDataOnline '.Show vbModal
    f.show vbModal
End Sub

Private Sub tblSimpan_Click()
isTambahTCO = True

Call prosesTCO(Replace(Replace(objO.FetchData_String("SELECT NEWID()"), "{", ""), "}", ""))
                
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
            Case 26
                gMainKeyCari = Text(Index)
                gSearchOption = "PART"
                frmCariAcess.show (vbModal)
        End Select
    End If
    
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    
        Select Case Index
            Case 26
                If Text(Index) <> "" Then
                    Set rsSearch = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
                    If rsSearch.BOF Then
                    Else
                        If isTambahTCO Then
                            Text(30) = Replace(Mid(NulSID(rsSearch!KodeItem), 3, 15), "-", "")
                            Text(31) = NulSID(rsSearch!NamaBarang)
                        End If
                        Text(32) = Val(rsSearch("HARGAJUALA"))
                    End If
                
            End If
                
        End Select
    
End Sub


Private Sub txt_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If gMainKeyCari <> "" Then
        txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        DiTab = True
        Sendkeys ("{TAB}")
        'txt(Index).SetFocus
        'SendKeys ("{TAB}")
    End If
    
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                
                If gMainKeyCari = "" Then gMainKeyCari = txt(Index)
                gSearchOption = "TOPOLOGY"
                If frmCari.Visible Then
                    frmCari.Hide
                End If
                frmCari.show (vbModal)
             
        
        End Select
        onFind = False
    ElseIf KeyCode = 13 Then
        If Index = 0 Then
            If NulSID(txt(Index)) <> "" Then
                Dim rsTopology As ADODB.Recordset
                Set rsTopology = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE KODEPOS = '" & NulSID(txt(0)) & "'")
                If Not rsTopology.EOF Then
                    rsTopology.MoveFirst
                    Text(22) = NulSID(rsTopology!Propinsi)
                    Text(23) = NulSID(rsTopology!Kota) 'nulsid(DGrid.Columns(1))
                    Text(24) = NulSID(rsTopology!kecamatan) 'nulsid(DGrid.Columns(2))
                    Text(25) = NulSID(rsTopology!kelurahan) 'nulsid(DGrid.Columns(3))
                    txt(0) = NulSID(rsTopology!KodePos)
                Else
                    MsgBox "Data kodepos tidak terdaftar.", vbInformation
                    Text(22) = ""
                    Text(23) = "" 'nulsid(DGrid.Columns(1))
                    Text(24) = "" 'nulsid(DGrid.Columns(2))
                    Text(25) = "" 'nulsid(DGrid.Columns(3))
                    txt(0) = "" 'nulsid(DGrid.Columns(4))
                End If
            End If
        Else
            'Call txt_Validate(Index, False)
        End If
    End If
End Sub

