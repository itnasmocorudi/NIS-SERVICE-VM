VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmSIDCustomer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Customer"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10110
      Left            =   450
      TabIndex        =   83
      Top             =   45
      Width           =   14025
      Begin VB.Frame FrAlm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   990
         TabIndex        =   14
         Top             =   1335
         Visible         =   0   'False
         Width           =   2715
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
            Index           =   63
            Left            =   0
            MaxLength       =   100
            TabIndex        =   15
            Top             =   15
            Width           =   2715
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
            Index           =   64
            Left            =   0
            MaxLength       =   100
            TabIndex        =   16
            Top             =   345
            Width           =   2715
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
            Index           =   65
            Left            =   0
            MaxLength       =   100
            TabIndex        =   17
            Top             =   675
            Width           =   2715
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
            Index           =   66
            Left            =   0
            MaxLength       =   100
            TabIndex        =   18
            Top             =   1005
            Width           =   2715
         End
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo1 
         Height          =   275
         Left            =   8445
         TabIndex        =   176
         Top             =   4635
         Visible         =   0   'False
         Width           =   1320
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   2328
         _ExtentY        =   485
         _StockProps     =   93
         Text            =   "SSOleDBCombo1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.ComboBox cboProv 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1365
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
         Index           =   11
         Left            =   9135
         MaxLength       =   10
         TabIndex        =   165
         Top             =   6930
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Index           =   10
         Left            =   5685
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2400
         Width           =   570
      End
      Begin VB.ComboBox cboKel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2340
         Width           =   2700
      End
      Begin VB.ComboBox cboKec 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2025
         Width           =   2700
      End
      Begin VB.ComboBox cboKota 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1710
         Width           =   2700
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   4845
         MaxLength       =   3
         TabIndex        =   20
         Top             =   2400
         Width           =   570
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   4845
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1365
         Width           =   2130
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11295
         Picture         =   "frmSIDCustomer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   8520
         Width           =   2010
      End
      Begin VB.CommandButton cmdHapus 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9300
         Picture         =   "frmSIDCustomer.frx":06CF
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   8520
         Width           =   2010
      End
      Begin VB.CommandButton cmdSimpan 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7305
         Picture         =   "frmSIDCustomer.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   8520
         Width           =   2010
      End
      Begin VB.CommandButton cmdPelanggan 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4095
         Picture         =   "frmSIDCustomer.frx":1252
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   345
         Width           =   2320
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4055
         Left            =   7890
         TabIndex        =   112
         Top             =   18000
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CheckBox chAsuransi 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5580
            TabIndex        =   160
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdPlgAsuransi 
            BackColor       =   &H00DCFCFB&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Picture         =   "frmSIDCustomer.frx":1B33
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   3480
            Width           =   2320
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   38
            Left            =   1050
            TabIndex        =   113
            Top             =   1170
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   39
            Left            =   1050
            TabIndex        =   114
            Top             =   1470
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Index           =   37
            Left            =   1050
            TabIndex        =   69
            Top             =   840
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   40
            Left            =   1050
            TabIndex        =   70
            Top             =   2475
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   41
            Left            =   3915
            TabIndex        =   71
            Top             =   2475
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   42
            Left            =   1050
            TabIndex        =   72
            Top             =   2790
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   43
            Left            =   1050
            TabIndex        =   73
            Top             =   3120
            Visible         =   0   'False
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Index           =   52
            Left            =   1050
            TabIndex        =   68
            Top             =   480
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   53
            Left            =   1050
            TabIndex        =   151
            Top             =   1800
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   54
            Left            =   1050
            TabIndex        =   152
            Top             =   2100
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   55
            Left            =   2865
            TabIndex        =   153
            Top             =   2100
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   270
            Left            =   2640
            TabIndex        =   161
            Top             =   480
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   0
            X2              =   6000
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   0
            Picture         =   "frmSIDCustomer.frx":238A
            Top             =   0
            Width           =   6000
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PSTN No."
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
            TabIndex        =   156
            Top             =   1845
            Width           =   675
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hp.  No."
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
            Left            =   315
            TabIndex        =   155
            Top             =   2145
            Width           =   600
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ";"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2755
            TabIndex        =   154
            Top             =   2145
            Width           =   75
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   270
            TabIndex        =   150
            Top             =   525
            Width           =   660
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
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
            TabIndex        =   121
            Top             =   3165
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
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
            Left            =   3120
            TabIndex        =   120
            Top             =   2520
            Width           =   225
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Limit"
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
            Left            =   600
            TabIndex        =   119
            Top             =   2835
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
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
            TabIndex        =   118
            Top             =   2520
            Width           =   570
         End
         Begin VB.Label Label35 
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
            Left            =   525
            TabIndex        =   117
            Top             =   885
            Width           =   405
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   600
            TabIndex        =   116
            Top             =   1515
            Width           =   330
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   435
            TabIndex        =   115
            Top             =   1215
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "&H000000FF&"
         Height          =   4995
         Left            =   45
         TabIndex        =   106
         Top             =   3915
         Width           =   7095
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
            Left            =   4530
            MaxLength       =   50
            TabIndex        =   202
            Top             =   3975
            Width           =   2415
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
            Index           =   12
            Left            =   960
            MaxLength       =   50
            TabIndex        =   200
            Top             =   3975
            Width           =   3000
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   191
            Top             =   3660
            Width           =   3000
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
            Left            =   4530
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   3975
            Visible         =   0   'False
            Width           =   2205
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
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   193
            Top             =   3975
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
            Left            =   945
            Style           =   2  'Dropdown List
            TabIndex        =   192
            Top             =   3660
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1935
            Left            =   45
            TabIndex        =   177
            Top             =   465
            Width           =   6975
            Begin VB.Frame frmAlamat2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1350
               Left            =   915
               TabIndex        =   27
               Top             =   600
               Visible         =   0   'False
               Width           =   2715
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
                  Index           =   3
                  Left            =   0
                  MaxLength       =   100
                  TabIndex        =   34
                  Top             =   1005
                  Width           =   2715
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
                  Left            =   0
                  MaxLength       =   100
                  TabIndex        =   33
                  Top             =   675
                  Width           =   2715
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
                  Index           =   1
                  Left            =   0
                  MaxLength       =   100
                  TabIndex        =   32
                  Top             =   345
                  Width           =   2715
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
                  Index           =   0
                  Left            =   0
                  MaxLength       =   100
                  TabIndex        =   31
                  Top             =   15
                  Width           =   2715
               End
            End
            Begin VB.ComboBox cboKotaS 
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
               Height          =   315
               Left            =   930
               Style           =   2  'Dropdown List
               TabIndex        =   190
               Top             =   975
               Width           =   2700
            End
            Begin VB.ComboBox cboKecS 
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
               Height          =   315
               Left            =   930
               Style           =   2  'Dropdown List
               TabIndex        =   189
               Top             =   1290
               Width           =   2700
            End
            Begin VB.ComboBox cboKelS 
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
               Height          =   315
               Left            =   930
               Style           =   2  'Dropdown List
               TabIndex        =   188
               Top             =   1605
               Width           =   2700
            End
            Begin VB.ComboBox cboProvS 
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
               Height          =   315
               Left            =   930
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   630
               Width           =   2700
            End
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
               Height          =   960
               IMEMode         =   3  'DISABLE
               Index           =   6
               Left            =   4770
               MaxLength       =   150
               MultiLine       =   -1  'True
               TabIndex        =   35
               Top             =   630
               Width           =   2130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               IMEMode         =   3  'DISABLE
               Index           =   5
               Left            =   4770
               MaxLength       =   3
               TabIndex        =   36
               Top             =   1620
               Width           =   570
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   4
               Left            =   5610
               MaxLength       =   3
               TabIndex        =   37
               Top             =   1620
               Width           =   570
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   270
               Index           =   62
               Left            =   915
               TabIndex        =   28
               Top             =   300
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8454143
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   31
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   270
               Index           =   63
               Left            =   2910
               TabIndex        =   29
               Top             =   300
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8454143
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   31
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   270
               Index           =   64
               Left            =   4770
               TabIndex        =   30
               Top             =   300
               Width           =   2130
               _ExtentX        =   3757
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8454143
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   31
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin ucXPButton.XPButton cmdBackS 
               Height          =   345
               Left            =   6465
               TabIndex        =   178
               Top             =   -90
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
               MICON           =   "frmSIDCustomer.frx":30C8
               PICN            =   "frmSIDCustomer.frx":30E4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSMask.MaskEdBox MaskEdBox10 
               Height          =   270
               Left            =   2625
               TabIndex        =   198
               Top             =   0
               Width           =   60
               _ExtentX        =   106
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   66
               Left            =   915
               TabIndex        =   82
               Top             =   0
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
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
            Begin MSMask.MaskEdBox MaskEdBox13 
               Height          =   270
               Left            =   3660
               TabIndex        =   208
               Top             =   645
               Width           =   60
               _ExtentX        =   106
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox14 
               Height          =   270
               Left            =   3660
               TabIndex        =   209
               Top             =   1620
               Width           =   60
               _ExtentX        =   106
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox15 
               Height          =   270
               Left            =   3660
               TabIndex        =   210
               Top             =   1305
               Width           =   60
               _ExtentX        =   106
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox16 
               Height          =   270
               Left            =   3660
               TabIndex        =   211
               Top             =   975
               Width           =   60
               _ExtentX        =   106
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " No."
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
               Index           =   5
               Left            =   4410
               TabIndex        =   212
               Top             =   1275
               Width           =   300
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   135
               TabIndex        =   199
               Top             =   45
               Width           =   660
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   " Kompleks"
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
               Index           =   9
               Left            =   4035
               TabIndex        =   197
               Top             =   1050
               Width           =   705
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jalan /"
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
               Index           =   8
               Left            =   4215
               TabIndex        =   196
               Top             =   795
               Width           =   480
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
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
               Left            =   390
               TabIndex        =   185
               Top             =   375
               Width           =   405
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RT/RW"
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
               Index           =   6
               Left            =   4005
               TabIndex        =   184
               Top             =   1620
               Width           =   510
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "           /"
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
               Left            =   4740
               TabIndex        =   183
               Top             =   1620
               Width           =   780
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   4
               Left            =   75
               TabIndex        =   182
               Top             =   1665
               Width           =   720
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   181
               Top             =   1335
               Width           =   795
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kab/Kota"
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
               Index           =   2
               Left            =   135
               TabIndex        =   180
               Top             =   1035
               Width           =   660
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   179
               Top             =   720
               Width           =   555
            End
         End
         Begin VB.CheckBox chLog 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Simpan ke Log"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5145
            TabIndex        =   53
            Top             =   4620
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin VB.CheckBox chSda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "sda"
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
            Height          =   255
            Left            =   6135
            TabIndex        =   26
            Top             =   40
            Width           =   735
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   31
            Left            =   930
            TabIndex        =   39
            Top             =   810
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   32
            Left            =   930
            TabIndex        =   42
            Top             =   1710
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Height          =   270
            Index           =   34
            Left            =   960
            TabIndex        =   44
            Top             =   2730
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   33
            Left            =   960
            TabIndex        =   43
            Top             =   2415
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   30
            Left            =   930
            TabIndex        =   38
            Top             =   480
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   35
            Left            =   945
            TabIndex        =   45
            Top             =   3015
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8454143
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   36
            Left            =   3435
            TabIndex        =   46
            Top             =   3030
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   4515
            TabIndex        =   50
            Top             =   3675
            Visible         =   0   'False
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">CCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   270
            Left            =   3990
            TabIndex        =   139
            Top             =   3660
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   960
            TabIndex        =   49
            Top             =   3705
            Visible         =   0   'False
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">CCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   270
            Left            =   2865
            TabIndex        =   140
            Top             =   3705
            Visible         =   0   'False
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   960
            TabIndex        =   48
            Top             =   3375
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   48
            Left            =   960
            TabIndex        =   51
            Top             =   4305
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Index           =   49
            Left            =   960
            TabIndex        =   52
            Top             =   4605
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   57
            Left            =   930
            TabIndex        =   40
            Top             =   1100
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   58
            Left            =   930
            TabIndex        =   41
            Top             =   1400
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   59
            Left            =   960
            TabIndex        =   47
            Top             =   3360
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   270
            Left            =   3990
            TabIndex        =   201
            Top             =   3975
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox12 
            Height          =   270
            Left            =   6660
            TabIndex        =   203
            Top             =   3975
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   225
            X2              =   7305
            Y1              =   4920
            Y2              =   4920
         End
         Begin VB.Image Image2 
            Height          =   465
            Left            =   60
            Picture         =   "frmSIDCustomer.frx":32EF
            Top             =   0
            Width           =   7065
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Height          =   195
            Left            =   480
            TabIndex        =   159
            Top             =   3390
            Width           =   360
         End
         Begin VB.Label Label37 
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
            Left            =   525
            TabIndex        =   158
            Top             =   3075
            Width           =   315
         End
         Begin VB.Label Label28 
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
            Left            =   2865
            TabIndex        =   157
            Top             =   3075
            Width           =   435
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tempat / Tgl Lahir"
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
            Left            =   150
            TabIndex        =   145
            Top             =   4365
            Width           =   690
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agama"
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
            TabIndex        =   144
            Top             =   4050
            Width           =   495
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hobi"
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
            Left            =   4155
            TabIndex        =   143
            Top             =   4005
            Width           =   315
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pekerjaan"
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
            Left            =   105
            TabIndex        =   142
            Top             =   3720
            Width           =   720
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   3570
            TabIndex        =   141
            Top             =   3375
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
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
            Left            =   180
            TabIndex        =   111
            Top             =   2460
            Width           =   660
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   315
            TabIndex        =   110
            Top             =   855
            Width           =   495
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   480
            TabIndex        =   109
            Top             =   1755
            Width           =   330
         End
         Begin VB.Label Label25 
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
            Left            =   465
            TabIndex        =   108
            Top             =   2775
            Width           =   375
         End
         Begin VB.Label Label15 
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
            Left            =   405
            TabIndex        =   107
            Top             =   510
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Tipe Pelanggan "
         Height          =   855
         Left            =   5820
         TabIndex        =   104
         Top             =   2730
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton OpIn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Internal"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton OpExt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "External"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   990
         TabIndex        =   6
         Top             =   1065
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   5
         Left            =   9150
         TabIndex        =   5
         Top             =   5625
         Visible         =   0   'False
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   6
         Left            =   9150
         TabIndex        =   79
         Top             =   6285
         Visible         =   0   'False
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Index           =   11
         Left            =   990
         TabIndex        =   80
         Top             =   3615
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   15
         Left            =   8430
         TabIndex        =   54
         Top             =   435
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   44
         Left            =   1020
         TabIndex        =   75
         Top             =   8925
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   28
         Left            =   12285
         TabIndex        =   81
         Top             =   3705
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   11880
         TabIndex        =   67
         Top             =   1065
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   18
         Left            =   8430
         TabIndex        =   55
         Top             =   1065
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   20
         Left            =   8430
         TabIndex        =   57
         Top             =   1860
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   8430
         TabIndex        =   58
         Top             =   2175
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   8430
         TabIndex        =   59
         Top             =   2775
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   23
         Left            =   8430
         TabIndex        =   85
         Top             =   2475
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   26
         Left            =   8430
         TabIndex        =   64
         Top             =   3705
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   27
         Left            =   8430
         TabIndex        =   65
         Top             =   4035
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   22
         Left            =   12150
         TabIndex        =   86
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   2565
         TabIndex        =   2
         Top             =   435
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   11520
         TabIndex        =   87
         Top             =   2175
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   8
         Left            =   990
         TabIndex        =   23
         Top             =   3000
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   8430
         TabIndex        =   98
         Top             =   765
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   11520
         TabIndex        =   99
         Top             =   435
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   29
         Left            =   12285
         TabIndex        =   100
         Top             =   4035
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   3500
         TabIndex        =   78
         Top             =   765
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2700
         TabIndex        =   102
         Top             =   765
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   9
         Left            =   990
         TabIndex        =   24
         Top             =   3300
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   10
         Left            =   3495
         TabIndex        =   25
         Top             =   3300
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   990
         TabIndex        =   22
         Top             =   2685
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   19
         Left            =   8430
         TabIndex        =   56
         Top             =   1365
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   435
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   45
         Left            =   8430
         TabIndex        =   60
         Top             =   3090
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   3900
         TabIndex        =   123
         Top             =   435
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   25
         Left            =   8430
         TabIndex        =   61
         Top             =   3390
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   46
         Left            =   9045
         TabIndex        =   62
         Top             =   3390
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
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
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   47
         Left            =   9870
         TabIndex        =   63
         Top             =   3390
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8454143
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   11520
         TabIndex        =   147
         Top             =   2775
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   1600
         TabIndex        =   148
         Top             =   435
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   990
         TabIndex        =   3
         Top             =   765
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   50
         Left            =   990
         TabIndex        =   0
         Top             =   435
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   51
         Left            =   855
         TabIndex        =   66
         Top             =   4890
         Visible         =   0   'False
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   270
         Index           =   56
         Left            =   9150
         TabIndex        =   9
         Top             =   5940
         Visible         =   0   'False
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   60
         Left            =   3000
         TabIndex        =   7
         Top             =   1065
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   61
         Left            =   4845
         TabIndex        =   8
         Top             =   1065
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdBack 
         Height          =   345
         Left            =   3720
         TabIndex        =   166
         Top             =   1365
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
         MICON           =   "frmSIDCustomer.frx":748E
         PICN            =   "frmSIDCustomer.frx":74AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   65
         Left            =   4110
         TabIndex        =   186
         Top             =   780
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9630
         TabIndex        =   207
         Top             =   1920
         Width           =   120
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11610
         TabIndex        =   206
         Top             =   2835
         Width           =   120
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11610
         TabIndex        =   205
         Top             =   2235
         Width           =   120
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10395
         TabIndex        =   204
         Top             =   3450
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Kompleks; No."
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
         Index           =   7
         Left            =   3765
         TabIndex        =   195
         Top             =   1860
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   13
         Left            =   315
         TabIndex        =   175
         Top             =   1455
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kab/Kota"
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
         Index           =   14
         Left            =   210
         TabIndex        =   174
         Top             =   1770
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   18
         Left            =   75
         TabIndex        =   173
         Top             =   2070
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   19
         Left            =   150
         TabIndex        =   172
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jalan /"
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
         Index           =   21
         Left            =   4305
         TabIndex        =   171
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "           /"
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
         Left            =   4815
         TabIndex        =   170
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT/RW"
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
         Index           =   23
         Left            =   4080
         TabIndex        =   169
         Top             =   2400
         Width           =   510
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
         Left            =   9030
         TabIndex        =   168
         Top             =   6975
         Visible         =   0   'False
         Width           =   60
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
         Index           =   24
         Left            =   7890
         TabIndex        =   167
         Top             =   6975
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   0
         Picture         =   "frmSIDCustomer.frx":76B5
         Top             =   -75
         Width           =   13470
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11640
         TabIndex        =   146
         Top             =   450
         Width           =   120
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9720
         TabIndex        =   138
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10380
         TabIndex        =   137
         Top             =   4080
         Width           =   120
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10395
         TabIndex        =   136
         Top             =   3750
         Width           =   120
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2370
         TabIndex        =   135
         Top             =   2370
         Width           =   120
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(YYYY)"
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
         Left            =   9120
         TabIndex        =   134
         Top             =   1905
         Width           =   480
      End
      Begin VB.Label Label39 
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
         Left            =   2940
         TabIndex        =   133
         Top             =   3345
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   465
         TabIndex        =   132
         Top             =   1140
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   8535
         TabIndex        =   131
         Top             =   5700
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   8700
         TabIndex        =   130
         Top             =   6360
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rangka"
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
         TabIndex        =   129
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   435
         TabIndex        =   128
         Top             =   3660
         Width           =   435
      End
      Begin VB.Label Label4 
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
         Left            =   495
         TabIndex        =   127
         Top             =   3045
         Width           =   375
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   210
         TabIndex        =   126
         Top             =   803
         Width           =   660
      End
      Begin VB.Label Label8 
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
         Left            =   555
         TabIndex        =   125
         Top             =   3345
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   210
         TabIndex        =   124
         Top             =   2730
         Width           =   660
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km Terakhir"
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
         Left            =   7440
         TabIndex        =   122
         Top             =   3135
         Width           =   840
      End
      Begin VB.Label Label24 
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
         Left            =   7710
         TabIndex        =   105
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wilayah"
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
         Left            =   2880
         TabIndex        =   103
         Top             =   810
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
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
         Left            =   11805
         TabIndex        =   101
         Top             =   4080
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   7830
         TabIndex        =   97
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
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
         TabIndex        =   96
         Top             =   9000
         Width           =   585
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
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
         Index           =   0
         Left            =   11865
         TabIndex        =   95
         Top             =   3750
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Index           =   0
         Left            =   11415
         TabIndex        =   94
         Top             =   1110
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   7320
         TabIndex        =   93
         Top             =   1110
         Width           =   975
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
         Left            =   7635
         TabIndex        =   92
         Top             =   1905
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
         Left            =   7860
         TabIndex        =   91
         Top             =   2220
         Width           =   420
      End
      Begin VB.Label Label19 
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
         Left            =   7800
         TabIndex        =   90
         Top             =   2820
         Width           =   480
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Mesin"
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
         Left            =   7605
         TabIndex        =   89
         Top             =   3750
         Width           =   705
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kunci"
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
         Left            =   7635
         TabIndex        =   88
         Top             =   4080
         Width           =   675
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   7650
         TabIndex        =   84
         Top             =   3435
         Width           =   645
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKP"
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
         Left            =   465
         TabIndex        =   149
         Top             =   4935
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9375
      Left            =   0
      Picture         =   "frmSIDCustomer.frx":B837
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmSIDCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Private kliPr   As Boolean
Private kliKo   As Boolean
Private kliKc   As Boolean
Private kliKl   As Boolean
Public isDariSearch    As Boolean
Private isLoadAktif As Boolean
Dim pds As Boolean

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
          Dim aa As String
          Dim xHasil As String
1        On Error GoTo GensSql_Error

2         xHasil = ""
3         If xUpdate Then
4             xHasil = "UPDATE " + xTabel + " SET "
5             i = 0
6             Do While gsSql1(i) <> ""
7                 If i <> 0 Then
8                     xHasil = xHasil + ","
9                 End If
                  
10                aa = UCase(gsSql1(i))
11                Select Case aa
                  Case "PPN", "DISC"
12                    xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
13                Case Else
14                    xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
15                End Select
16                i = i + 1
17            Loop
18        Else
19            xHasil = "INSERT INTO " + xTabel + " ("
20            i = 0
21            Do While gsSql1(i) <> ""
22                If i <> 0 Then
23                    xHasil = xHasil + ","
24                End If
25                xHasil = xHasil + gsSql1(i)
26                i = i + 1
27            Loop
28            xHasil = xHasil + ") VALUES ("
              
29            i = 0
30            Do While gsSql1(i) <> ""
31                If i <> 0 Then
32                    xHasil = xHasil + ","
33                End If
                  
34                aa = UCase(gsSql1(i))
35                Select Case aa
                  Case "PPN", "DISC"
36                    xHasil = xHasil + gsSql2(i)
37                Case Else
38                    xHasil = xHasil + "'" + gsSql2(i) + "'"
39                End Select
40                i = i + 1
41            Loop
42            xHasil = xHasil + ")"
43        End If
44        GensSql = xHasil

45       On Error GoTo 0
46       Exit Function

GensSql_Error:

47        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GensSql of Form frmSIDCustomer")
End Function

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cboKecS_Click()
    If cboKecS.Tag = "" Then
        kliPr = True
        Call cboKecS_Change
        'cboKelS.Clear
        'cboKecS.Clear
    End If
End Sub

Private Sub cboKelS_Click()
    If cboKelS.Tag = "" Then
        kliPr = True
        Call cboKelS_Change
    End If
End Sub

Private Sub cboKotaS_Click()
    If cboKotaS.Tag = "" Then
        kliPr = True
        Call cboKotaS_Change
        cboKelS.Clear
        'cboKecS.Clear
    End If
End Sub

Private Sub cboPek_Change()
    If UCase(cboPek.Text) = "LAINNYA" Then
        txt(7).BackColor = tColor
        txt(7).SetFocus
    Else
        txt(7).BackColor = fColor
        Sendkeys ("{TAB}")
    End If
End Sub

Private Sub cboProv_Change()
    If kliPr = True Then
        If cboProv.Text <> "" Then
            Dim rsX As New ADODB.Recordset, z As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProv.Text)) & "' Order By Kota")
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
                    cboKota.Text = cboKota.Tag
                    cboKota.Tag = ""
                    
                    If cboKota.Enabled Then cboKota.SetFocus
                End If
            End With
        Else
            cboKota.Tag = "": cboKota.Clear
        End If
        If cboProv.Text = "LAINNYA" Then
            hideComboProv (False)
            txt(63).SetFocus
        Else
            hideComboProv (True)
        End If
        kliPr = False
    End If
End Sub

Private Sub cboProv_Click()
    If cboProv.Tag = "" Then
        kliPr = True
        Call cboProv_Change
        cboKel.Clear
        cboKec.Clear
    End If
End Sub

Private Sub cboProvS_Change()
    If cboProvS.Text <> "" Then
        Dim rsX As New ADODB.Recordset, z As Integer
        Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvS.Text)) & "' Order By Kota")
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
                cboKotaS.Text = cboKotaS.Tag
                cboKotaS.Tag = ""
                If cboKotaS.Enabled Then
                    cboKotaS.SetFocus
                End If
            End If
        End With
    Else
        cboKotaS.Tag = "": cboKotaS.Clear
    End If
    If cboProvS.Text = "LAINNYA" Then
        Call hideComboProvS(False)
    Else
        Call hideComboProvS(True)
    End If
End Sub

Private Sub cboProvS_Click()
    If cboProvS.Tag = "" Then
        kliPr = True
        Call cboProvS_Change
        cboKelS.Clear
        cboKecS.Clear
    End If
End Sub

Private Sub chAsuransi_Click()
    If chAsuransi.value = Checked Then
        Text(52).Enabled = True
        Text(37).Enabled = True
        Text(40).Enabled = True
        Text(41).Enabled = True
        Text(42).Enabled = True
        Text(43).Enabled = True
        Text(52).SetFocus
    ElseIf chAsuransi.value = Unchecked Then
        Text(52).Enabled = False
        Text(37).Enabled = False
        Text(40).Enabled = False
        Text(41).Enabled = False
        Text(42).Enabled = False
        Text(43).Enabled = False
        Text(37) = ""
        n = Text(40).Mask
        Text(40).Mask = ""
        Text(40) = ""
        Text(40).Mask = n
        n = Text(41).Mask
        Text(41).Mask = ""
        Text(41) = ""
        Text(41).Mask = n
        Text(42) = "0"
        Text(43) = ""
    End If
End Sub

Private Sub chAsuransi_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chLog_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub disableSuratMenyurat(value As Boolean)
    'txt(66).Enabled = value
    txt(6).Enabled = value
    txt(5).Enabled = value
    txt(4).Enabled = value
    Text(33).Enabled = value
    Text(34).Enabled = value
    Text(35).Enabled = value
    Text(36).Enabled = value
    
    Text(62).Enabled = value
    Text(63).Enabled = value
    Text(64).Enabled = value
    
'    If value = False Then
'        If nulsid(Text(59)) <> "" Then
'            Text(59).Enabled = value
'        End If
'    Else
'        Text(59).Enabled = value
'    End If
    Text(59).Enabled = value
    
    cboPek.Enabled = value
    txt(7).Enabled = value
    cboAgama.Enabled = value
    cboHobi.Enabled = value
    Text(48).Enabled = value
    Text(49).Enabled = value
    
    'If cmdBack.Visible = True Then
        txt(0).Enabled = value
        txt(1).Enabled = value
        txt(2).Enabled = value
        txt(3).Enabled = value
'        cmdBack.Enabled = value
'    Else
'        Call hideComboProvS(True)
'        cboProvS.Enabled = value
'
'        cboKotaS.Enabled = value
'
'        cboKecS.Enabled = value
'
'        cboKelS.Enabled = value
'    End If
    
    
End Sub

Private Sub chSda_Click()
    If chSda.value = Checked Then
        Text(66) = NulSID(Text(2))
        Text(66).Enabled = False
        Call text_Validate(66, False)
'        Text(62) = nulsid(Text(4))
'        Text(63) = nulsid(Text(60))
'        Text(64) = nulsid(Text(61))
'        'If cmdBack.Visible = True Then
'            Call hideComboProvS(False)
'            txt(0) = nulsid(txt(63))
'            txt(1) = nulsid(txt(64))
'            txt(2) = nulsid(txt(65))
'            txt(3) = nulsid(txt(66))
''        Else
''            Call hideComboProvS(True)
''            cboProvS.Text = cboProv.Text
''            Call cboProvS_Change
''
''            cboKotaS.Text = cboKota.Text
''            Call cboKotaS_Change
''
''            cboKecS.Text = cboKec.Text
''            Call cboKecS_Change
''
''            cboKelS.Text = cboKel.Text
''            Call cboKelS_Change
''        End If
'        txt(6) = nulsid(txt(8))
'        txt(5) = nulsid(txt(9))
'        txt(4) = nulsid(txt(10))
'        Text(33) = nulsid(Text(7))
'        Text(34) = nulsid(Text(8))
'        Text(35) = nulsid(Text(9))
'        Text(36) = nulsid(Text(10))
        
        Call disableSuratMenyurat(False)
    ElseIf chSda.value = Unchecked Then
        disableSuratMenyurat (True)
        Text(62) = ""
        Text(63) = ""
        Text(64) = ""
        cmdBackS.Visible = False
        Call hideComboProvS(False)
        
        txt(0) = ""
        txt(1) = ""
        txt(2) = ""
        txt(3) = ""
            
        txt(6) = ""
        txt(5) = ""
        txt(4) = ""
        Text(33) = ""
        Text(34) = ""
        Text(35) = ""
        Text(36) = ""
        Text(66) = ""
        Text(66).Enabled = True
        
        txt(7) = ""
        txt(12) = ""
        txt(13) = ""
    End If
End Sub

Private Sub chSda_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdBack_Click()
    hideComboProv (True)
    Call refreshComboBox
End Sub

Private Sub cmdBackS_Click()
    hideComboProvS (True)
    Call refreshComboBox
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL' AND FS_LOG_DEL = '' AND NO_CHASSIS = '" & Trim(Trim(Text(50)) & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")) & "'")
    If Not rsTemp.EOF Then
        MsgBox "Kendaraan ini tidak boleh dihapus, karena sudah memiliki transaksi. "
    Else
        If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
            sSQL = "delete CUSTDAT where NO_CHASSIS ='" & Trim(Text(50)) & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'"
            Call objO.UpdateData(sSQL)
            Call hapusForm
        End If
    End If
End Sub

Private Sub cmdPelanggan_Click()
    Call tampilkanFormSearch
    Text(0).SetFocus
'    frmSIDMasterPelanggan.Show vbModal, Me
    
'    With frmtPlgn
'        .Text(57) = Text(2)
'        .Text(1) = Text(4)
'        .Text(2) = Left(Text(5), 40)
'        .Text(3) = Text(6)
'        .Text(4) = Text(8) + "," + Trim(Text(9)) + "," + Trim(Text(10))
'        .Text(6) = Trim(Text(11))
'        gBantuan2 = "kosong"
'        frmtPlgn.Show vbModal, Me
'    End With
End Sub

Private Sub cmdPlgAsuransi_Click()
    With frmtPlgn
        .Text(57) = Trim(Text(52))
        .Text(1) = Trim(Text(37))
        .Text(2) = Trim(Text(38))
        .Text(3) = Trim(Text(39))
        .Text(4) = Text(8) + "," + Trim(Text(9)) + "," + Trim(Text(10))
        .Text(6) = Trim(Text(11))
        gBantuan2 = "kosong"
        frmtPlgn.show vbModal, Me
    End With
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim rsSearch As ADODB.Recordset
      Dim dataLengkap As String

1         On Error GoTo vbError
2         If Trim(Text(0)) <> "" Or Trim(Text(1)) <> "" Then
3             boleh = True
              
              ' Cek Field mandatory yg kosong
4             If Trim(Text(50)) = "" And Text(1) = "" And Text(0) = "" Then
5                 boleh = False
6                 pesan = pesan + "No Rangka. "
7             End If
8             If Text(2) = "" Then
9                 boleh = False
10                pesan = pesan + "Kode Pelanggan. "
11            End If
12            If Text(8) = "" And Text(9) = "" And Text(10) = "" Then
13                boleh = False
14                pesan = pesan + "No Telpon. "
15            End If
16            If Text(34) = "" And Text(35) = "" And Text(36) = "" Then
17                boleh = False
18                pesan = pesan + "No Telpon Surat Menyurat. "
19            End If
20            If NulSID(txt(7)) = "" Then
21                boleh = False
22                pesan = pesan + " Pekerjaan . "
23            End If
24            If Text(15) = "" Then
25                boleh = False
26                pesan = pesan + "Dealer Penjual. "
27            Else
28                Set rsSearch = objO.FetchData("SELECT FS_KD_DEALER FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Trim(Text(15)) & "'")
29                If rsSearch.EOF Then
30                    boleh = False
31                    pesan = pesan & vbCr & "Dealer Penjual tidak sesuai master. " & vbCr
32                End If
33            End If
34            If Text(18) = ">>->>->>>>" Then
35                boleh = False
36                pesan = pesan + "Tanggal Delivery. "
37            Else
                  ''2016-05-27 || penambahan validasi delivery date
38                If Not CekTgl(Text(18)) Then
39                   boleh = False
40                   pesan = pesan + "Tanggal Delivery. "
41                ElseIf Trim(Text(18)) = "01/01/1900" Then
42                   boleh = False
43                   pesan = pesan + "Tanggal Delivery tidak boleh 01/01/1900."
44                ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Trim(Text(18)), "dd/mm/yyyy")) > 0 Then
45                   boleh = False
46                   pesan = pesan + "Tanggal Delivery tidak boleh melebihi tanggal hari ini."
47                Else
48                   Text(18) = TglDDMMYYYY(Text(18))
49                End If
50            End If
51            If Text(19) = "" Then
52                boleh = False
53                pesan = pesan + "DOTAM. "
54            End If
55            If Text(20) = "" Then
56                boleh = False
57                pesan = pesan + "Tahun Kendaraan. "
58            ElseIf IsNumeric(Trim(Text(20))) = False Then
59                boleh = False
60                pesan = pesan + "Tahun Kendaraan harus angka. "
61            ElseIf year(isnow) - CInt(Trim(Text(20))) > 50 Then
62                boleh = False
63                pesan = pesan + "Tahun Kendaraan minimal tahun " & year(isnow) - 50 & ". "
64            ElseIf CInt(Trim(Text(20))) - year(isnow) > 0 Then
65                boleh = False
66                pesan = pesan + "Tahun Kendaraan maksimal tahun " & year(isnow) & ". "
67            End If
68            If Text(21) = "" Then
69                boleh = False
70                pesan = pesan + "Model. "
71            Else
                  
72                If Log.RPLHistori = 0 Then
73                    Set rsSearch = objO.FetchData("SELECT MDL,MODEL FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Trim(Text(21)) & "'")
74                Else
75                    Set rsSearch = objO.FetchData("SELECT MDL,MODEL FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & Trim(Text(21)) & "'")
76                End If
77                If rsSearch.EOF Then
78                    boleh = False
79                    pesan = pesan & vbCr & "Model tidak terdapat pada master. " & vbCr
80                Else
81                    If NulSID(Text(22)) <> "" Then
82                        If NulSID(Text(22)) <> NulSID(rsSearch!model) Then
83                            boleh = False
84                            pesan = pesan & vbCr & "Kelas kendaraan tidak sesuai master Model. Enter pada kode model" & vbCr
                          
85                        End If
86                    End If
87                End If
88            End If
89            If Text(22) = "" Then
90                boleh = False
91                pesan = pesan & vbCr & "Ada kesalahan pada master model. " & vbCr
92            Else
93                If Log.RPLHistori = 1 Then
94                    Set rsSearch = objO.FetchData("SELECT CLASSCODE FROM [" & dbReplikasiHO & "]..CARCLASS WITH(NOLOCK) WHERE CLASSCODE = '" & Trim(Text(22)) & "'")
95                Else
96                    Set rsSearch = objO.FetchData("SELECT CLASSCODE FROM CARCLASS WITH(NOLOCK) WHERE CLASSCODE = '" & Trim(Text(22)) & "'")
97                End If
                  ''Set rsSearch = objO.FetchData("SELECT CLASSCODE FROM CARCLASS WITH(NOLOCK) WHERE CLASSCODE = '" & Trim(text(22)) & "'")
98                If rsSearch.EOF Then
99                    boleh = False
100                   pesan = pesan & vbCr & "Ada kesalahan pada master model. " & vbCr
101               End If
102           End If
103           If Text(24) = "" Then
104               boleh = False
105               pesan = pesan + "Warna. "
106           End If
107           If Text(26) = "" Then
108               boleh = False
109               pesan = pesan + "No Mesin. "
110           End If
111           If Text(27) = "" Then
112               boleh = False
113               pesan = pesan + "No Kunci. "
114           End If
115           If Trim(Text(45)) = "" Then
116               Text(45) = 0
117           End If
      '        If Text(59) = "" Then
      '            MsgBox "Alamat Email tidak boleh kosong" & Chr(13) & "Input konfirmasi untuk melanjutkan proses. (Jika Customer memang tidak punya alamat Email)", vbCritical, "Kelangkapan Data Email"
      '            If Trim(InputBox("Masukkan kata konfirmasi " & Chr(13) & "Konfirmasi : Pelanggan belum memiliki email")) <> "Pelanggan belum memiliki email" Then
      '                boleh = False
      '                pesan = pesan + "Email. "
      '            End If
      '        End If
              
118           If NulSID(NulSID(Text(62)) & " " & NulSID(Text(63)) & " " & NulSID(Text(64))) = "" Then
119               boleh = False
120               pesan = pesan + " Nama Surat Menyurat. "
121           End If
              
122           If NulSID(txt(0)) = "" Then
123               boleh = False
124               pesan = pesan + " Propinsi Surat Menyurat. "
125           End If
              
126           If NulSID(txt(1)) = "" Then
127               boleh = False
128               pesan = pesan + " Kota Surat Menyurat. "
129           End If
              
130           If NulSID(txt(2)) = "" Then
131               boleh = False
132               pesan = pesan + " Kecamatan Surat Menyurat. "
133           End If
              
134           If NulSID(txt(3)) = "" Then
135               boleh = False
136               pesan = pesan + " Kelurahan Surat Menyurat. "
137           End If
              
138           If NulSID(txt(6)) = "" Then
139               boleh = False
140               pesan = pesan + " Alamat Surat Menyurat. "
141           End If
              
      '        If nulsid(txt(5)) = "" Or Val(txt(5)) = 0 Then
      '            boleh = False
      '            pesan = pesan + " RT Surat Menyurat. "
      '        End If
      '
      '        If nulsid(txt(4)) = "" Or Val(txt(4)) = 0 Then
      '            boleh = False
      '            pesan = pesan + " RW Surat Menyurat. "
      '        End If
              
142           If NulSID(Text(46)) = "" Then
143               boleh = False
144               pesan = pesan + " No Polisi. "
145           End If
              
              '20/12/2012  Cust Non Axtif
              Dim rsNon As ADODB.Recordset
146           Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & IIf(NulSID(Text(65)) = "", Text(2), NulSID(Text(65))) & "' and nonaktifflag = 0")
147           If rsNon.RecordCount = 0 Then
148               boleh = False
149               pesan = "Kode Langgan : " & Text(2) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut"
150           End If
              
151           If boleh Then
                  Dim rsCari As ADODB.Recordset
152               sSQL = "SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(0)) & "' and KOTA = '" & NulSID(txt(1)) & "' and KECAMATAN = '" & NulSID(txt(2)) & "' and KELURAHAN = '" & NulSID(txt(3)) & "'"
153               Set rsCari = objO.FetchData(sSQL)
154               If rsCari.EOF Then
155                   MsgBox "PROPINSI / KOTA / KECAMATAN / KELURAHAN / KODEPOS tidak cocok", vbInformation
156                   Exit Sub
157               End If
158               Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 0)
159               erTranAll = False
160               Call objHT.OpenTransactionAll
                  
161               Call ClearsSql
162               gsSql1(0) = "NO_CHASSIS"
163               gsSql2(0) = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
164               gsSql1(1) = "FS_KD_PLG"
165               gsSql2(1) = IIf(NulSID(Text(65)) = "", Text(2), NulSID(Text(65)))
166               gsSql1(2) = "FS_KD_WIL"
167               gsSql2(2) = Text(3)
168               gsSql1(3) = "CUSTOMER"
169               gsSql2(3) = NulSID(Text(4)) & " " & NulSID(Text(60)) & " " & NulSID(Text(61))
170               dataLengkap = ""
      '            If cmdBack.Visible = True Then
171                   dataLengkap = NulSID(txt(8)) & " RT." & NulSID(txt(9)) & " RW." & NulSID(txt(10)) & " " & NulSID(txt(66)) & "," & NulSID(txt(65)) & "," & NulSID(txt(64)) & "," & NulSID(txt(63))
      '            Else
      '                dataLengkap = NulSID(txt(8)) & " RT." & NulSID(txt(9)) & " RW." & NulSID(txt(10)) & " " & cboKel.Text & "," & cboKec.Text & "," & cboKota.Text & "," & cboProv.Text
      '            End If
172               gsSql1(4) = "CUST_ADRES"
173               gsSql2(4) = dataLengkap
174               gsSql1(5) = "KOTA"
      '            If cmdBack.Visible = True Then
175                   gsSql2(5) = NulSID(txt(64))
      '            Else
      '                gsSql2(5) = cboKota.Text
      '            End If
176               gsSql1(6) = "KODEPOS"
177               gsSql2(6) = Text(7)
178               gsSql1(7) = "TELPON"
179               gsSql2(7) = Text(8)
180               gsSql1(8) = "GSM"
181               gsSql2(8) = Text(9)
182               gsSql1(9) = "CDMA"
183               gsSql2(9) = Text(10)
184               gsSql1(10) = "NPWP"
185               gsSql2(10) = Text(11)
                  
186               gsSql1(11) = "USEJOBDSC"
187               gsSql2(11) = NulSID(txt(7)) 'IIf(cboPek.Text = "LAINNYA", nulsid(txt(7)), cboPek.Text) ''Text(12)
188               gsSql1(12) = "USEAGAMA"
189               gsSql2(12) = NulSID(txt(12)) 'cboAgama.Text 'Text(13)
190               gsSql1(13) = "USEFS_NM_HOBI"
191               gsSql2(13) = NulSID(txt(13)) 'cboHobi.Text ' Text(14)
192               gsSql1(14) = "FS_KD_DEALER"
193               gsSql2(14) = Text(15)
194               gsSql1(15) = "FS_NM_DEALER"
195               gsSql2(15) = Text(16)
196               gsSql1(16) = "CODE"
197               gsSql2(16) = Text(17)
198               gsSql1(17) = "DDATE"
199               gsSql2(17) = Format(IIf(Replace(Replace(Text(18), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(18), ">", ""), "-", "/")), "dd/mm/yyyy")
200               gsSql1(18) = "DOTAM"
201               gsSql2(18) = Format(IIf(Replace(Replace(Text(19), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(19), ">", ""), "-", "/")), "mm/dd/yyyy")
202               gsSql1(19) = "FS_TH_MOBIL"
203               gsSql2(19) = Text(20)
204               gsSql1(20) = "MDL"
205               gsSql2(20) = Text(21)
206               gsSql1(21) = "FMDL"
207               gsSql2(21) = Text(23)
208               gsSql1(22) = "MODEL"
209               gsSql2(22) = Text(22)
210               gsSql1(23) = "COLOR"
211               gsSql2(23) = Text(24)
212               gsSql1(24) = "NO_POLISI"
213               gsSql2(24) = Trim(Text(25)) & IIf(Trim(Text(25)) <> "", "-", "") & Trim(Text(46)) & IIf(Trim(Text(47)) <> "", "-", "") & Text(47)
214               gsSql1(25) = "NO_ENGINE"
215               gsSql2(25) = Text(26)
216               gsSql1(26) = "KEYNO"
217               gsSql2(26) = Text(27)
218               gsSql1(27) = "PPN"
219               gsSql2(27) = SimpanKoma2Titik(Text(28))
220               gsSql1(28) = "DISC"
221               gsSql2(28) = Val(Text(29))
                  
                  ''--------------------------------------
                  ''SURAT MENYURAT
                  ''--------------------------------------
222               gsSql1(29) = "USENAME"
223               gsSql2(29) = NulSID(NulSID(Text(62)) & " " & NulSID(Text(63)) & " " & NulSID(Text(64))) 'Text(30)
224               dataLengkap = ""
      '210               If cmdBackS.Visible = True Then
225                   dataLengkap = NulSID(txt(6)) & " RT." & NulSID(txt(5)) & " RW." & NulSID(txt(4)) & " " & NulSID(txt(3)) & "," & NulSID(txt(2)) & "," & NulSID(txt(0))
      '212               Else
      '213                   dataLengkap = NulSID(Txt(6)) & " RT." & NulSID(Txt(5)) & " RW." & NulSID(Txt(4)) & " " & cboKelS.text & "," & cboKecS.text & "," & cboProvS.text
      '214               End If
226               gsSql1(30) = "USEADDR"
227               gsSql2(30) = dataLengkap 'Text(31) + Text(57) + Text(58)
228               gsSql1(31) = "USECITY"
229               gsSql2(31) = NulSID(txt(1)) 'IIf(cmdBackS.Visible = True, NulSID(Txt(1)), cboKota.text) 'Text(32)
230               gsSql1(32) = "USEPOS"
231               gsSql2(32) = Text(33)
232               gsSql1(33) = "USEPSTN"
233               gsSql2(33) = Text(34)
234               gsSql1(34) = "USEGSM"
235               gsSql2(34) = Text(35)
236               gsSql1(35) = "USECDMA"
237               gsSql2(35) = Text(36)
238               gsSql1(36) = "ASURANSI"
239               gsSql2(36) = Text(37)
240               gsSql1(37) = "ASR_ADDR"
241               gsSql2(37) = Text(38)
242               gsSql1(38) = "ASR_CITY"
243               gsSql2(38) = Text(39)
244               gsSql1(39) = "ASR_TGLM"
245               gsSql2(39) = Format(IIf(Replace(Replace(Text(40), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(40), ">", ""), "-", "/")), "mm/dd/yyyy")
246               gsSql1(40) = "ASR_TGLS"
247               gsSql2(40) = Format(IIf(Replace(Replace(Text(41), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(41), ">", ""), "-", "/")), "mm/dd/yyyy")
248               gsSql1(41) = "ASR_LIMIT"
249               gsSql2(41) = Text(42)
250               gsSql1(42) = "ASR_NOTE"
251               gsSql2(42) = Text(43)
252               gsSql1(43) = "NOTE"
253               gsSql2(43) = Text(44)
254               gsSql1(44) = "FS_KM_AKHIR"
255               gsSql2(44) = CDbl(Text(45))
256               gsSql1(45) = "USERADD"
257               gsSql2(45) = gSA
258               gsSql1(46) = "USET4LAHIR"
259               gsSql2(46) = Text(48)
260               gsSql1(47) = "USETGLLAHIR"
261               gsSql2(47) = Format(IIf(Replace(Replace(Text(49), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(49), ">", ""), "-", "/")), "mm/dd/yyyy")
262               gsSql1(48) = "PKP"
263               gsSql2(48) = Trim(Text(51))
264               gsSql1(49) = "KDPLG_ASR"
265               gsSql2(49) = Trim(Text(52))
266               gsSql1(50) = "USEEMAIL"
267               gsSql2(50) = Trim(Text(59))
                  
268               If Log.loyalityCustomer = 1 Then
269                   Set rsSearch = objHT.FetchDataBengkel("select * froM RplNasApp..UnitCustomer where NoRangka = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
270                   If rsSearch.EOF Then
271                       Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM RplNasApp..UnitCustomer WHERE IDCUSTSERVICE = '" & Trim(Text(2)) & "'")
272                       If Not rsSearch.EOF Then
273                           sSQL = "INSERT INTO [RplNasApp]..UnitCustomer(ddate,fmdl,noEngine,IDCustService,NoRangka,KdDlrPenjual,nmDlrPenjual,noPolisi,mdl,thmobil,userentry,tglentry)"
274                           sSQL = sSQL & " VALUES ('" & Format(IIf(Replace(Replace(Text(18), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(18), ">", ""), "-", "/")), "dd/mm/yyyy") & "','" & NulSID(Text(23)) & "','" & NulSID(Text(26)) & "','" & NulSID(Text(2)) & "','" & NulSID(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")) & "','" & NulSID(Text(15)) & "','" & NulSID(Text(16)) & "','" & NulSID(Trim(Text(25)) & IIf(Trim(Text(25)) <> "", "-", "") & Trim(Text(46)) & IIf(Trim(Text(47)) <> "", "-", "") & Text(47)) & "','" & NulSID(Text(21)) & "','" & NulSID(Text(20)) & "','" & gSA & "',getdate())"
275                           Call objHT.InsertData_Bengkel(sSQL)
276                       End If
                          
277                   End If
278               End If
                  
279               Set rsSearch = objO.FetchData("Select * from CUSTDAT WITH(NOLOCK) where NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' ")
280               If Not rsSearch.BOF Then
281                   Set rsTemp = objO.FetchData("SELECT * FROM HIST_PEMILIK WITH(NOLOCK) WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
282                   If rsTemp.EOF Then
283                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", False))
284                   Else
285                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", True) + " WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
286                   End If
287                   rsSearch.MoveFirst
                      
288                   sSQL = GensSql("CUSTDAT", True) + " where NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'"
289                   If rsSearch("ASURANSI") <> Trim(Text(37)) Then
290                       Set rsTemp = objO.FetchData("SELECT * FROM HIST_ASURANSI WITH(NOLOCK) WHERE fs_kd_plg = '" & Text(2) & "' AND no_Chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' ORDER BY urut DESC")
291                       If Not rsTemp.EOF Then
292                           If rsTemp.RecordCount > 1 Then
293                               Call objHT.UpdateData_Bengkel("UPDATE HIST_ASURANSI SET urut = '" & CDbl(rsTemp("urut")) + 1 & "' WHERE fs_kd_plg = '" & Text(2) & "' AND no_Chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
294                           ElseIf rsTemp.RecordCount = 1 Then
295                               Call objHT.UpdateData_Bengkel("UPDATE HIST_ASURANSI SET urut = '1' WHERE fs_kd_plg = '" & Text(2) & "' AND no_Chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
296                           End If
297                       End If
298                       If Trim(Text(37)) <> "" Then
299                           Call objHT.InsertData_Bengkel("INSERT INTO HIST_ASURANSI (NO_CHASSIS,FS_KD_PLG,CUSTOMER,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,KDPLG_ASR) VALUES ('" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "','" & Text(2) & "','" & Text(4) & "','" & Text(37) & "','" & Text(38) & "','" & Text(39) & "','" & Format(gsSql2(39), "MM/DD/YYYY") & "','" & Format(gsSql2(40), "MM/DD/YYYY") & "'," & CDbl(Text(42)) & ",'" & Text(43) & "','" & Trim(Text(52)) & "')")
300                       End If
301                   End If
302               Else
303                   sSQL = GensSql("CUSTDAT", False)
304                   Set rsTemp = objO.FetchData("SELECT * FROM HIST_PEMILIK WITH(NOLOCK) WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
305                   If rsTemp.EOF Then
306                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", False))
307                   Else
308                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", True) + "WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
309                   End If
310                   If Text(37) <> "" Then
311                       Call objHT.InsertData_Bengkel("INSERT INTO HIST_ASURANSI (NO_CHASSIS,FS_KD_PLG,CUSTOMER,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,KDPLG_ASR) VALUES ('" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "','" & Text(2) & "','" & Text(4) & "','" & Text(37) & "','" & Text(38) & "','" & Text(39) & "','" & Format(gsSql2(39), "MM/DD/YYYY") & "','" & Format(gsSql2(40), "MM/DD/YYYY") & "'," & CDbl(Text(42)) & ",'" & Text(43) & "','" & Trim(Text(52)) & "')")
312                   End If
313               End If
314               Call objHT.UpdateData_Bengkel(sSQL)
                  
                  ' UPDATE DATA CUSTOMER KE SPKCABANG
315               If Trim(Text(30)) <> "" Then
                      ' KLO BELUM ADA PDS TANYA MO PDS APA GAK, KLO PDS TIDAK UPDATE
316                   If Not pds Then
317                       If MsgBox("Akan melakukan transaksi PDS ?", vbDefaultButton2 + vbYesNo, "PDS") = vbNo Then
318                               If chLog = 1 Then
319                                   Call objHT.InsertData_SPK("INSERT INTO SCUST VALUES ('" & Trim(gsSql2(0)) & "','" & Trim(Text(30)) & "','" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "','" & Trim(Text(33)) & "','" & Trim(Text(32)) & "','','" & Trim(Text(34)) & "','" & Trim(Text(35)) & "','" & Trim(Text(36)) & "', '" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN") & "','" & Trim(Text(59)) & "')")
                                      'Call objHT.InsertData_SPK("INSERT INTO SCUST VALUES ('" & Trim(gsSql2(0)) & "','" & Trim(Text(30)) & "','" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "','" & Trim(Text(33)) & "','" & Trim(Text(32)) & "','','" & Trim(Text(34)) & "','" & Trim(Text(35)) & "','" & Trim(Text(36)) & "', '" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN") & "')")
320                               End If
321                               Call objHT.UpdateData_SPK("UPDATE SPK SET SNAMA = '" & Trim(Text(30)) & "', SALAMAT = '" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "', SKODEPOS = '" & Trim(Text(33)) & "', SKOTA = '" & Trim(Text(32)) & "',  " _
                                                  & "STELP = '" & Trim(Text(34)) & "', SGSM = '" & Trim(Text(35)) & "', SCDMA = '" & Trim(Text(36)) & "', EMAIL = '" & Trim(Text(59)) & "' WHERE NORANGKA = '" & Trim(gsSql2(0)) & "' AND SPK_STATUS <> '1'")
                                  'Call objHT.UpdateData_SPK("UPDATE SPK SET SNAMA = '" & Trim(Text(30)) & "', SALAMAT = '" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "', SKODEPOS = '" & Trim(Text(33)) & "', SKOTA = '" & Trim(Text(32)) & "',  " _
                                  '                & "STELP = '" & Trim(Text(34)) & "', SGSM = '" & Trim(Text(35)) & "', SCDMA = '" & Trim(Text(36)) & "' WHERE NORANGKA = '" & Trim(gsSql2(0)) & "' AND SPK_STATUS <> '1'")
322                       End If
323                   End If
324               End If
                  
                  ' UPDATE NO RANGKA DI BOOK
325               If gBantuan = "booking" Then
326                   Call objHT.UpdateData_Bengkel("UPDATE book SET no_chassis = '" & gsSql2(0) & "' WHERE nofak = '" & gBantuan2 & "'")
327               End If
                  
                  ''-------------------------------------------
                  '' 20171127 || UPDATE / INSERT UNIT CUSTOMER
                  ''-------------------------------------------
328               Call ClearsSql
329               gsSql1(0) = "IDCustSV"
330               gsSql2(0) = NulSID(Text(66))
                  
      '
      '            gsSql1(1) = "NoKTPSV"
      '            gsSql2(1) = ""
      '
      '            gsSql1(2) = "NoNPWPSV"
      '            gsSql2(2) = ""
      '
      '            gsSql1(3) = "TitlesSV"
      '            gsSql2(3) = ""
                  
331               gsSql1(1) = "NamaCust1SV"
332               gsSql2(1) = NulSID(Text(62))
                  
333               gsSql1(2) = "NamaCust2SV"
334               gsSql2(2) = NulSID(Text(63))
                  
335               gsSql1(3) = "NamaCust3SV"
336               gsSql2(3) = NulSID(Text(64))
                  
337               gsSql1(4) = "TglLahirSV"
338               gsSql2(4) = Format(IIf(Replace(Replace(Text(49), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(49), ">", ""), "-", "/")), "mm/dd/yyyy") 'Format(nulsid(Text(49)), "MM/DD/YYYY")
                  
339               gsSql1(5) = "AgamaSV"
340               gsSql2(5) = UCase(NulSID(txt(12)))
                  
341               gsSql1(6) = "PekerjaanSV"
342               gsSql2(6) = NulSID(txt(7)) 'IIf(cboPek.Text = "LAINNYA", nulsid(txt(7)), cboPek.Text)
                  
343               gsSql1(7) = "PropinsiSV"
344               gsSql2(7) = NulSID(txt(0)) 'IIf(cmdBackS.Visible, nulsid(txt(0)), cboProvS.Text)
                  
345               gsSql1(8) = "KotaSV"
346               gsSql2(8) = NulSID(txt(1)) 'IIf(cmdBackS.Visible, nulsid(txt(1)), cboKotaS.Text)
                  
347               gsSql1(9) = "KecamatanSV"
348               gsSql2(9) = NulSID(txt(2)) 'IIf(cmdBackS.Visible, nulsid(txt(2)), cboKecS.Text)
                  
349               gsSql1(10) = "KelurahanSV"
350               gsSql2(10) = NulSID(txt(3)) 'IIf(cmdBackS.Visible, nulsid(txt(3)), cboKelS.Text)
                  
351               gsSql1(11) = "KodePosSV"
352               gsSql2(11) = NulSID(Text(33))
                  
353               gsSql1(12) = "AlamatSV"
354               gsSql2(12) = NulSID(txt(6))
                  
355               gsSql1(13) = "RTSV"
356               gsSql2(13) = NulSID(txt(5))
                  
357               gsSql1(14) = "RWSV"
358               gsSql2(14) = NulSID(txt(4))
      '
      '            gsSql1(18) = "IdTelpSV"
      '            gsSql2(18) = ""
                  
359               gsSql1(15) = "TelpSV"
360               gsSql2(15) = Left(NulSID(Text(34)), 10)
                  
361               gsSql1(16) = "HP1SV"
362               gsSql2(16) = NulSID(Text(35))
                  
363               gsSql1(17) = "HP2SV"
364               gsSql2(17) = NulSID(Text(36))
                  
365               gsSql1(18) = "EmailSV"
366               gsSql2(18) = NulSID(Text(59))
                  
367               gsSql1(19) = "HobiSV"
368               gsSql2(19) = NulSID(txt(13))
                  
      '            gsSql1(24) = "OlahragaSV"
      '            gsSql2(24) = ""
      '
      '            gsSql1(25) = "SukuSV"
      '            gsSql2(0) = ""
      '
      '            gsSql1(26) = "PreferensiSV"
      '            gsSql2(0) = ""
                  
      '            gsSql1(202) = "noPolis"
      '            gsSql2(0) = ""
      '
      '            gsSql1(28) = "TglAwalPolis"
      '            gsSql2(0) = ""
      '
      '            gsSql1(29) = "TglAkhirPolis"
      '            gsSql2(0) = ""
                  
369               gsSql1(20) = "NoPolisi"
370               gsSql2(20) = Trim(Text(25)) & IIf(Trim(Text(25)) <> "", "-", "") & Trim(Text(46)) & IIf(Trim(Text(47)) <> "", "-", "") & Text(47)
                  
371               gsSql1(21) = "NoEngine"
372               gsSql2(21) = NulSID(Text(26))
                  
373               gsSql1(22) = "NoKey"
374               gsSql2(22) = NulSID(Text(27))
                  
375               gsSql1(23) = "MDL"
376               gsSql2(23) = Text(21)
                  
377               gsSql1(24) = "Class"
378               gsSql2(24) = Text(22)
                  
379               gsSql1(25) = "Fmdl"
380               gsSql2(25) = Text(23)
                  
381               gsSql1(26) = "Color"
382               gsSql2(26) = NulSID(Text(24))
                  
383               gsSql1(27) = "KdDlrPenjual"
384               gsSql2(27) = NulSID(Text(15))
                  
385               gsSql1(28) = "NmDlrPenjual"
386               gsSql2(28) = NulSID(Text(16))
                  
387               gsSql1(29) = "dDate"
388               gsSql2(29) = Format(NulSID(Text(18)), "mm/dd/yyyy")
                  
389               gsSql1(30) = "thMobil"
390               gsSql2(30) = NulSID(Text(20))
                  
391               gsSql1(31) = "TglEntry"
392               gsSql2(31) = "GETDATE()"
                  
393               gsSql1(32) = "UserEntry"
394               gsSql2(32) = gSA
                   
                  
395               gsSql1(33) = "FS_KM_AKHIR"
396               gsSql2(33) = NulSID(Text(45))
                  
397               gsSql1(34) = "NoRangka"
398               gsSql2(34) = NulSID(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
                              
399               gsSql1(35) = "IDCustService"
400               gsSql2(35) = NulSID(Text(2))
                  
401               gsSql1(36) = "CexSurat"
402               gsSql2(36) = IIf(chSda.value = Checked, 1, 0)
                  
                  
                  
403               Set rsSearch = objHT.FetchDataBengkel("Select * from [SINGLEIDORI]..UNITCUSTOMER WITH(NOLOCK) where NoRangka ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' ")
404               If rsSearch.EOF Then
405                   Call objHT.UpdateData_Bengkel(GensSql2("[SINGLEIDORI]..UNITCUSTOMER", False))
406               Else
407                   gsSql1(37) = "IDCustSV"
408                   gsSql2(37) = NulSID(Text(66))
409                   Call objHT.UpdateData_Bengkel(GensSql2("[SINGLEIDORI]..UNITCUSTOMER", True) + " WHERE NoRangka ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
410               End If
                  
                  ''-------------------------------------------
                  
                  ' TRANSFER HISTORY DARI BENGKEL LAMA
               '   Dim noRgk As String
               '   Dim rsHlp As ADodb.Recordset
                  
               '      noRgk = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
               '           Set rsHlp = objO.FetchDataAAS("SELECT * FROM CUSTDAT WHERE NO_CHASSIS = '" & noRgk & "'")
               '           If Not rsHlp.EOF Then
               '               If MsgBox("Akan transfer history Bengkel Lama ?", vbDefaultButton2 + vbYesNo, "PDS") = vbYes Then
               '                   Call transferHistoryLama(noRgk, rsHlp)
               '                  updateHistory = False
               '               Else
               '                   updateHistory = True
               '               End If
               '           End If
                      
                      
                              
411               If erTranAll = False Then
412                   Call objHT.CloseTransactionAll
413                   Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 1)
414               Else
415                   Call objHT.CloseTransactionAll_Err
416                   Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 2)
417                   MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
418               End If
419               With frmOrder
420                   If .Visible Then
                          
                          
421                       .Text(1) = Text(2)
422                       .Text(14) = gsSql2(20)
423                       .txtKodeLgnLama = IIf(NulSID(Text(65)) = "", Text(2), NulSID(Text(65)))
                          '.text(0).SetFocus
424                       If .cmdPilih.Tag = "No. Rangka" Then
425                       .Text(0) = NulSID(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
                          
426                       .cmdPilih.Tag = "No. Rangka"
427                       .cmdPilih.Picture = .ImageList1.ListImages(2).Picture
428                       Else
429                           If NulSID(.Text(0)) = "" Then
                                  
430                               .Text(0) = NulSID(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
                                  
431                               .cmdPilih.Tag = "No. Rangka"
432                               .cmdPilih.Picture = .ImageList1.ListImages(2).Picture
433                           End If
434                       End If
                          
435                       Unload Me
                          
436                   End If
437               End With
                  
438               Text(25) = ""
439               Call hapusForm
440               Unload Me
441           Else
442               MsgBox "Lengkapi field berikut ..." & vbCr & pesan, vbCritical, "Field Mandatory"
443           End If
444       End If
445       Exit Sub
          
vbError:
446       If erTranAll = True Then
447           Call objHT.CloseTransactionAll_Err
448           Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 2)
449           MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmSIDCustomer"), vbInformation, "Gagal Simpan"
450       End If
451       Exit Sub
End Sub

Private Sub Form_Activate()
    ''20171221 || REVISI MENAMPILKAN UNIT LANGSUNG JIKA SUDAH ADA DATA
    ''------------------------------------------------------------------------
    If isLoadAktif = False And NulSID(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")) <> "" Then
        Call panggilUnitCustomer(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
        isLoadAktif = True
    End If
End Sub

Private Sub Form_Load()
    
    Call hapusForm
    pds = False
    chAsuransi.value = Unchecked
    chAsuransi_Click
    Call refreshComboBox
    hideComboProv (False)
    Call hideComboProvS(False)
    'Call loadComboBox
End Sub

Private Sub refreshComboBox()
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
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub OpExt_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub OpIn_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub Text_GotFocus(Index As Integer)
          Dim DiTab As Boolean
1         DiTab = False
              
2         If Text(Index).BackColor = &HC0E0FF Then
3             DiTab = True
4         Else
5             If gMainKeyCari <> "" Then
6                 If Index = 0 Then
7                     If gBantuan <> "" Then
8                         If NulSID(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")) <> gBantuan Then
9                             If MsgBox("Ada perubahan no rangka." & vbCrLf & " Apakah ingin mengikuti perubahan?", vbYesNo) = vbYes Then
10                                Text(Index) = Right(gBantuan, 8)
11                                Text(1) = Mid(gBantuan, 4, Len(gBantuan) - 11)
12                                Text(50) = Left(gBantuan, 3)
13                            End If
14                        End If
                          
15                    End If
16                Else
17                    Text(Index) = gMainKeyCari
18                End If
                  
              
19                gMainKeyCari = ""
20                gBantuan = ""
21                DiTab = True
22            End If
23        End If
24        If DiTab Then
25            Sendkeys ("{TAB}")
26        Else
27            Sendkeys ("{HOME}+{END}")
28        End If
          
29        Select Case Index
              Case 50
30                If Text(Index) <> "" And Text(1) <> "" Then
31                    PanggilCustomer
32                End If
33        End Select
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                Call tampilkanFormSearch
                Text(Index).SetFocus
            Case 2
                Call tampilkanFormSearch
                Text(Index).SetFocus
            Case 66
                Call tampilkanFormSearchSurat
                Text(Index).SetFocus
            Case 13
                gMainKeyCari = Text(Index)
                gSearchOption = "AGAMA"
                frmCari.show (vbModal)
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "HOBI"
                frmCari.show (vbModal)
            Case 15
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 21
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 24
                If Text(21) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "WARNA"
                    gBantuan = Trim(Text(21))
                    frmCariSPK.show (vbModal)
                Else
                    MsgBox "Isi Model Kendaraan"
                    Text(21).SelLength = Len(Text(21))
                    Text(21).SetFocus
                    Exit Sub
                End If
            Case 37
                gMainKeyCari = Text(Index)
                gSearchOption = "ASURANSI"
                frmCari.show (vbModal)
            Case 50
                gMainKeyCari = Text(Index)
                gSearchOption = "COUNTRY"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 7, 33, 46
            KeyAscii = CekNumerik(KeyAscii)
        Case 47, 50, 1, 0
            KeyAscii = IIf(CekNumerik(KeyAscii) = 0, CekAlfabet(KeyAscii), CekNumerik(KeyAscii))
    End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 20
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
                Text(Index).SetFocus
                Exit Sub
            End If
            If CDbl(Trim(Text(Index))) > year(isnow) Or CDbl(Trim(Text(Index))) < (year(isnow) - 60) Then
                MsgBox "Tahun Kendaraan salah"
                Text(Index) = ""
                Text(Index).SetFocus
            End If
        Case 42
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
                Exit Sub
            End If
        Case 45
            If Text(Index) = "" Then
                Text(Index) = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub loadComboBox()
    Dim rsA As New ADODB.Recordset
    
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
End Sub

Private Sub panggilUnitCustomer(noRangka As String)
1         Call PanggilCustomer
           ''panggil data unit customer berdasarkan norangka
2          sSQL = "select NoRangka,IDCustService,IDCustSV,NoKTPSV,NoNPWPSV,"
3          sSQL = sSQL & " TitlesSV , NamaCust1SV, NamaCust2SV, NamaCust3SV, TglLahirSV"
4          sSQL = sSQL & " ,AgamaSV,PekerjaanSV,PropinsiSV,KotaSV,KecamatanSV,KelurahanSV,"
5          sSQL = sSQL & " KodePosSV , AlamatSV, RTSV, RWSV, IdTelpSV, TelpSV"
6          sSQL = sSQL & " ,HP1SV,HP2SV,EmailSV,HobiSV,OlahragaSV,SukuSV,PreferensiSV,NoPolis,"
7          sSQL = sSQL & " TglAwalPolis , TglAkhirPolis, NoPolisi, NoEngine, NoKey, MDL, Class, Fmdl"
8          sSQL = sSQL & " ,Color,KdDlrPenjual,NmDlrPenjual,Ddate,Thmobil,TglEntry,UserEntry,KodeColor"
9          sSQL = sSQL & " from [SingleIDOri]..UnitCustomer with (nolock) where NoRangka = '" & noRangka & "'"
           
           'Call loadComboBox
           
           '' UNIT CUSTOMER || --------------- ""
10         Set rsTemp = objO.FetchData(sSQL)
11         If Not rsTemp.EOF Then
12            rsTemp.MoveFirst '
              
13            Call disableSuratMenyurat(True)
14            Text(62) = NulSID(rsTemp!NamaCust1SV)
15            Text(63) = NulSID(rsTemp!NamaCust2SV)
16            Text(64) = Left(NulSID(rsTemp!NamaCust3SV), 31)
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & nulsid(rsTemp!PropinsiSV) & "'")
      '        If Not rsHelp.EOF Then
      '            cboProvS.Text = UCase(nulsid(rsTemp!PropinsiSV))
      '            Call cboProvS_Change
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT Kota FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE kota = '" & nulsid(rsTemp!KotaSV) & "'")
      '            If Not rsHelp.EOF Then
      '            cboKotaS.Text = UCase(nulsid(rsTemp!KotaSV))
      '            End If
      '            Call cboKotaS_Change
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT Kecamatan FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE kota = '" & nulsid(rsTemp!KotaSV) & "' and Kecamatan = '" & nulsid(rsTemp!KecamatanSV) & "'")
      '            If Not rsHelp.EOF Then
      '            cboKecS.Text = UCase(nulsid(rsTemp!KecamatanSV))
      '            End If
      '            Call cboKecS_Change
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT Kelurahan FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE Kecamatan = '" & nulsid(rsTemp!KecamatanSV) & "' and Kelurahan = '" & nulsid(rsTemp!KelurahanSV) & "'")
      '            If Not rsHelp.EOF Then
      '            cboKelS.Text = UCase(nulsid(rsTemp!KelurahanSV))
      '            End If
      '            Call cboKel_Change
      '        Else
17                Call hideComboProvS(False)
18                txt(0) = UCase(NulSID(rsTemp!PropinsiSV))
19                txt(1) = UCase(NulSID(rsTemp!KotaSV))
20                txt(2) = UCase(NulSID(rsTemp!KecamatanSV))
21                txt(3) = UCase(NulSID(rsTemp!KelurahanSV))
      '        End If
              
22            txt(6) = NulSID(rsTemp!AlamatSV)
23            txt(5) = Format(NulSID(rsTemp!RTSV), "0##") 'nulsid(rsTemp!RTSV)
24            txt(4) = Format(NulSID(rsTemp!RWSV), "0##") 'nulsid(rsTemp!RWSV)
              
25            Text(33) = NulSID(rsTemp!KodePosSV)
              
26            Text(34) = NulSID(rsTemp!IdTelpSV) & NulSID(rsTemp!TelpSV)
27            Text(35) = NulSID(rsTemp!HP1SV)
28            Text(36) = NulSID(rsTemp!HP2SV)
              
      '        Set rsHelp = objO.FetchData("Select Upper(Pekerjaan) as Pekerjaan from [SingleIdOri]..Pekerjaan With (NoLock) WHERE Pekerjaan = '" & nulsid(rsTemp!Pekerjaansv) & "'")
      '        If Not rsHelp.EOF Then
      '            cboPek.Text = UCase(nulsid(rsTemp!Pekerjaansv))
      '        Else
      '            cboPek.Text = "LAINNYA"
      '            Call cboPek_Change
      '
29                txt(7) = UCase(NulSID(rsTemp!Pekerjaansv))
      '        End If
          
              'If UCase(nulsid(rsTemp!Hobisv)) <> "" Then
30                txt(13) = UCase(NulSID(rsTemp!Hobisv))
              'End If
          
              'If UCase(nulsid(rsTemp!Agamasv)) <> "" Then
31                txt(12) = UCase(NulSID(rsTemp!Agamasv))
              'End If
              
32            Text(59) = UCase(NulSID(rsTemp!Emailsv))
33            Text(49) = Format(NulSID(rsTemp!tglLahirsv), "dd-mm-yyyy")
              
              
              'Text(59) = nulsid(rsTemp!)
34            Text(15) = NulSID(rsTemp!KdDlrPenjual)
35            Text(16) = NulSID(rsTemp!NmDlrPenjual)
              
36            Text(18) = Format(NulSID(rsTemp!dDate), "DD-MM-YYYY")
37            Text(18) = Format(NulSID(rsTemp!dDate), "DD-MM-YYYY")
              
38            Text(19) = "01-01-3000"
39            Text(20) = NulSID(rsTemp!thMobil)
              
40            Text(21) = NulSID(rsTemp!MDL)
              'Call text_Validate(21, False)
              '20180221 agar tidak muncul show modaly
              
41            If Log.RPLHistori = 0 Then
42                Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(21) & "'")
43            Else
44                Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(21) & "'")
45            End If
46            If Not rsSearch.EOF Then
47                Text(21) = rsSearch("MDL")
48                Text(23) = rsSearch("Desc")
49                Text(22) = rsSearch("MODEL")
50                gGroupModel = rsSearch("FS_GRP_MDL")
51            End If
52            Text(24) = NulSID(rsTemp!Color)
              
53            Text(26) = NulSID(rsTemp!NoEngine)
54            Text(27) = NulSID(rsTemp!nokey)
55            If Len(NulSID(rsTemp("NOPOLISI"))) > 0 Then
56                nPol = NulSID(rsTemp("NOPOLISI"))
                  Dim nDpn1 As String
                  Dim nTgh1 As String
                  Dim nBlk1 As String
57                nDpn1 = ""
58                nTgh1 = ""
59                nBlk1 = ""
60                Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
61                Text(25) = Trim(nDpn1)
62                Text(46) = Trim(nTgh1)
63                Text(47) = Trim(nBlk1)
64            Else
65                Text(25) = ""
66                Text(46) = ""
67                Text(47) = ""
68            End If
69            If NulSID(rsTemp!idcustservice) <> "" Then
70                Text(2) = NulSID(rsTemp!idcustservice)
71            End If
72            Text(66) = ""
73            If NulSID(rsTemp!IDCustSV) <> "" Then
74                Text(66) = NulSID(rsTemp!IDCustSV)
75                Call text_Validate(66, False)
76            End If
              
77            Call panggilCustomerSID
78         Else
              'Call tampilkanFormSearch
79         End If
           ''------------------------------------
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 0
3                 If Text(Index) <> "" Then
      '                Set rsSearch = objO.FetchData("select * from custdat WITH(NOLOCK) where no_chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
      '                If rsSearch.EOF Then
      '                    If Text(Index) = "" Then
      '                        Cancel = True
      '                        gMainKeyCari = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
      '                        gSearchOption = "CUSTOMER2"
      '                        frmCari.Show (vbModal)
      '                    End If
      '                    PanggilCustomer
      '                Else
                          'PanggilCustomer
      '                End If
4                     Call panggilUnitCustomer(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
                      
5                 End If
6             Case 2
7                 If Trim(Text(Index)) <> "" Then
                  
8                 Call panggilCustomerSID
      '            Text(3) = Left(Text(Index), 3)
      '
      '                    'Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Armst where KodeLgn = '" & Trim(text(Index)) & "'")
      '                    '20/12/2012 Cust Non Axtif
      '                    Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Armst WITH(NOLOCK) where KodeLgn = '" & Trim(Text(Index)) & "' and nonaktifflag  = 0")
      '                    If rsSearch.BOF Then
      '                        temp = MsgBox("Tambah Kode Pelanggan baru ?", vbYesNo, "KODE PELANGGAN")
      '                        If temp = vbYes Then
      '                            cmdPelanggan_Click
      '                        Else
      '                            Cancel = True
      '                            gMainKeyCari = Text(Index)
      '                            gSearchOption = "KODEPLG"
      '                            frmCariAcess.Show (vbModal)
      '                        End If
      '                    Else
      '                        If IsNull(rsSearch("NamaLgn")) Then
      '                            Text(4) = ""
      '                        Else
      '                            Text(4) = Left(Replace(Trim(rsSearch("NamaLgn")), vbNewLine, ""), 60)
      '                        End If
      '                        Text(5) = Left(Trim(nulsid(rsSearch("Alamat1"))) & " " & Trim(nulsid(rsSearch("Alamat2"))), 60)
      '
      '                            If Len(Trim(nulsid(rsSearch("Alamat1"))) & " " & Trim(nulsid(rsSearch("Alamat2")))) > 60 Then
      '                                Text(56) = Mid(Trim(nulsid(rsSearch("Alamat1"))) & " " & Trim(nulsid(rsSearch("Alamat2"))), 61, Len(Trim(nulsid(rsSearch("Alamat1"))) & " " & Trim(nulsid(rsSearch("Alamat2")))) - 60)
      '                            Else
      '                                Text(56) = ""
      '                            End If
      '
      '                        If IsNull(rsSearch("Alamat3")) Then
      '                            Text(6) = ""
      '                        Else
      '                            Text(6) = rsSearch("Alamat3")
      '                        End If
      '                        If IsNull(rsSearch("Telepon")) Then
      '                            Text(8) = ""
      '                            Text(9) = ""
      '                            Text(10) = ""
      '                        Else
      '                            Dim nDpn1 As String
      '                            Dim nTgh1 As String
      '                            Dim nBlk1 As String
      '                            nDpn1 = ""
      '                            nTgh1 = ""
      '                            nBlk1 = ""
      '                            Call pecahNoTelp(Trim(rsSearch("Telepon")), nDpn1, nTgh1, nBlk1, ",")
      '                            Text(8) = Trim(nDpn1)
      '                            Text(9) = Trim(nTgh1)
      '                            Text(10) = Trim(nBlk1)
      '                        End If
      '                        If IsNull(rsSearch("Npwp")) Then
      '                            Text(11) = ""
      '                        Else
      '                            Text(11) = rsSearch("Npwp")
      '                        End If
      '                        If IsNull(rsSearch("KodeWil")) Then
      '                            Text(3) = ""
      '                        Else
      '                            Text(3) = rsSearch("KodeWil")
      '                        End If
      '                    End If
      '                If Log.loyalityCustomer = 1 Then
      '                    Set rsTemp = objO.FetchData("SELECT a.alamatemail as email FROM RPLNASAPP..MASTERUSER a INNER JOIN RPLNASAPP..NewCustomer b ON a.alamatemail = b.email WHERE a.verifikasi = 1 and b.idcust = '" & Trim(Text(2)) & "' ")
      '                    If Not rsTemp.EOF Then
      '                        Text(59).Enabled = False
      '                        Text(59) = nulsid(rsTemp!Email)
      '                    End If
      '
      '                End If
      '                Set rsSearch = objO.FetchDataACC("select KodeWil,NamaWil from Arwilayah WITH(NOLOCK) where KodeWil = '" & Left(Text(3), 3) & "'")
      '                If Not rsSearch.BOF Then
      '                    Text(6) = rsSearch("NamaWil")
      '                End If
      '
9                 End If
10            Case 13
11                If Text(Index) <> "" Then
12                    Set rsSearch = objO.FetchData("select * from t_agama WITH(NOLOCK) where agama = '" & Text(Index) & "'")
13                    If rsSearch.BOF Then
14                        Cancel = True
15                        gMainKeyCari = Text(Index)
16                        gSearchOption = "AGAMA"
17                        frmCari.show (vbModal)
18                    End If
19                End If
20            Case 14
21                If Text(Index) <> "" Then
22                    Set rsSearch = objO.FetchData("select * from t_hobi WITH(NOLOCK) where fs_nm_hobi = '" & Text(Index) & "'")
23                    If rsSearch.BOF Then
24                        Cancel = True
25                        gMainKeyCari = Text(Index)
26                        gSearchOption = "HOBI"
27                        frmCari.show (vbModal)
28                    End If
29                End If
30            Case 15
31                If Text(Index) <> "" Then
32                    Set rsSearch = objO.FetchData("select * from t_dealer WITH(NOLOCK) where fs_kd_dealer = '" & Text(Index) & "' and FS_LOG_DEL = ''")
33                    If rsSearch.BOF Then
34                        Cancel = True
35                        gMainKeyCari = Text(Index)
36                        gBantuan = "ALL"
37                        gSearchOption = "DEALER"
38                        frmCari.show (vbModal)
39                    Else
40                        Text(16) = rsSearch("FS_NM_DEALER")
41                    End If
42                Else
43                    Text(16) = ""
44                End If
45            Case 18, 19, 49
46                If Text(Index) <> ">>->>->>>>" Then
47                    If Not CekTgl(Text(Index)) Then
48                        Text(Index) = ">>->>->>>>"
49                        Sendkeys ("+{TAB}")
50                        Exit Sub
51                    End If
52                End If
53            Case 21
54                If Text(Index) <> "" Then
                      
55                    If Log.RPLHistori = 0 Then
56                        Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(Index) & "'")
57                    Else
58                        Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(Index) & "'")
59                    End If
60                    If rsSearch.BOF Then
61                        Cancel = True
62                        gMainKeyCari = Text(Index)
63                        gSearchOption = "MDL"
64                        frmCari.show (vbModal)
65                        gGroupModel = ""
66                    Else
67                        Text(21) = rsSearch("MDL")
68                        Text(23) = rsSearch("Desc")
69                        Text(22) = rsSearch("MODEL")
70                        gGroupModel = rsSearch("FS_GRP_MDL")
      '                    If Text(1) <> Trim(rsSearch("NO_CHASSIS")) And Text(1) <> "" Then
      '                        If MsgBox("No Rangka Berubah !" & vbCr & "Mengikuti perubahan ?", vbYesNo, "UBAH KODE RANGKA") = vbYes Then
      ''                            Text(1) = rsSearch("NO_CHASSIS")
      '                        End If
      '                    ElseIf Text(1) = "" Then
      '                        Text(1) = nulsid(rsSearch("NO_CHASSIS"))
      '                    End If
71                    End If
72                End If
73            Case 37
74                If Text(Index) <> "" Then
75                    Set rsSearch = objO.FetchData("select * from t_asuransi WITH(NOLOCK) where asuransi = '" & Trim(Text(Index)) & "'")
76                    If rsSearch.BOF Then
77                        Cancel = True
78                        gMainKeyCari = Text(Index)
79                        gSearchOption = "ASURANSI"
80                        frmCari.show (vbModal)
81                        Text(38) = ""
82                        Text(39) = ""
83                    Else
84                        Text(38) = Trim(rsSearch("Alamat"))
85                        Text(39) = Trim(rsSearch("Kota"))
86                        Text(52) = Trim(rsSearch("KDPLG_ASR"))
87                        Text(53) = Trim(rsSearch("PSTN"))
88                        Text(54) = Trim(rsSearch("GSM"))
89                        Text(55) = Trim(rsSearch("CDMA"))
90                    End If
91                End If
92            Case 45
93                If Text(Index) <> "" Or Text(Index) <> "0" Then
94                    Set rsSearch = objO.FetchData("SELECT * FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Replace(Trim(Text(50)), ">", "") & Trim(Replace(Text(1), ">", "")) & Trim(Replace(Text(0), ">", "")) & "' ORDER BY KM")
95                    If Not rsSearch.EOF Then
96                        If CDbl(Text(Index)) < CDbl(rsSearch("KM")) Then
97                            MsgBox "KM tidak boleh lebih kecil dari KM Terakhir"
98                            Text(Index) = CDbl(rsSearch("KM"))
99                        End If
100                   End If
101               End If
102           Case 47
103               If Text(Index) <> "" Then
104                   nPol = Trim(Replace(Text(25), ">", "")) & Trim(Replace(Text(46), ">", "")) & Trim(Replace(Text(47), ">", ""))
105                   Set rsSearch = objO.FetchData("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI = '" & Trim(nPol) & "' AND NO_CHASSIS <> '" & Replace(Trim(Text(50)), ">", "") & Trim(Replace(Text(1), ">", "")) & Trim(Replace(Text(0), ">", "")) & "'")
106                   If Not rsSearch.EOF Then
107                       MsgBox "No. Polisi digunakan unit kendaraan lain !!!" & vbCr & "Pastikan No. Polisi yang anda masukkan benar !!"
108                       Text(25) = ""
109                       Text(46) = ""
110                       Text(47) = ""
111                       Text(25).SetFocus
112                       Exit Sub
113                   End If
114               End If
115           Case 50
116               If Text(Index) <> "" Then
117                   If Log.RPLHistori = 0 Then
118                       Set rsSearch = objO.FetchData("SELECT * FROM T_COUNTRY WITH(NOLOCK) WHERE KODE = '" & Trim(Text(Index)) & "'")
119                   Else
120                       Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_COUNTRY WITH(NOLOCK) WHERE KODE = '" & Trim(Text(Index)) & "'")
121                   End If
122                   If rsSearch.EOF Then
123                      If MsgBox("Unit bukan Kendaraan Toyota ?", vbYesNo, "Konfirmasi No Rangka") = vbNo Then
124                           gMainKeyCari = Text(Index)
125                           gSearchOption = "COUNTRY"
126                           frmCari.show (vbModal)
127                      End If
128                   End If
129               End If
130           Case 66
131               If NulSID(Text(Index)) <> "" Then
132                   Set rsSearch = objO.FetchData("SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH(NOLOCK) WHERE IDCUST = '" & Trim(Text(Index)) & "'")
133                   If Not rsSearch.EOF Then
134                       Text(62) = NulSID(rsSearch!namacust1)
135                       Text(63) = NulSID(rsSearch!namacust2)
136                       Text(64) = Left(NulSID(rsSearch!namacust3), 31)
      '                    Dim rsSearchAlamat As ADODB.Recordset
      '                    Set rsSearchAlamat = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE PROPINSI = '" & nulsid(rsSearch!Propinsi) & "'")
      '                    If Not rsSearchAlamat.EOF Then
      '                        Call hideComboProvS(True)
      '                        cboProvS.Text = UCase(nulsid(rsSearch!Propinsi))
      '                        Call cboProvS_Change
      '
      '                        Set rsSearchAlamat = objO.FetchData("SELECT DISTINCT Kota FROM [SingleIdOri]..TopologyDaerah WHERE kota = '" & nulsid(rsSearch!Kota) & "'")
      '                        If Not rsSearchAlamat.EOF Then
      '                            cboKotaS.Text = UCase(nulsid(rsSearch!Kota))
      '                        End If
      '                        Call cboKotaS_Change
      '
      '                        Set rsSearchAlamat = objO.FetchData("SELECT DISTINCT Kecamatan FROM [SingleIdOri]..TopologyDaerah WHERE kota = '" & nulsid(rsSearch!Kota) & "' and Kecamatan = '" & nulsid(rsSearch!kecamatan) & "'")
      '                        If Not rsSearchAlamat.EOF Then
      '                            cboKecS.Text = UCase(nulsid(rsSearch!kecamatan))
      '                        End If
      '                        Call cboKecS_Change
      '
      '                        Set rsSearchAlamat = objO.FetchData("SELECT DISTINCT Kelurahan FROM [SingleIdOri]..TopologyDaerah WHERE Kecamatan = '" & nulsid(rsSearch!kecamatan) & "' and Kelurahan = '" & nulsid(rsSearch!kelurahan) & "'")
      '                        If Not rsSearchAlamat.EOF Then
      '                            cboKelS.Text = UCase(nulsid(rsSearch!kelurahan))
      '                        End If
      '                        Call cboKelS_Change
      '
      '                    Else
137                           Call hideComboProvS(False)
138                           txt(0) = UCase(NulSID(rsSearch!Propinsi))
139                           txt(1) = UCase(NulSID(rsSearch!Kota))
140                           txt(2) = UCase(NulSID(rsSearch!kecamatan))
141                           txt(3) = UCase(NulSID(rsSearch!kelurahan))
                           
      '                    End If
142                       txt(6) = NulSID(rsSearch!Alamat)
143                       txt(5) = NulSID(rsSearch!rt)
144                       txt(4) = NulSID(rsSearch!rw)
145                       Text(33) = NulSID(rsSearch!KodePos)
146                       Text(34) = NulSID(rsSearch!IdTelp) & NulSID(rsSearch!telp)
147                       Text(35) = NulSID(rsSearch!hp1)
148                       Text(36) = NulSID(rsSearch!hp2)
                          
      '                    Set rsHelp = objO.FetchData("Select Upper(Pekerjaan) as Pekerjaan from [SingleIdOri]..Pekerjaan With (NoLock) WHERE Pekerjaan = '" & nulsid(rsSearch!Pekerjaan) & "'")
      '                    If Not rsHelp.EOF Then
      '                        cboPek.Text = UCase(nulsid(rsSearch!Pekerjaan))
      '                    Else
      '                        cboPek.Text = "LAINNYA"
      '                        Call cboPek_Change
      '
149                           txt(7) = UCase(NulSID(rsSearch!pekerjaan))
      '                    End If
                      
                          'If UCase(nulsid(rsSearch!Hobi)) <> "" Then
150                       txt(13) = UCase(NulSID(rsSearch!hobi))
                          'End If
                      
                          'If UCase(nulsid(rsSearch!Agama)) <> "" Then
151                       txt(12) = UCase(NulSID(rsSearch!Agama))
                          'End If
                          
152                       Text(59) = UCase(NulSID(rsSearch!email))
153                       Text(49) = Format(NulSID(rsSearch!tglLahir), "dd-mm-yyyy")
                          
154                       Call disableSuratMenyurat(False)
155                   Else
156                       MsgBox "IDCust " & NulSID(Text(Index)) & " tidak ditemukan", vbInformation
157                       Text(66) = ""
158                   End If
159               End If
160           Case 33
161               If NulSID(Text(Index)) <> "" Then
                      Dim rsTopology As ADODB.Recordset
162                   Set rsTopology = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE KODEPOS = '" & NulSID(Text(Index)) & "'")
163                   If Not rsTopology.EOF Then
164                       rsTopology.MoveFirst
165                       txt(0) = NulSID(rsTopology!Propinsi)
166                       txt(1) = NulSID(rsTopology!Kota) 'nulsid(DGrid.Columns(1))
167                       txt(2) = NulSID(rsTopology!kecamatan) 'nulsid(DGrid.Columns(2))
168                       txt(3) = NulSID(rsTopology!kelurahan) 'nulsid(DGrid.Columns(3))
169                       Text(Index) = NulSID(rsTopology!KodePos)
170                   Else
171                       MsgBox "Data kodepos tidak terdaftar.", vbInformation
172                       txt(0) = ""
173                       txt(1) = "" 'nulsid(DGrid.Columns(1))
174                       txt(2) = "" 'nulsid(DGrid.Columns(2))
175                       txt(3) = "" 'nulsid(DGrid.Columns(3))
176                       Text(Index) = "" 'nulsid(DGrid.Columns(4))
177                   End If
178               End If
179       End Select

180      On Error GoTo 0
181      Exit Sub

text_Validate_Error:

182       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmSIDCustomer")
End Sub

Private Sub hapusForm()
          Dim i As Integer
1         isLoadAktif = False
2        On Error GoTo hapusForm_Error
3         txt(8) = ""
4         txt(9) = ""
5         txt(10) = ""
6         For i = 63 To 66
7             txt(i) = ""
8         Next
          
9         For i = 0 To 6
10            txt(i) = ""
11        Next
          
12        For i = 62 To 66
13            Text(i) = ""
14        Next
          
          'cboProv.Text = ""
15        cboProv.Clear
16        cboKota.Clear
17        cboKec.Clear
18        cboKel.Clear
19        isDariSearch = False
20        Call refreshComboBox
          
21        For i = 0 To 59
22            Select Case i
                  Case 8, 9, 10, 34, 35, 36
23                    Text(i) = ">>>>>>>>>>>>>>"
24                Case 25
25                    Text(i) = ""
26                    Text(i).Mask = "??"
27                Case 46
28                    Text(i) = ""
29                    Text(i).Mask = "9999"
30                Case 47
31                    Text(i) = ""
32                    Text(i).Mask = "CCC"
33                Case 18, 19, 40, 41, 49
34                    n = Text(i).Mask
35                    Text(i).Mask = ""
36                    Text(i) = ""
37                    Text(i).Mask = n
38                Case 20, 28, 29, 42, 45
39                    Text(i) = 0
40                Case Else
41                    Text(i) = ""
42            End Select
43        Next

44       On Error GoTo 0
45       Exit Sub

hapusForm_Error:

46        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmSIDCustomer")
          
End Sub

Private Function cekHistoryHO(noRgk As String) As Boolean
Dim rsHlp As ADODB.Recordset
    '  Dim rsTemp As ADODB.Recordset

    On Error GoTo vbError
    
    If Log.RPLHistori = 1 Then
        Set rsHlp = objO.FetchData("SELECT TOP 1 CUSTOMER,LEFT(cust_adres,120) as CUST_ADRES,KOTA, '' AS kodepos, TELPON, NO_POLISI, " & _
                            "NO_ENGINE, '' AS wmi, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, FS_KD_DEALER, " & _
                            "FS_NM_DEALER, DDATE, '' AS remark, CODE, NPWP , PPN, 0 AS DISC, '' AS keyno, " & _
                            "front AS [USER], '' AS SPK, invDATE AS TGL_REP, JOB, '' AS JOBCD, km AS FS_KM_AKHIR, " & _
                            "'' AS USEJOBDSC, invDATE AS RCDATE, '' AS KD_K, '' AS CR, '' AS CDK, " & _
                            "'1900-01-01' AS TGL_D, '' ASCDD, '' AS CRDTCD, FS_TH_MOBIL, '' AS USEFS_NM_HOBI, " & _
                            "FS_KD_PLG, FS_KD_WIL, '1900-01-01' AS dotam, CUSTOMER AS USENAME, LEFT(cust_adres,120) AS USEADDR, " & _
                            "KOTA AS USECITY, '' AS USEPOS, '' AS USEPSTN, '' AS USEGSM, '' AS USECDMA, '' AS USET4LAHIR, " & _
                            "'1900-01-01' AS USETGLLAHIR, '' AS ASURANSI, '' AS ASR_ADDR, '' AS ASR_CITY, " & _
                            "'1900-01-01' AS ASR_TGLM, '1900-01-01' AS ASR_TGLS,'0' AS  ASR_LIMIT, '' AS ASR_NOTE, " & _
                            "'' AS NOTE, '' AS GSM, '' AS CDMA, '' AS USEAGAMA, FRONT AS USERADD, 0 AS FLAGKM, " & _
                            "0 AS HEAVYUSE, '' AS pkp, '' AS KDPLG_ASR, '' AS USEEMAIL  " & _
                            "from [" & dbReplikasiHO & "]..T_invoiceNew WITH(NOLOCK) WHERE TPSVC <> 'W' and THS = 0 and PDS = 0 and NO_CHASSIS = '" & noRgk & "' and Right(NOINV, 1) <> 'B' Order By InvDate Desc ")
        GoTo tampilkanData
    End If
If Log.LogHOConn = 1 Then
    If ConHO = True Then
        Set rsHlp = objO.FetchDataHO("SELECT TOP 1 CUSTOMER,LEFT(cust_adres,120) as CUST_ADRES,KOTA, '' AS kodepos, TELPON, NO_POLISI, " & _
                            "NO_ENGINE, '' AS wmi, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, FS_KD_DEALER, " & _
                            "FS_NM_DEALER, DDATE, '' AS remark, CODE, NPWP , PPN, 0 AS DISC, '' AS keyno, " & _
                            "front AS [USER], '' AS SPK, invDATE AS TGL_REP, JOB, '' AS JOBCD, km AS FS_KM_AKHIR, " & _
                            "'' AS USEJOBDSC, invDATE AS RCDATE, '' AS KD_K, '' AS CR, '' AS CDK, " & _
                            "'1900-01-01' AS TGL_D, '' ASCDD, '' AS CRDTCD, FS_TH_MOBIL, '' AS USEFS_NM_HOBI, " & _
                            "FS_KD_PLG, FS_KD_WIL, '1900-01-01' AS dotam, CUSTOMER AS USENAME, LEFT(cust_adres,120) AS USEADDR, " & _
                            "KOTA AS USECITY, '' AS USEPOS, '' AS USEPSTN, '' AS USEGSM, '' AS USECDMA, '' AS USET4LAHIR, " & _
                            "'1900-01-01' AS USETGLLAHIR, '' AS ASURANSI, '' AS ASR_ADDR, '' AS ASR_CITY, " & _
                            "'1900-01-01' AS ASR_TGLM, '1900-01-01' AS ASR_TGLS,'0' AS  ASR_LIMIT, '' AS ASR_NOTE, " & _
                            "'' AS NOTE, '' AS GSM, '' AS CDMA, '' AS USEAGAMA, FRONT AS USERADD, 0 AS FLAGKM, " & _
                            "0 AS HEAVYUSE, '' AS pkp, '' AS KDPLG_ASR, '' AS USEEMAIL  " & _
                            "from T_invoiceNew WITH(NOLOCK) WHERE TPSVC <> 'W' and THS = 0 and PDS = 0 and NO_CHASSIS = '" & noRgk & "' and Right(NOINV, 1) <> 'B' Order By InvDate Desc ")

tampilkanData:
        If Not rsHlp.EOF Then
            
            Text(2) = rsHlp("FS_KD_PLG")
            Text(4) = Left(rsHlp("CUSTOMER"), 60)
            Text(5) = Left(Trim(rsHlp("CUST_ADRES")), 60)
            If Len(Trim(rsHlp("CUST_ADRES"))) > 60 Then
                Text(56) = Mid(NulSID(rsHlp("CUST_ADRES")), 61, Len(rsHlp("CUST_ADRES")) - 60)
            End If
            
            Text(6) = rsHlp("KOTA")
            Text(7) = rsHlp("KODEPOS")
            Text(8) = rsHlp("TELPON")
            Text(9) = rsHlp("GSM")
            Text(10) = rsHlp("CDMA")
            
            Text(11) = rsHlp("NPWP")
            Text(12) = rsHlp("USEJOBDSC")
            Text(13) = rsHlp("USEAGAMA")
            Text(14) = rsHlp("USEFS_NM_HOBI")
            
            '' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            '' Text(15) = rsHlp("FS_KD_DEALER")
            '' Text(16) = rsHlp("FS_NM_DEALER")
            '' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            
            
            ''-------------------------------------------------
            ''2016-02-01 | diganti dengan nilai dari salesunit
            ''-------------------------------------------------
            If Log.RPLHistori = 0 Then
                Set rsTemp = objO.FetchData("SELECT * FROM SALESUNIT WHERE NORANGKA = '" & Trim(noRgk) & "' ORDER BY tglentry desc")
            Else
                Set rsTemp = objO.FetchData("select ISNULL(b.fs_kd_dealer,'') as FS_KD_DEALER,ISNULL(b.fs_nm_dealer,'') as FS_NM_DEALER from [" & dbReplikasiHO & "]..Summary a with (nolock) left join [" & dbReplikasiHO & "]..t_dealer b with (nolock) on a.KodeDealer = b.kodecabang where a.NoRangka = '" & Trim(noRgk) & "'")
            End If
            Text(15) = ""
            Text(16) = ""
            If Not rsTemp.EOF Then
                rsTemp.MoveFirst
                Text(15) = rsTemp("FS_KD_DEALER")
                Text(16) = rsTemp("FS_NM_DEALER")
            End If
            Text(17) = rsHlp("CODE")
            Text(18) = TglDDMMYYYY(NulTgl(rsHlp("DDATE")))
            Text(19) = TglDDMMYYYY(NulTgl(rsHlp("DOTAM")))
            Text(20) = rsHlp("FS_TH_MOBIL")
            Text(21) = rsHlp("MDL")
            Text(22) = rsHlp("MODEL")
            Text(23) = rsHlp("FMDL")
            Text(24) = rsHlp("COLOR")
            If Len(rsHlp("NO_POLISI")) > 0 Then
                Dim nPol1 As String
                nPol1 = rsHlp("NO_POLISI")
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""
                Call pecahNoPol(nPol1, nDpn1, nTgh1, nBlk1)
                Text(25) = Trim(nDpn1)
                Text(46) = Trim(nTgh1)
                Text(47) = Trim(nBlk1)
            Else
                Text(25) = ""
                Text(46) = ""
                Text(47) = ""
            End If
            
            Text(26) = rsHlp("NO_ENGINE")
            Text(27) = rsHlp("KEYNO")
            Text(28) = rsHlp("PPN")
            Text(29) = rsHlp("DISC")
            Text(30) = rsHlp("USENAME")
            Text(31) = Left(NulSID(rsHlp("USEADDR")), 60)
            Text(57) = ""
            Text(58) = ""
            If Len(NulSID(rsHlp("USEADDR"))) > 60 Then
                Text(57) = Mid(NulSID(rsHlp("USEADDR")), 61, Len(NulSID(rsHlp("USEADDR"))) - 60)
                If Len(NulSID(rsHlp("USEADDR"))) > 120 Then
                    Text(58) = Mid(NulSID(rsHlp("USEADDR")), 121, Len(NulSID(rsHlp("USEADDR"))) - 120)
                End If
            End If
            Text(32) = rsHlp("USECITY")
            Text(33) = rsHlp("USEPOS")
            Text(34) = rsHlp("USEPSTN")
            Text(35) = rsHlp("USEGSM")
            Text(36) = rsHlp("USECDMA")
            If rsHlp("ASURANSI") = "" Then
                chAsuransi.value = Unchecked
            Else
                chAsuransi.value = Checked
            End If
            Text(37) = rsHlp("ASURANSI")
            Text(38) = rsHlp("ASR_ADDR")
            Text(39) = rsHlp("ASR_CITY")
            Text(40) = TglDDMMYYYY(rsHlp("ASR_TGLM"))
            Text(41) = TglDDMMYYYY(rsHlp("ASR_TGLS"))
            Text(42) = rsHlp("ASR_LIMIT")
            Text(43) = rsHlp("ASR_NOTE")
            Text(44) = rsHlp("NOTE")
            Text(45) = rsHlp("FS_KM_AKHIR")
            Text(51) = rsHlp("PKP")
            Text(59) = NulSID(Trim(rsHlp("USEEMAIL")))
            
            ' INPUT HISTORY SERVICE
            'Call panggilHistory(noRgk, nPol1)
            cekHistoryHO = True
        Else
            cekHistoryHO = False
        End If
    End If
End If
Exit Function
vbError:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekHistoryHO of Form frmSIDCustomer")
    cekHistoryHO = False
    Exit Function
End Function

Private Sub panggilHistory(noRgk As String, noPols As String)
1         On Error GoTo vbError
          Dim rsHist As ADODB.Recordset
              
          ' AMBIL HISTORY DARI HO
          Dim rsHlp2 As ADODB.Recordset
          Dim dBaseHO As String
          
2     If Log.LogHOConn = 1 Then
3         If ConHO = True Then
4             dBaseHO = HODb
5             Set rsHlp = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
6             If Not rsHlp.EOF Then
7                 Do While Not rsHlp.EOF
8                     If rsHlp("NO_CHASSIS") <> "" Then
                          ' INPUT HISTORY
9                         Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "' AND KM= '" & rsHlp("KM") & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC, NO_WO DESC")
10                        If rsHlp2.EOF Then
11                            Call objO.InsertData("INSERT INTO HISTORY VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & Format(rsHlp("DATE"), "mm/dd/yyyy") & "'," & rsHlp("PART_G") & ", " _
                                           & "" & rsHlp("PART_NG") & "," & rsHlp("OIL_G") & "," & rsHlp("OIL_NG") & "," & rsHlp("MTRL") & "," & rsHlp("LABOR") & "," & rsHlp("BODY") & "," & rsHlp("SBLT") & "," & rsHlp("MET") & "," & rsHlp("TOTAL") & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "'," _
                                           & "'" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "'," _
                                           & "'" & rsHlp("MEKANIK4") & "','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "'," _
                                           & "'" & rsHlp("KLPKERJA7") & "','" & rsHlp("KLPKERJA8") & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & rsHlp("FN_RETURN") & "','" & rsHlp("FS_RC") & "'," _
                                           & "'" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "','" & rsHlp("FS_KD_WIL") & "','" & rsHlp("THS") & "','" & rsHlp("MOU") & "','" & rsHlp("PDS") & "','" & rsHlp("FS_LOG_DEL") & "','" & rsHlp("FD_TGL_DEL") & "', " _
                                           & "'" & rsHlp("ALASAN_DEL") & "', '','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')")
12                        End If
                          
                          ' INPUT PART
                          Dim stringAwal As String
                          Dim stringPotong1 As String
                          Dim stringPotong2 As String
                          Dim nNext As Integer
                          Dim hasil(15) As String
13                        If rsHlp("MEMOPART") <> "" Then
14                            Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "'")
15                            If rsHlp2.EOF Then
16                                stringAwal = Trim(rsHlp("MEMOPART"))
17                                Do While Len(stringAwal) > 0
18                                    stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
19                                    nNext = 0
20                                    Do While Len(stringPotong1) > 0
21                                        stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
22                                        hasil(nNext) = Trim(stringPotong2)
23                                        nNext = nNext + 1
24                                        stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
25                                    Loop
26                                    stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
                                      
                                      ' SIMPAN
27                                    Call objO.InsertData("INSERT INTO T_PART VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & IIf(Trim(hasil(5)), 1, 0) & "','" & Trim(hasil(6)) & "', " _
                                                   & "'" & Trim(hasil(7)) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "')")
28                                Loop
29                            End If
30                        End If
                          
                          ' INPUT SUBLET
31                        If rsHlp("MEMOSUBLET") <> "" Then
32                            Set rsHlp2 = objO.FetchData("SELECT * FROM HIST_SUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "' AND NO_CHASSIS = '" & noRgk & "'")
33                            If rsHlp2.EOF Then
34                                stringAwal = Trim(rsHlp("MEMOSUBLET"))
35                                Do While Len(stringAwal) > 0
36                                    stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
37                                    nNext = 0
38                                    Do While Len(stringPotong1) > 0
39                                        stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
40                                        hasil(nNext) = Trim(stringPotong2)
41                                        nNext = nNext + 1
42                                        stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
43                                    Loop
44                                    stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
                                      
                                      ' SIMPAN
45                                    Call objO.InsertData("INSERT INTO HIST_SUBLET VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & Trim(hasil(5)) & "','" & Trim(hasil(6)) & "', " _
                                                   & "'" & Trim(hasil(7)) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "','" & Trim(hasil(10)) & "','" & Trim(hasil(11)) & "','" & Trim(hasil(12)) & "','" & Trim(hasil(13)) & "','" & Trim(hasil(14)) & "')")
46                                Loop
47                            End If
48                        End If
49                    End If
50                    rsHlp.MoveNext
51                Loop
52            End If
53        End If
54    End If
55    Exit Sub
vbError:
56    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilHistory of Form frmSIDCustomer")
57        Exit Sub
End Sub

Private Function cekBengkelLama(noRgk As String) As Boolean
      Dim rsHlp As ADODB.Recordset
1        On Error GoTo cekBengkelLama_Error

2             Set rsHlp = objO.FetchDataAAS("SELECT CUSTOMER,LEFT(cust_adres,120) as CUST_ADRES,KOTA,KODEPOS,TELPON,NO_POLISI,NO_ENGINE,WMI,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,REMARK,CODE,NPWP,PPN,DISC,KEYNO,[USER],SPK,TGL_REP,JOB,JOBCD,FS_KM_AKHIR,USEJOBDSC,RCDATE,KD_K,CR,CDK,TGL_D,CDD,CRDTCD,FS_TH_MOBIL,USEFS_NM_HOBI,FS_KD_PLG,FS_KD_WIL,DOTAM,USENAME,USEADDR,USECITY,USEPOS,USEPSTN,USEGSM,USECDMA,USET4LAHIR,USETGLLAHIR,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,NOTE,GSM,CDMA,USEAGAMA,USERADD,FLAGKM,HEAVYUSE,PKP,KDPLG_ASR,USEEMAIL,KECAMATAN,USEKECAMATAN FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & noRgk & "'")
3             If Not rsHlp.EOF Then
                  
4                 Text(2) = rsHlp("FS_KD_PLG")
5                 Text(4) = Left(rsHlp("CUSTOMER"), 60)
6                 Text(5) = Left(rsHlp("CUST_ADRES"), 60)
7                 If Len(rsHlp("CUST_ADRES")) > 60 Then
8                     Text(56) = Mid(NulSID(rsHlp("CUST_ADRES")), 61, Len(rsHlp("CUST_ADRES")) - 60)
9                 End If
                  
10                Text(6) = rsHlp("KOTA")
11                Text(7) = rsHlp("KODEPOS")
12                Text(8) = rsHlp("TELPON")
13                Text(9) = ""
14                Text(10) = ""
                  
15                Text(11) = rsHlp("NPWP")
16                Text(12) = ""
17                Text(13) = ""
18                Text(14) = rsHlp("FS_NM_HOBI")
19                Text(15) = rsHlp("FS_KD_DEALER")
20                Text(16) = rsHlp("FS_NM_DEALER")
21                Text(17) = rsHlp("CODE")
22                Text(18) = TglDDMMYYYY(NulTgl(Replace(Replace(rsHlp("DDATE"), ">", "0"), " ", 0)))
23                Text(19) = "01-01-1900"
24                Text(20) = rsHlp("FS_TH_MOBIL")
25                Text(21) = rsHlp("MDL")
26                Text(22) = rsHlp("MODEL")
27                Text(23) = rsHlp("FMDL")
28                Text(24) = rsHlp("COLOR")
29                If Len(rsHlp("NO_POLISI")) > 0 Then
                      Dim nPol1 As String
30                    nPol1 = rsHlp("NO_POLISI")
                      Dim nDpn1 As String
                      Dim nTgh1 As String
                      Dim nBlk1 As String
31                    nDpn1 = ""
32                    nTgh1 = ""
33                    nBlk1 = ""
34                    Call pecahNoPol(nPol1, nDpn1, nTgh1, nBlk1)
35                    Text(25) = Trim(nDpn1)
36                    Text(46) = Trim(nTgh1)
37                    Text(47) = Trim(nBlk1)
38                Else
39                    Text(25) = ""
40                    Text(46) = ""
41                    Text(47) = ""
42                End If
                  
43                Text(26) = rsHlp("NO_ENGINE")
44                Text(27) = rsHlp("KEYNO")
45                Text(28) = rsHlp("PPN")
46                Text(29) = rsHlp("DISC")
47                Text(30) = rsHlp("CUSTOMER")
48                Text(31) = Left(NulSID(rsHlp("CUST_ADRES")), 60)
49                Text(57) = ""
50                Text(58) = ""
51                If Len(NulSID(rsHlp("CUST_ADRES"))) > 60 Then
52                    Text(57) = Mid(NulSID(rsHlp("CUST_ADRES")), 61, Len(NulSID(rsHlp("CUST_ADRES"))) - 60)
53                    If Len(NulSID(rsHlp("CUST_ADRES"))) > 120 Then
54                        Text(58) = Mid(NulSID(rsHlp("CUST_ADRES")), 121, Len(NulSID(rsHlp("CUST_ADRES"))) - 120)
55                    End If
56                End If
                  
57                Text(32) = rsHlp("KOTA")
58                Text(33) = rsHlp("KODEPOS")
59                Text(34) = rsHlp("TELPON")
60                Text(35) = ""
61                Text(36) = ""
62                chAsuransi.value = Unchecked
63                Text(37) = ""
64                Text(38) = ""
65                Text(39) = ""
66                Text(40) = "01-01-1900"
67                Text(41) = "01-01-1900"
68                Text(42) = "0"
69                Text(43) = ""
70                Text(44) = ""
71                Text(45) = rsHlp("FS_KM_AKHIR")
72                Text(51) = ""
                  
73                cekBengkelLama = True
74            Else
75                cekBengkelLama = False
76            End If

77       On Error GoTo 0
78       Exit Function

cekBengkelLama_Error:

79        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekBengkelLama of Form frmSIDCustomer")
End Function

Private Function cekUnitCustomer(noRangka As String) As Boolean
          Dim rsHlp As ADODB.Recordset
          Dim rsHlp2 As ADODB.Recordset
            
1        On Error GoTo cekUnitCustomer_Error
          
2             Set rsHlp = objO.FetchData("SELECT * FROM [SINGLEIDORI]..UNITCUSTOMER WHERE norangka = '" & noRangka & "'")
3             If Not rsHlp.EOF Then
4                 Text(15) = NulSID(rsHlp!KdDlrPenjual)
5                 Text(16) = NulSID(rsHlp!NmDlrPenjual)
                  
6                 Text(18) = Format(NulSID(rsHlp!dDate), "DD-MM-YYYY")
7                 Text(18) = Format(NulSID(rsHlp!dDate), "DD-MM-YYYY")
                  
8                 Text(19) = "01-01-3000"
9                 Text(20) = NulSID(rsHlp!thMobil)
                  
10                Text(21) = NulSID(rsHlp!MDL)
                  'Call text_Validate(21, False)
                  '20180221 agar tidak muncul show modaly
                  
11                If Log.RPLHistori = 0 Then
12                    Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(21) & "'")
13                Else
14                    Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(21) & "'")
15                End If
16                If Not rsSearch.EOF Then
17                    Text(21) = rsSearch("MDL")
18                    Text(23) = rsSearch("Desc")
19                    Text(22) = rsSearch("MODEL")
20                    gGroupModel = rsSearch("FS_GRP_MDL")
21                End If
22                Text(24) = NulSID(rsHlp!Color)
                  
23                Text(26) = NulSID(rsHlp!NoEngine)
24                Text(27) = NulSID(rsHlp!nokey)
25                If NulSID(rsHlp!idcustservice) <> "" Then
26                    Text(2) = NulSID(rsHlp!idcustservice)
27                ElseIf NulSID(rsHlp!idcustsales) <> "" Then
28                    Text(2) = NulSID(rsHlp!idcustsales)
29                End If
30                cekUnitCustomer = True
31            Else
32                cekUnitCustomer = False
33            End If

34       On Error GoTo 0
35       Exit Function

cekUnitCustomer_Error:

36        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cekUnitCustomer of Form frmSIDCustomer"
End Function

Private Function cekSPK(noRangka As String) As Boolean
      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset
1        On Error GoTo cekSPK_Error

2             Set rsHlp = objO.FetchDataSPK("select (case when a.alm1nama=1 then a.nama_faktur+' '+a.alamat1_faktur else a.nama_faktur end) as NamaFaktur, " _
                                         & "(case when a.alm1nama=1 then a.alamat2_faktur+' '+a.alamat3_faktur+' '+a.alamat4_faktur+' '+a.alamat5_faktur else a.alamat1_faktur+' '+a.alamat2_faktur+' '+a.alamat3_faktur+' '+a.alamat4_faktur+' '+a.alamat5_faktur end) as AlamatFaktur, " _
                                         & "(case when a.kotamadya='LAINNYA' then a.wil_lain else a.kotamadya end) as Kota, " _
                                         & "b.kodepos,b.telp ,b.gsmfaktur,b.cdmafaktur,b.agama,b.hoby,ISNULL(b.tglcr1,'01/01/3000') AS TGLCR1,ISNULL(c.tgldo,'01/01/3000') AS TGLDO,ISNULL(b.tahun,0) AS TAHUN,c.kodekategori,d.namakategory,e.namaind,b.nopol,c.nomesin,c.[key], " _
                                         & "a.snama , a.salamat, a.skota, a.skodepos, a.stelp, a.sgsm, a.scdma, a.kode_lgn from spk a WITH(NOLOCK) left join afi b WITH(NOLOCK) on a.spk_no=b.spk_no " _
                                         & "left join inmst c WITH(NOLOCK) on a.norangka=c.kodeitem left join inkategory d WITH(NOLOCK) on c.kodekategori=d.kodekategory left join inwarnaext e WITH(NOLOCK) on c.kodewarna=e.kodewarna " _
                                         & "where a.norangka = '" & noRangka & "'")
3             If Not rsHlp.EOF Then
              
4                 Text(2) = NulSID(rsHlp("KODE_LGN"))
5                 Text(3) = Left(Trim(NulSID(rsHlp("KODE_LGN"))), 3)
6                 Text(4) = Left(Replace(NulSID(rsHlp("NAMAFAKTUR")), vbNewLine, ""), 60)
7                 Text(5) = Left(Replace(NulSID(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 60)
8                 If Len(Replace(Trim(NulSID(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) > 60 Then
9                     Text(56) = Mid(Replace(NulSID(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 61, Len(Replace(NulSID(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) - 60)
10                End If
11                Text(6) = NulSID(rsHlp("KOTA"))
12                Text(7) = NulSID(rsHlp("KODEPOS"))
13                Text(8) = NulSID(rsHlp("TELP"))
14                Text(9) = NulSID(rsHlp("GSMFAKTUR"))
15                Text(10) = NulSID(rsHlp("CDMAFAKTUR"))
                  
16                Text(15) = gKdCbg
17                Text(16) = gNmCbg
              
18                Text(18) = TglDDMMYYYY(NulTgl(rsHlp("TGLCR1")))
19                Text(19) = TglDDMMYYYY(NulTgl(rsHlp("TGLDO")))
20                Text(20) = CDbl(IIf(rsHlp("TAHUN") = "", 0, rsHlp("TAHUN")))
21                Text(21) = NulSID(rsHlp("KODEKATEGORI"))
                  
22                If Log.RPLHistori = 0 Then
23                    Text(22) = objO.FetchData_String("SELECT MODEL FROM MODEL WITH(NOLOCK) WHERE MDL = '" & NulSID(rsHlp("KODEKATEGORI")) & "'")
24                Else
25                    Text(22) = objO.FetchData_String("SELECT MODEL FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & NulSID(rsHlp("KODEKATEGORI")) & "'")
26                End If
27                Text(23) = NulSID(rsHlp("NAMAKATEGORY"))
28                Text(24) = NulSID(rsHlp("NAMAIND"))
29                nPol = Trim(NulSID(rsHlp("NOPOL")))
                  Dim nDpn1 As String
                  Dim nTgh1 As String
                  Dim nBlk1 As String
30                nDpn1 = ""
31                nTgh1 = ""
32                nBlk1 = ""
33                If nPol <> "" Then
34                    Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
35                    Text(25) = Trim(nDpn1)
36                    Text(46) = Trim(nTgh1)
37                    Text(47) = Trim(nBlk1)
38                End If
              
39                Text(26) = NulSID(rsHlp("NOMESIN"))
40                Text(27) = NulSID(rsHlp("KEY"))
                  
41                Text(30) = Left(Replace(NulSID(rsHlp("SNAMA")), vbNewLine, ""), 60)
42                Text(31) = Left(Replace(NulSID(Trim(rsHlp("SALAMAT"))), vbNewLine, ""), 60)
43                Text(57) = ""
44                Text(58) = ""
45                If Len(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, "")) > 60 Then
46                    Text(57) = Mid(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, ""), 61, Len(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, "")) - 60)
47                    If Len(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, "")) > 120 Then
48                        Text(58) = Mid(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, ""), 121, Len(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, "")) - 120)
49                    End If
50                End If
                  
51                Text(32) = NulSID(rsHlp("SKOTA"))
52                Text(33) = NulSID(rsHlp("SKODEPOS"))
53                Text(34) = NulSID(rsHlp("STELP"))
54                Text(35) = NulSID(rsHlp("SGSM"))
55                Text(36) = NulSID(rsHlp("SCDMA"))
                  
56                sNamaSpk = NulSID(rsHlp("SNAMA"))
57                sAlamatSpk = NulSID(rsHlp("SALAMAT"))
58                sKotaSpk = NulSID(rsHlp("SKOTA"))
59                sKodeposSpk = NulSID(rsHlp("SKODEPOS"))
60                sTelpSpk = NulSID(rsHlp("STELP"))
61                sGsmSpk = NulSID(rsHlp("SGSM"))
62                sCdmaSpk = NulSID(rsHlp("SCDMA"))
                  
63                chAsuransi.value = Unchecked
64                cekSPK = True
65            Else
66                cekSPK = False
67            End If

68       On Error GoTo 0
69       Exit Function

cekSPK_Error:

70        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekSPK of Form frmSIDCustomer")
End Function

Private Function cekSPKHO(noRangka As String) As Boolean
1        On Error GoTo cekSPKHO_Error

2     If Log.LogHOConn = 1 Then
      Dim rsHlp As ADODB.Recordset
3             Set rsHlp = objO.FetchDataSPKHO("SELECT NAMAFAKTUR, ALAMATFAKTUR, KOTA, KODEPOS, TELP, GSMFAKTUR, CDMAFAKTUR, AGAMA, HOBY, TGLCR1, TGLDO, TAHUN, " _
                                           & "a.KODEKATEGORI , NAMAKATEGORY, NAMAIND, noPol, A.NOMESIN, [KEYNO], SNAMA, SALAMAT, SKOTA, SKODEPOS, STELP, SGSM, SCDMA, a.dealer " _
                                           & "FROM SPK99..DATACUST A WITH(NOLOCK) INNER JOIN SPK00..INKATEGORY B WITH(NOLOCK) ON A.KODEKATEGORI = B.KODEKATEGORY " _
                                           & "INNER JOIN SPK00..INWARNAEXT C WITH(NOLOCK) ON A.KODEWARNA = C.KODEWARNA INNER JOIN SPK00..INMST D WITH(NOLOCK) ON A.NORANGKA = D.KODEITEM " _
                                           & "WHERE NORANGKA = '" & noRangka & "'")
4             If Not rsHlp.EOF Then
5                 Text(4) = Left(NulSID(rsHlp("NAMAFAKTUR")), 60)
6                 Text(5) = NulSID(Left(Replace(Trim(NulSID(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 60))
7                 If Len(Replace(Trim(NulSID(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) > 60 Then
8                     Text(56) = Mid(Replace(NulSID(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 61, Len(Replace(Trim(NulSID(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) - 60)
9                 End If
10                Text(6) = NulSID(rsHlp("KOTA"))
11                Text(7) = NulSID(rsHlp("KODEPOS"))
12                Text(8) = NulSID(rsHlp("TELP"))
13                Text(9) = NulSID(rsHlp("GSMFAKTUR"))
14                Text(10) = NulSID(rsHlp("CDMAFAKTUR"))
15                Text(15) = cariDealer(Trim(rsHlp("DEALER")), "kode")
16                Text(16) = cariDealer(Trim(rsHlp("DEALER")), "nama")
17                Text(18) = TglDDMMYYYY(NulTgl(rsHlp("TGLCR1")))
18                Text(19) = TglDDMMYYYY(NulTgl(rsHlp("TGLDO")))
19                Text(20) = CDbl(NulA(rsHlp("TAHUN")))
20                Text(21) = NulSID(rsHlp("KODEKATEGORI"))
                  
21                If Log.RPLHistori = 0 Then
22                    Text(22) = objO.FetchData_String("SELECT MODEL FROM MODEL WITH(NOLOCK) WHERE MDL = '" & NulSID(rsHlp("KODEKATEGORI")) & "'")
23                Else
24                    Text(22) = objO.FetchData_String("SELECT MODEL FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & NulSID(rsHlp("KODEKATEGORI")) & "'")
25                End If
26                Text(23) = NulSID(rsHlp("NAMAKATEGORY"))
27                Text(24) = NulSID(rsHlp("NAMAIND"))
28                nPol = Trim(NulSID(rsHlp("NOPOL")))
                  Dim nDpn1 As String
                  Dim nTgh1 As String
                  Dim nBlk1 As String
29                nDpn1 = ""
30                nTgh1 = ""
31                nBlk1 = ""
32                If nPol <> "" Then
33                    Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
34                    Text(25) = Trim(nDpn1)
35                    Text(46) = Trim(nTgh1)
36                    Text(47) = Trim(nBlk1)
37                End If
                              
38                Text(26) = NulSID(rsHlp("NOMESIN"))
39                Text(27) = NulSID(rsHlp("KEYNO"))
40                Text(30) = NulSID(rsHlp("SNAMA"))
41                Text(31) = Left(Replace(NulSID(Trim(rsHlp("SALAMAT"))), vbNewLine, ""), 60)
42                Text(57) = ""
43                Text(58) = ""
44                If Len(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, "")) > 60 Then
45                    Text(57) = Mid(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, ""), 61, Len(Replace(NulSID(rsHlp("SALAMAT")), vbNewLine, "")) - 60)
46                    If Len(Replace(NulSID(Trim(rsHlp("SALAMAT"))), vbNewLine, "")) > 120 Then
47                        Text(58) = Mid(Replace(NulSID(Trim(rsHlp("SALAMAT"))), vbNewLine, ""), 121, Len(Replace(NulSID(Trim(rsHlp("SALAMAT"))), vbNewLine, "")) - 120)
48                    End If
49                End If
50                Text(32) = NulSID(rsHlp("SKOTA"))
51                Text(33) = NulSID(rsHlp("SKODEPOS"))
52                Text(34) = NulSID(rsHlp("STELP"))
53                Text(35) = NulSID(rsHlp("SGSM"))
54                Text(36) = NulSID(rsHlp("SCDMA"))
55                chAsuransi.value = Unchecked
56                cekSPKHO = True
57            Else
58                cekSPKHO = False
59            End If
60        End If

61       On Error GoTo 0
62       Exit Function

cekSPKHO_Error:

63        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekSPKHO of Form frmSIDCustomer")
End Function

Private Function cariDealer(IniDealer As String, tipe As String) As String
1        On Error GoTo cariDealer_Error

2         Select Case tipe
              Case "kode"
3                 cariDealer = objO.FetchData_String("SELECT FS_KD_DEALER FROM T_DEALER WITH(NOLOCK) WHERE KODECABANG = '" & Trim(IniDealer) & "'")
4             Case "nama"
5                 cariDealer = objO.FetchData_String("SELECT FS_NM_DEALER FROM T_DEALER WITH(NOLOCK) WHERE KODECABANG = '" & Trim(IniDealer) & "'")
6         End Select

7        On Error GoTo 0
8        Exit Function

cariDealer_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariDealer of Form frmSIDCustomer")
End Function

Private Function cekHistoryOnlyPDS(noChs As String) As Boolean
1        On Error GoTo cekHistoryOnlyPDS_Error

2         cekHistoryOnlyPDS = False
3         Set rsSearch = objO.FetchData("SELECT PDS FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noChs) & "'")
4         If Not rsSearch.EOF Then
5             If rsSearch.RecordCount = 1 Then
6                 If Trim(rsSearch("PDS")) = "1" Then
7                     cekHistoryOnlyPDS = True
8                 End If
9             End If
10        End If

11       On Error GoTo 0
12       Exit Function

cekHistoryOnlyPDS_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekHistoryOnlyPDS of Form frmSIDCustomer")
End Function

Private Function cekHistoryPDS(noChs As String) As Boolean
1        On Error GoTo cekHistoryPDS_Error

2         cekHistoryPDS = objO.FetchData_Boolean("SELECT NO_CHASSIS FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noChs) & "' AND PDS = '1'")

3        On Error GoTo 0
4        Exit Function

cekHistoryPDS_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekHistoryPDS of Form frmSIDCustomer")
End Function

Private Function cekDataSingleID(table As String, where As String) As Boolean
    Dim rsHasil As ADODB.Recordset
    Set rsHasil = objO.FetchData("SELECT * FROM [SINGLEIDORI].." & table & " WITH (NOLOCK) " & where)
    If rsHasil.EOF Then
        cekDataSingleID = False
    Else
        cekDataSingleID = True
    End If
End Function

Private Sub tampilkanFormSearch()
    gMainKeyCari = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
     
    With frmSIDSearchCust
        .Text(1) = gMainKeyCari
        .Text(4) = NulSID(Text(2))
        gSearchOption = "frmSIDCustomer"
        .Left = Me.asxPowerBanner1.Width '+ 100
        .Top = Me.Top + 450
        .show vbModal
    End With
End Sub

Private Sub tampilkanFormSearchSurat()
    gMainKeyCari = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
     
    With frmSIDSearchCust
        .Text(1) = gMainKeyCari
        .Text(4) = NulSID(Text(2))
        gSearchOption = "frmSIDCustomerSurat"
        .Left = Me.asxPowerBanner1.Width '+ 100
        .Top = Me.Top + 450
        .show vbModal
    End With
End Sub

Private Sub hideComboProv(value As Boolean)
    cboProv.Visible = value
    cboKota.Visible = value
    cboKec.Visible = value
    cboKel.Visible = value
    
    FrAlm.Visible = Not value
    'cmdBack.Visible = Not value
End Sub

Private Sub hideComboProvS(value As Boolean)
    cboProvS.Visible = value
    cboKotaS.Visible = value
    cboKecS.Visible = value
    cboKelS.Visible = value
    
    frmAlamat2.Visible = Not value
    'cmdBackS.Visible = Not value
End Sub


Private Sub panggilCustomerSID()
          Dim rsCust As ADODB.Recordset
          Dim rsHelp As ADODB.Recordset
          Dim isNewID As Boolean
1         isNewID = False
2         If cekDataSingleID("NewCustomer", "Where idCust = '" & NulSID(Text(2)) & "'") = True Then
              ''jika id ada di new customer
3             If cekDataSingleID("bridging", "where KodeDealer = '" & Right(gDb, 2) & "' and idcust = '" & NulSID(Text(2)) & "'") = False Then
4                 Call objO.InsertData("INSERT INTO [SINGLEIDORI]..BRIDGING (KodeDealer,KodeLgn,IDCust,TglEntry,UserEntry) VALUES ('" & Right(gDb, 2) & "','" & NulSID(Text(2)) & "','" & NulSID(Text(2)) & "',GETDATE(),'" & gSA & "')")
5                 Text(65) = NulSID(Text(2))
6             Else
                  ''2018-01-23 || KELOLOSAN DATA TEXT(65) BERISI KOSONG
                  ''jika di bridging ada data , kodelgn lama diambil dari bridging
                  ''------------------------------------------------------------------
7                 Set rsCust = objO.FetchData("SELECT * FROM [SINGLEIDORI]..bridging WITH (NOLOCK) where KodeDealer = '" & Right(gDb, 2) & "' and idcust = '" & NulSID(Text(2)) & "'")
8                 If Not rsCust.EOF Then
9                     Text(65) = NulSID(rsCust!kodeLgn)
10                End If
11            End If
              
12            Set rsCust = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & IIf(NulSID(Text(65)) = "", Text(2), NulSID(Text(65))) & "'")
13            If rsCust.EOF Then
                  
14                Set rsCust = objO.FetchData("SELECT * FROM [SINGLEIDORI]..NewCustomer WITH (NOLOCK) WHERE IDCUST = '" & NulSID(Text(2)) & "'")
15                If Not rsCust.EOF Then
                      Dim kodeLgn As String
                      Dim namaLengkap As String, Alamat As String, provinsi As String, Kota As String, kec As String, kel As String, vStatus As String
16                    provinsi = "": Kota = "": kec = "": kel = ""
                      
17                    If UCase(rsCust!TypeCustomer) = UCase("Pemerintahan") Then
18                        vStatus = "P"
19                    Else
20                        vStatus = "S"
21                    End If
22                    Alamat = NulSID(rsCust!Alamat)
                       
23                    Alamat = Alamat & " Kel." & NulSID(rsCust!kelurahan) & " Kec." & NulSID(rsCust!kecamatan) & " " & NulSID(rsCust!Kota) & " Prov." & NulSID(rsCust!Propinsi)
24                    provinsi = UCase(NulSID(rsCust!Propinsi))
25                    Kota = UCase(NulSID(rsCust!Kota))
26                    kec = UCase(NulSID(rsCust!kecamatan))
27                    kel = UCase(NulSID(rsCust!kelurahan))
                           
28                    namaLengkap = NulSID(NulSID(rsCust!namacust1) & " " & NulSID(rsCust!namacust2) & " " & NulSID(rsCust!namacust3))
29                    kodeLgn = IIf(NulSID(Text(65)) = "", Text(2), NulSID(Text(65)))
                      
30                    sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3" _
                                 & " ,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry,HubungDengan,Telepon1)" _
                                & " VALUES ('" & kodeLgn & "','" & NulSID(namaLengkap) & "','" & Left(NulSID(kodeLgn), 3) & "'" _
                                & ",'" & Mid(NulSID(Alamat), 1, 40) & "','" & Mid(NulSID(Alamat), 41, 40) & "','" & Mid(NulSID(Alamat), 81, 40) & "','" & Replace(Replace(NulSID(rsCust!IdTelp & rsCust!telp) & "," & NulSID(rsCust!hp1) & "," & NulSID(rsCust!hp2), ">", ""), " ", "") & "','" & NulSID(rsCust!NoFax) & "'" _
                                & ",'" & NulSID(rsCust!noNPWP) & "','D0','D0','D0','D0','D0','" & vStatus & "','" & gSA & "','" & Format(isnow, "MM/DD/YYYY") & "','" & NulSID(rsCust!NamaPIC) & "','" & NulSID(rsCust!TelpPIC) & "')"
31                    Call objO.InsertDataACC(sSQL)
                       '' insert data armst
32                    sSQL = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                      "        StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,DBSumber,Keterangan,ClientName,KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak, SIUP, No_Mou ) "
33                    sSQL = sSQL & " VALUES ( '" & NulSID(kodeLgn) & "', " & _
                      "                  '" & s2d(NulSID(namaLengkap)) & "', " & _
                      "                  '" & s2d(Left(NulSID(kodeLgn), 3)) & "', " & _
                      "                  '" & s2d(Mid(NulSID(Alamat), 1, 40)) & "', " & _
                      "                  '" & s2d(Mid(NulSID(Alamat), 41, 40)) & "', " & _
                      "                  '" & s2d(Mid(NulSID(Alamat), 81, 40)) & "', " & _
                      "                  '" & s2d(NulSID(rsCust!hp1)) & "', " & _
                      "                  '" & s2d(NulSID(rsCust!NoFax)) & "', " & _
                      "                  '" & s2d(NulSID(NulSID(rsCust!noNPWP))) & "', " & _
                      "                  'D0', 0, " & _
                      "                  '" & s2d(NulSID(rsCust!NamaPIC)) & "', 'GRP', " & _
                      "                  '" & s2d(vStatus) & "', '', " & _
                      "                  " & "0" & ", 'Customer GRP', 0, " & _
                      "                  '" & gSA & "', GETDATE(),0,0,0, " & _
                      "                  '" & Trim(gDb) & "','Data Baru','" & CompName & "' ,'" & gSA & "',Getdate(), " & _
                      "                  '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', " & _
                      "                  " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                      "                  " & 0 & ", '" & 0 & "', '', '', '' ) "   'A.Rev1711291-SingleId
34                    Call objO.InsertDataACC(sSQL)
35                End If
36            End If
37            isNewID = True
38            GoTo Lanjutkan
39        Else
              ''jika id tidak ada di new customer dicari di bridging
40            If MsgBox("Idcust tidak ditemukan apakah ingin melanjutkan pencarian?", vbYesNo) = vbYes Then
41                If NulSID(Text(65)) = "" Then
42                    Call tampilkanFormSearch
43                    Exit Sub
44                Else
45                    If cekDataSingleID("bridging", "where KodeDealer = '" & Right(gDb, 2) & "' and KodeLgn = '" & NulSID(txt(65)) & "'") = True Then
46                        GoTo Lanjutkan
47                    Else
48                        Call tampilkanFormSearch
49                        Exit Sub
50                    End If
51                End If
52            Else
53                Text(2).Text = ""
54                Text(3).Text = ""
55                Text(65).Text = ""
56                Exit Sub
57            End If
              
58        End If
Lanjutkan:
59        If isNewID = True Then
      '        sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IdCust = '" & nulsid(Text(2)) & "'"
60            sSQL = "select b.*,a.kodelgn"
61            sSQL = sSQL & " from [SingleIDOri]..Bridging a WITH (NOLOCK) "
62            sSQL = sSQL & " inner join [SingleIDOri]..NewCustomer b WITH (NOLOCK) on a.IDCust = b.IDCust"
63            sSQL = sSQL & " where a.KodeDealer = '" & Right(gDb, 2) & "' and a.idcust = '" & NulSID(Text(2)) & "'"
64        Else
65            sSQL = "select b.*,a.kodelgn"
66            sSQL = sSQL & " from [SingleIDOri]..Bridging a WITH (NOLOCK) "
67            sSQL = sSQL & " inner join [SingleIDOri]..NewCustomer b WITH (NOLOCK) on a.IDCust = b.IDCust"
68            sSQL = sSQL & " where a.KodeDealer = '" & Right(gDb, 2) & "' and a.KodeLgn = '" & NulSID(txt(65)) & "'"
              
69        End If
70        Set rsCust = objO.FetchData(sSQL)
          
71        If rsCust.EOF Then
              'jika data no rangka kosong maka ditampilkan pencarian single id
72            Call tampilkanFormSearch
73            Exit Sub
74        End If
          
          'Call loadComboBox

75        Text(2) = NulSID(rsCust!idcust)
76        Text(3) = Left(NulSID(rsCust!idcust), 3) '"XXX"
77        Text(3).Enabled = False
78        Text(65) = NulSID(rsCust!kodeLgn)
79        Text(4) = Left(NulSID(rsCust!namacust1), 31)
80        Text(4).Enabled = False
81        Text(60) = Left(NulSID(rsCust!namacust2), 31)
82        Text(60).Enabled = False
83        Text(61) = Left(NulSID(rsCust!namacust3), 31)
84        Text(61).Enabled = False
          
85        Text(7) = NulSID(rsCust!KodePos)
86        Text(7).Enabled = False
87        Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(rsCust!Propinsi) & "'")
      '    If Not rsHelp.EOF Then
      '        cboProv.Text = UCase(nulsid(rsCust!Propinsi))
      '        Call cboProv_Change
      '        cboProv.Enabled = False
      '
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT Kota FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE kota = '" & nulsid(rsCust!Kota) & "'")
      '        If Not rsHelp.EOF Then
      '        cboKota.Text = UCase(nulsid(rsCust!Kota))
      '        End If
      '        Call cboKota_Change
      '        cboKota.Enabled = False
      '
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT Kecamatan FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE kota = '" & nulsid(rsCust!Kota) & "' and Kecamatan = '" & nulsid(rsCust!kecamatan) & "'")
      '        If Not rsHelp.EOF Then
      '        cboKec.Text = UCase(nulsid(rsCust!kecamatan))
      '        End If
      '        Call cboKec_Change
      '        cboKec.Enabled = False
      '
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT Kelurahan FROM [SingleIdOri]..TopologyDaerah WITH (NOLOCK) WHERE Kecamatan = '" & nulsid(rsCust!kecamatan) & "' and Kelurahan = '" & nulsid(rsCust!kelurahan) & "'")
      '        If Not rsHelp.EOF Then
      '        cboKel.Text = UCase(nulsid(rsCust!kelurahan))
      '        End If
      '        Call cboKel_Change
      '        cboKel.Enabled = False
      '    Else
88            Call hideComboProv(False)
89            txt(63) = UCase(NulSID(rsCust!Propinsi))
90            txt(63).Enabled = False
91            txt(64) = UCase(NulSID(rsCust!Kota))
92            txt(64).Enabled = False
93            txt(65) = UCase(NulSID(rsCust!kecamatan))
94            txt(65).Enabled = False
95            txt(66) = UCase(NulSID(rsCust!kelurahan))
96            txt(66).Enabled = False
      '    End If
           
97        txt(8) = NulSID(rsCust!Alamat)
98        txt(8).Enabled = False
99        txt(9) = NulSID(rsCust!rt)
100       txt(9).Enabled = False
101       txt(10) = NulSID(rsCust!rw)
102       txt(10).Enabled = False
          
103       Text(8) = NulSID(rsCust!IdTelp) & NulSID(rsCust!telp)
104       Text(8).Enabled = False
105       Text(9) = NulSID(rsCust!hp1)
106       Text(9).Enabled = False
107       Text(10) = NulSID(rsCust!hp2)
108       Text(10).Enabled = False
109       Text(11) = NulSID(rsCust!noNPWP)
110       Text(11).Enabled = False
          
111       If NulSID(Text(66)) <> "" Then
112           Call text_Validate(66, False)
113       End If
          
      '    Set rsHelp = objO.FetchData("Select Upper(Pekerjaan) as Pekerjaan from [SingleIdOri]..Pekerjaan With (NoLock) WHERE Pekerjaan = '" & nulsid(rsCust!Pekerjaan) & "'")
      '    If Not rsHelp.EOF Then
      '        cboPek.text = UCase(nulsid(rsCust!Pekerjaan))
      '    Else
      '        cboPek.text = "LAINNYA"
      '        Call cboPek_Change
      '
      '        txt(7) = UCase(nulsid(rsCust!Pekerjaan))
      '    End If
      '
      '    If UCase(nulsid(rsCust!Hobi)) <> "" Then
      '        cboHobi.text = UCase(nulsid(rsCust!Hobi))
      '    End If
      '
      '    If UCase(nulsid(rsCust!Agama)) <> "" Then
      '        cboAgama.text = UCase(nulsid(rsCust!Agama))
      '    End If
      '
      '    text(59) = UCase(nulsid(rsCust!Email))
      '    text(49) = Format(nulsid(rsCust!tglLahir), "dd-mm-yyyy")
End Sub

Private Sub cboKel_Change()
        If cboKel.Text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah WITH (NOLOCK) Where UPPER(Kelurahan) = '" & Trim(UCase(cboKel.Text)) & "'")
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                Text(7).Text = rsX("KodePos")
                If txt(8).Enabled Then txt(8).SetFocus
            Else
                txt(8).SetFocus
            End If
        End If
End Sub

Private Sub cboKec_Change()
    If cboKec.Text <> "" Then
        Dim rsX As New ADODB.Recordset, X As Integer
        Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah WITH (NOLOCK) Where UPPER(Kecamatan) = '" & Trim(UCase(cboKec.Text)) & "' Order By Kelurahan")
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
                cboKel.Text = cboKel.Tag
                cboKel.Tag = ""
                If cboKel.Enabled Then cboKel.SetFocus
            End If
        End With
    Else
        cboKel.Tag = "": cboKel.Clear
    End If
End Sub

Private Sub cboKota_Change()
    If cboKota.Text <> "" Then
        Dim rsX As New ADODB.Recordset, X As Integer
        Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah WITH (NOLOCK) Where UPPER(Kota) = '" & Trim(UCase(cboKota.Text)) & "' Order By Kecamatan")
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
                cboKec.Text = cboKec.Tag
                cboKec.Tag = ""
                If cboKec.Enabled Then cboKec.SetFocus
            End If
        End With
    Else
        cboKec.Tag = "": cboKec.Clear
    End If
End Sub

Private Sub PanggilCustomer(Optional str As String)
      Dim nPol As String
      Dim nChassis As String
      Dim n As String
      Dim newFeed As Boolean
      Dim rsCust As ADODB.Recordset

1        On Error GoTo PanggilCustomer_Error
2         Text(65) = ""
3         newFeed = False
4         If str <> "" Then
5             sSQL = "Select CUSTOMER,LEFT(cust_adres,120) as CUST_ADRES,KOTA,KODEPOS,TELPON,NO_POLISI,NO_ENGINE,WMI,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,REMARK,CODE,NPWP,PPN,DISC,KEYNO,[USER],SPK,TGL_REP,JOB,JOBCD,FS_KM_AKHIR,USEJOBDSC,RCDATE,KD_K,CR,CDK,TGL_D,CDD,CRDTCD,FS_TH_MOBIL,USEFS_NM_HOBI,FS_KD_PLG,FS_KD_WIL,DOTAM,USENAME,USEADDR,USECITY,USEPOS,USEPSTN,USEGSM,USECDMA,USET4LAHIR,USETGLLAHIR,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,NOTE,GSM,CDMA,USEAGAMA,USERADD,FLAGKM,HEAVYUSE,PKP,KDPLG_ASR,USEEMAIL,KECAMATAN,USEKECAMATAN from CUSTDAT WITH(NOLOCK) where NO_CHASSIS ='" & str & "'"
6         Else
7             sSQL = "Select CUSTOMER,LEFT(cust_adres,120) as CUST_ADRES,KOTA,KODEPOS,TELPON,NO_POLISI,NO_ENGINE,WMI,NO_CHASSIS,MDL,MODEL,FMDL,COLOR,FS_KD_DEALER,FS_NM_DEALER,DDATE,REMARK,CODE,NPWP,PPN,DISC,KEYNO,[USER],SPK,TGL_REP,JOB,JOBCD,FS_KM_AKHIR,USEJOBDSC,RCDATE,KD_K,CR,CDK,TGL_D,CDD,CRDTCD,FS_TH_MOBIL,USEFS_NM_HOBI,FS_KD_PLG,FS_KD_WIL,DOTAM,USENAME,USEADDR,USECITY,USEPOS,USEPSTN,USEGSM,USECDMA,USET4LAHIR,USETGLLAHIR,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,NOTE,GSM,CDMA,USEAGAMA,USERADD,FLAGKM,HEAVYUSE,PKP,KDPLG_ASR,USEEMAIL,KECAMATAN,USEKECAMATAN from CUSTDAT WITH(NOLOCK) where NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'"
8             str = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
9         End If
10        Set rsCust = objO.FetchData(sSQL)
11        If rsCust.BOF Then
12            If cekUnitCustomer(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
13                If cekSPK(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
14                    If cekHistoryHO(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
15                        If cekSPKHO(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
                             ' If cekBengkelLama(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
                             ' End If
16                        End If
17                    End If
18                End If
19            End If
20            Call text_Validate(2, False)
21        Else
22            pds = cekHistoryPDS(rsCust("NO_CHASSIS"))
23            If cekHistoryOnlyPDS(rsCust("NO_CHASSIS")) Then
24                If Not cekSPK(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) Then
25                    n = cekSPKHO(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")))
26                Else
27                    GoTo LANJUTAN
28                End If
29            Else
30                nChassis = IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
31                If NulSID(rsCust("FS_KD_PLG")) <> Trim(Text(2)) And Trim(Text(2)) <> "" Then
                      
      '                If MsgBox("Data Pelanggan Berbeda. Pakai data Baru ?", vbYesNo, "KODE PELANGGAN") = vbYes Then
      '                    newFeed = True
      '                    Exit Sub
      '                End If
32                End If
                  
LANJUTAN:
33                If Not newFeed Then
                      'text(2) = nulsid(rsCust("FS_KD_PLG"))
34                    Text(65) = NulSID(rsCust("FS_KD_PLG"))
35                    Text(4) = Left(NulSID(rsCust("CUSTOMER")), 60)
36                    Text(5) = Left(NulSID(rsCust("CUST_ADRES")), 60)
37                    If Len(Trim(NulSID(rsCust("CUST_ADRES")))) > 60 Then
38                        Text(56) = Mid(NulSID(rsCust("CUST_ADRES")), 61, Len(NulSID(rsCust("CUST_ADRES"))) - 60)
39                    End If
40                    Text(6) = NulSID(rsCust("KOTA"))
41                    Text(7) = NulSID(rsCust("KODEPOS"))
42                    Text(8) = NulSID(rsCust("TELPON"))
43                    Text(9) = NulSID(rsCust("GSM"))
44                    Text(10) = NulSID(rsCust("CDMA"))
                      
45                    Text(11) = NulSID(rsCust("NPWP"))
46                    Text(51) = NulSID(rsCust("PKP"))
47                    Text(12) = NulSID(rsCust("USEJOBDSC"))
48                    Text(13) = NulSID(Trim(rsCust("USEAGAMA")))
49                    Text(14) = NulSID(rsCust("USEFS_NM_HOBI"))
                      
50                    Text(28) = NulSID(rsCust("PPN"))
51                    Text(29) = NulSID(rsCust("DISC"))
                      
                      '================= DATA MRS =================
52                    Text(30) = NulSID(rsCust("USENAME"))
53                    Text(31) = Left(Trim(NulSID(rsCust("USEADDR"))), 60)
54                    If Len(Trim(NulSID(rsCust("USEADDR")))) > 60 Then
55                        Text(57) = Mid(Trim(NulSID(rsCust("USEADDR"))), 61, 60)
56                        If Len(Trim(NulSID(rsCust("USEADDR")))) > 120 Then
57                            Text(58) = Mid(Trim(NulSID(rsCust("USEADDR"))), 121, Len(Trim(NulSID(rsCust("USEADDR")))) - 120)
58                        End If
59                    End If
60                    Text(32) = NulSID(rsCust("USECITY"))
61                    Text(33) = NulSID(rsCust("USEPOS"))
62                    Text(34) = NulSID(rsCust("USEPSTN"))
63                    Text(35) = NulSID(rsCust("USEGSM"))
64                    Text(36) = NulSID(rsCust("USECDMA"))
65                    Text(48) = NulSID(rsCust("USET4LAHIR"))
66                    n = Text(49).Mask
67                    Text(49).Mask = ""
68                    Text(49) = NulSID(IIf(Replace(rsCust("USETGLLAHIR"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("USETGLLAHIR"), "/", "-") = "01-01-1900", "", Replace(rsCust("USETGLLAHIR"), "/", "-"))))
69                    Text(49).Mask = n
70                    Text(59) = NulSID(Trim(rsCust("USEEMAIL")))
                      
                      '================= DATA ASURANSI =================
71                    If Trim(rsCust("ASURANSI")) = "" Then
72                        chAsuransi.value = Unchecked
73                    Else
74                        chAsuransi.value = Checked
75                    End If
76                    Text(52) = NulSID(rsCust("KDPLG_ASR"))
77                    Text(37) = NulSID(rsCust("ASURANSI"))
78                    Text(38) = NulSID(rsCust("ASR_ADDR"))
79                    Text(39) = NulSID(rsCust("ASR_CITY"))
80                    n = Text(40).Mask
81                    Text(40).Mask = ""
82                    Text(40) = NulSID(IIf(Replace(rsCust("ASR_TGLM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("ASR_TGLM"), "/", "-") = "01-01-1900", "", Replace(rsCust("ASR_TGLM"), "/", "-"))))
83                    Text(40).Mask = n
84                    n = Text(41).Mask
85                    Text(41).Mask = ""
86                    Text(41) = NulSID(IIf(Replace(rsCust("ASR_TGLS"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("ASR_TGLS"), "/", "-") = "01-01-1900", "", Replace(rsCust("ASR_TGLS"), "/", "-"))))
87                    Text(41).Mask = n
                      
                      Dim rsAsr As ADODB.Recordset
88                    Set rsAsr = objO.FetchData("SELECT * FROM T_ASURANSI WITH(NOLOCK) WHERE ASURANSI = '" & Trim(Text(37)) & "'")
89                    If Not rsAsr.EOF Then
90                        Text(53) = NulSID(rsAsr("PSTN"))
91                        Text(54) = NulSID(rsAsr("GSM"))
92                        Text(55) = NulSID(rsAsr("CDMA"))
93                    End If
94                    Text(42) = NulSID(rsCust("ASR_LIMIT"))
95                    Text(43) = NulSID(rsCust("ASR_NOTE"))
96                    Text(44) = NulSID(rsCust("NOTE"))
97                End If
                  '================= DATA KENDARAAN =================
98                Text(45) = NulSID(rsCust("FS_KM_AKHIR"))
                  ''--------------------------------------
                  ''2016-02-01 | ambil data dari salesunit
                  ''--------------------------------------
99                 Text(15) = NulSID(rsCust("FS_KD_DEALER"))
100                Text(16) = NulSID(rsCust("FS_NM_DEALER"))
                  ''---------------------------------------
                  
      '            Text(15) = ""
      '            Text(16) = ""
                  '' 2016-04-11 || jika fs_kd_dealer kosong maka dicari dari salesunit
101               If NulSID(rsCust!FS_KD_DEALER) = "" Then
                     ''2016-07-29 "" SUMMARY MENYIMPAN SEMUA DATA PENJUALAN SEKALIPUN ada RETUR SEHINGGA YANG DIAKUI PENJUALAN TERAKHIR
102                  If Log.RPLHistori = 0 Then
103                       Set rsTemp = objO.FetchData("SELECT FS_KD_DEALER,FS_NM_DEALER FROM SALESUNIT WHERE FS_KD_DEALER <> '' AND NORANGKA = '" & Trim(str) & "' ORDER BY TGLENTRY DESC ")
104                  Else
105                       Set rsTemp = objO.FetchData("select ISNULL(b.fs_kd_dealer,'') as FS_KD_DEALER,ISNULL(b.fs_nm_dealer,'') as FS_NM_DEALER from [" & dbReplikasiHO & "]..Summary a with (nolock) left join [" & dbReplikasiHO & "]..t_dealer b with (nolock) on a.KodeDealer = b.kodecabang where isnull(b.FS_KD_DEALER,'') <> '' AND a.NoRangka = '" & Trim(str) & "'")
106                  End If
107                  If Not rsTemp.EOF Then
108                      Text(15) = NulSID(rsTemp!FS_KD_DEALER)
109                      Text(16) = NulSID(rsTemp("FS_NM_DEALER"))
110                  End If
111               End If
                              
112               Text(17) = NulSID(rsCust("CODE"))
113               Text(18).Mask = ""
114               Text(18) = IIf(TglDDMMYYYY(NulTgl(rsCust("DDATE"))) = "01-01-1900" Or TglDDMMYYYY(NulTgl(rsCust("DDATE"))) = "01-01-3000", "", TglDDMMYYYY(NulTgl(rsCust("DDATE")))) 'nulsid(IIf(Replace(rsCust("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsCust("DDATE"), "/", "-"))))
115               Text(18).Mask = "99-99-9999"
116               Text(19).Mask = ""
117               Text(19) = IIf(TglDDMMYYYY(NulTgl(rsCust("DOTAM"))) = "01-01-1900" Or TglDDMMYYYY(NulTgl(rsCust("DOTAM"))) = "01-01-3000", "", TglDDMMYYYY(NulTgl(rsCust("DOTAM")))) 'nulsid(IIf(Replace(rsCust("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsCust("DOTAM"), "/", "-"))))
118               Text(19).Mask = "99-99-9999"
119               Text(20) = NulSID(rsCust("FS_TH_MOBIL"))
120               Text(21) = NulSID(rsCust("MDL"))
121               Text(22) = NulSID(rsCust("MODEL"))
122               Text(23) = NulSID(rsCust("FMDL"))
123               Text(24) = NulSID(rsCust("COLOR"))
124               If Len(NulSID(rsCust("NO_POLISI"))) > 0 Then
125                   nPol = NulSID(rsCust("NO_POLISI"))
                      Dim nDpn1 As String
                      Dim nTgh1 As String
                      Dim nBlk1 As String
126                   nDpn1 = ""
127                   nTgh1 = ""
128                   nBlk1 = ""
129                   Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
130                   Text(25) = Trim(nDpn1)
131                   Text(46) = Trim(nTgh1)
132                   Text(47) = Trim(nBlk1)
133               Else
134                   Text(25) = ""
135                   Text(46) = ""
136                   Text(47) = ""
137               End If
138               Text(26) = NulSID(rsCust("NO_ENGINE"))
139               Text(27) = NulSID(rsCust("KEYNO"))
                  
140               If Log.loyalityCustomer = 1 Then
141                   Set rsTemp = objO.FetchData("SELECT * FROM [RplNasApp]..masteruser WHERE verifikasi = 1 and alamatemail = '" & Trim(Text(59)) & "' ")
142                   If Not rsTemp.EOF Then
143                       Text(59).Enabled = False
144                   End If
                      
145                   Set rsTemp = objO.FetchData("SELECT b.* FROM RPLNASAPP..MASTERUSER a INNER JOIN RPLNASAPP..NewCustomer b ON a.alamatemail = b.email WHERE a.VERIFIKASI = 1 and isnull(b.isUpdate,0) = 1 AND a.ALAMATEMAIL = '" & Trim(Text(59)) & "' ")
146                   If Not rsTemp.EOF Then
147                       Text(2).Enabled = False
148                       Text(4) = UCase(Trim(NulSID(rsTemp!namacust1) & " " & NulSID(rsTemp!namacust2) & " " & NulSID(rsTemp!namacust3)))
149                       If Len(Trim(NulSID(rsTemp("alamat")) & " RT." & NulSID(rsTemp("rt")) & " RW." & NulSID(rsTemp("rw")) & " RT." & NulSID(rsTemp("rt")) & " Kelurahan." & NulSID(rsTemp("kelurahan")))) > 60 Then
150                           Text(5) = UCase(Left(Trim(NulSID(rsTemp("alamat")) & " RT." & NulSID(rsTemp("rt")) & " RW." & NulSID(rsTemp("rw")) & " RT." & NulSID(rsTemp("rt")) & " Kelurahan." & NulSID(rsTemp("kelurahan"))), 60))
151                           Text(56) = UCase(Mid(Trim(NulSID(rsTemp("alamat")) & " RT." & NulSID(rsTemp("rt")) & " RW." & NulSID(rsTemp("rw")) & " RT." & NulSID(rsTemp("rt")) & " Kelurahan." & NulSID(rsTemp("kelurahan"))), 61, 60))
152                       Else
153                           Text(5) = UCase(Left(Trim(NulSID(rsTemp("alamat")) & " RT." & NulSID(rsTemp("rt")) & " RW." & NulSID(rsTemp("rw")) & " RT." & NulSID(rsTemp("rt")) & " Kelurahan." & NulSID(rsTemp("kelurahan"))), 60))
154                       End If
155                       Text(60) = UCase(NulSID(rsTemp!kecamatan))
156                       Text(6) = UCase(NulSID(rsTemp!Kota))
157                       Text(7) = UCase(NulSID(rsTemp!KodePos))
158                       Text(14) = UCase(NulSID(rsTemp!hobi))
159                       Text(13) = UCase(NulSID(rsTemp!Agama))
160                       Text(12) = UCase(NulSID(rsTemp!pekerjaan))
161                       Text(8) = UCase(NulSID(rsTemp!telp))
162                       Text(9) = UCase(NulSID(rsTemp!hp1))
163                       Text(10) = UCase(NulSID(rsTemp!hp2))
164                   End If
165               End If
166           End If
167       End If

168      On Error GoTo 0
169      Exit Sub

PanggilCustomer_Error:

170       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilCustomer of Form frmSIDCustomer")
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If gMainKeyCari <> "" Then
        txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        Sendkeys ("{TAB}")
    End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 63, 64, 65, 66, 11
                If gMainKeyCari = "" Then gMainKeyCari = txt(Index)
                gSearchOption = "TOPOLOGY"
                If frmCari.Visible Then
                    frmCari.Hide
                End If
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
                frmCari.show (vbModal)
            Case 0, 1, 2, 3, 33
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
                If gMainKeyCari = "" Then gMainKeyCari = txt(Index)
                gSearchOption = "TOPOLOGY"
                If frmCari.Visible Then
                    frmCari.Hide
                End If
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
                frmCari.show (vbModal)
            Case 7
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDPEKERJAAN"
                frmCari.show (vbModal)
            Case 13
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDHOBI"
                frmCari.show (vbModal)
            Case 12
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDAGAMA"
                frmCari.show (vbModal)
        End Select
    ElseIf KeyCode = 13 Then
        Call txt_Validate(Index, False)
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        'SendKeys "{TAB}"
          
    Else
        If KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            Select Case Index
                Case 9, 10, 12, 13, 14, 15, 40, 41, 43, 44, 45, 46, 55, 56, 58, 59, 60, 61, 4, 5
                    If KeyAscii <= 57 And KeyAscii >= 48 Then
                        KeyAscii = KeyAscii
                    Else
                        KeyAscii = 0
                    End If
'                Case 7, 12, 13, 0, 1, 2, 3
'                    KeyAscii = 0
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        End If
    End If
End Sub


Private Sub cboKotaS_Change()
    If cboKotaS.Text <> "" Then
        Dim rsX As New ADODB.Recordset, X As Integer
        Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah WITH (NOLOCK) Where UPPER(Kota) = '" & Trim(UCase(cboKotaS.Text)) & "' Order By Kecamatan")
        With cboKecS
            .Clear: X = 0: cboKotaS.Tag = ""
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                Do Until rsX.EOF
                    If X = 0 Then cboKecS.Tag = rsX("Kecamatan")
                    .AddItem NulSID(rsX("Kecamatan"))
                    rsX.MoveNext
                    X = X + 1
                Loop
                cboKecS.Text = cboKecS.Tag
                cboKecS.Tag = ""
                If cboKecS.Enabled Then
                cboKecS.SetFocus
                End If
            End If
        End With
    Else
        cboKecS.Tag = "": cboKecS.Clear
    End If
End Sub

Private Sub cboKecS_Change()
    If cboKecS.Text <> "" Then
        Dim rsX As New ADODB.Recordset, X As Integer
        Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah WITH (NOLOCK) Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecS.Text)) & "' Order By Kelurahan")
        With cboKelS
            .Clear: X = 0: cboKelS.Tag = ""
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                Do Until rsX.EOF
                    If X = 0 Then cboKelS.Tag = rsX("Kelurahan")
                    .AddItem NulSID(rsX("Kelurahan"))
                    rsX.MoveNext
                    X = X + 1
                Loop
                'cboKelS.text = cboKelS.Tag
                cboKelS.Tag = ""
                If cboKelS.Enabled Then
                cboKelS.SetFocus
                End If
            End If
        End With
    Else
        cboKelS.Tag = "": cboKelS.Clear
    End If
End Sub

Private Sub cboKelS_Change()
    If cboKelS.Text <> "" Then
        Dim rsX As New ADODB.Recordset, X As Integer
        Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah WITH (NOLOCK) Where UPPER(Kelurahan) = '" & Trim(UCase(cboKelS.Text)) & "'")
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Text(33).Text = rsX("KodePos")
             
        End If
    Else
        Text(33).Text = ""
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    If Index = 4 Or Index = 5 Then
        txt(Index) = Format(NulSID(txt(Index)), "0##")
    End If
    If gMainKeyCari <> "" Then
        txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        'SendKeys ("{TAB}")
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
        If Index = 12 Then
        If NulSID(txt(Index)) <> "" Then
            Dim rsAgama As ADODB.Recordset
            Set rsAgama = objO.FetchData("SELECT * FROM [SINGLEIDORI]..agama WITH (NOLOCK) WHERE agama = '" & NulSID(txt(Index)) & "'")
            If rsAgama.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDAGAMA"
                frmCari.show (vbModal)
            End If
        End If
    ElseIf Index = 7 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..PEKERJAAN WITH (NOLOCK) WHERE pekerjaan = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDPEKERJAAN"
                frmCari.show (vbModal)
            End If
        End If
    ElseIf Index = 13 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..HOBI WITH (NOLOCK) WHERE HOBI = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDHOBI"
                frmCari.show (vbModal)
            End If
        End If
     
    ElseIf Index = 0 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
'                txt(Index) = ""
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
            End If
        End If
    ElseIf Index = 1 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(0)) & "' and KOTA = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
'                txt(Index) = ""
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
            End If
        End If
    ElseIf Index = 2 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(0)) & "' and KOTA = '" & NulSID(txt(1)) & "' and KECAMATAN = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
'                txt(Index) = ""
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
            End If
        End If
    ElseIf Index = 3 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(0)) & "' and KOTA = '" & NulSID(txt(1)) & "' and KECAMATAN = '" & NulSID(txt(2)) & "' and KELURAHAN = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
'                txt(Index) = ""
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
            End If
        End If
    End If
End Sub
