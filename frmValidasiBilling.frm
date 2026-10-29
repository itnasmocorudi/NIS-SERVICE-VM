VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmValidasiBilling 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Validasi Billing"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   12975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmValidasiBilling.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   55
      Top             =   -360
      Width           =   495
   End
   Begin VB.Frame frmeWO 
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
      Height          =   7890
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   10605
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   930
         TabIndex        =   44
         Top             =   2100
         Width           =   3855
         Begin VB.OptionButton opTpSvc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Warranty"
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
            Left            =   1920
            TabIndex        =   47
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opTpSvc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Program"
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
            Left            =   960
            TabIndex        =   46
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton opTpSvc 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Regular"
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
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00DCFCFB&
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6480
         Width           =   10575
      End
      Begin VB.Frame frmeValid 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   0
         TabIndex        =   37
         Top             =   6600
         Width           =   10575
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1080
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   9375
         End
         Begin ucXPButton.XPButton cmdSimpan 
            Height          =   495
            Left            =   4680
            TabIndex        =   57
            Top             =   660
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
            MICON           =   "frmValidasiBilling.frx":26F5
            PICN            =   "frmValidasiBilling.frx":2711
            PICH            =   "frmValidasiBilling.frx":2C24
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
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan :"
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
            TabIndex        =   38
            Top             =   240
            Width           =   945
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   930
         TabIndex        =   5
         Top             =   510
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   930
         TabIndex        =   6
         Top             =   795
         Width           =   4815
         _ExtentX        =   8493
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
         TabIndex        =   7
         Top             =   1110
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   6
         Left            =   930
         TabIndex        =   8
         Top             =   1410
         Width           =   1365
         _ExtentX        =   2408
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
         Left            =   6750
         TabIndex        =   9
         Top             =   510
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
         Index           =   10
         Left            =   6750
         TabIndex        =   10
         Top             =   795
         Width           =   2895
         _ExtentX        =   5106
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
         Index           =   13
         Left            =   6750
         TabIndex        =   11
         Top             =   1410
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   6750
         TabIndex        =   12
         Top             =   1110
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
         Index           =   11
         Left            =   9690
         TabIndex        =   13
         Top             =   795
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
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   210
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   930
         TabIndex        =   15
         Top             =   210
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   6750
         TabIndex        =   16
         Top             =   210
         Width           =   1410
         _ExtentX        =   2487
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
         Index           =   15
         Left            =   6750
         TabIndex        =   17
         Top             =   1725
         Width           =   2655
         _ExtentX        =   4683
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
         Index           =   17
         Left            =   9270
         TabIndex        =   18
         Top             =   510
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
         Index           =   7
         Left            =   2400
         TabIndex        =   19
         Top             =   1410
         Width           =   1605
         _ExtentX        =   2831
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
         Index           =   8
         Left            =   4140
         TabIndex        =   20
         Top             =   1410
         Width           =   1605
         _ExtentX        =   2831
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   2800
         TabIndex        =   21
         Top             =   217
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   6750
         TabIndex        =   22
         Top             =   2010
         Width           =   2655
         _ExtentX        =   4683
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
         Index           =   18
         Left            =   930
         TabIndex        =   41
         Top             =   1730
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   21
         Left            =   9270
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   "999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   19
         Left            =   930
         TabIndex        =   51
         Top             =   2400
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   20
         Left            =   6750
         TabIndex        =   52
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin TrueOleDBGrid70.TDBGrid JobGrid 
         Height          =   2295
         Left            =   0
         TabIndex        =   53
         Top             =   2760
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   4048
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
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
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).DataField=   ""
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(35)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(39)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(43)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(44)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(46)=   "Column(9).Order=10"
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
         Caption         =   "Job Order"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Named:id=33:Normal"
         _StyleDefs(77)  =   ":id=33,.parent=0"
         _StyleDefs(78)  =   "Named:id=34:Heading"
         _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(80)  =   ":id=34,.wraptext=-1"
         _StyleDefs(81)  =   "Named:id=35:Footing"
         _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=36:Selected"
         _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=37:Caption"
         _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(87)  =   "Named:id=38:HighlightRow"
         _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=39:EvenRow"
         _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(91)  =   "Named:id=40:OddRow"
         _StyleDefs(92)  =   ":id=40,.parent=33"
         _StyleDefs(93)  =   "Named:id=41:RecordSelector"
         _StyleDefs(94)  =   ":id=41,.parent=34"
         _StyleDefs(95)  =   "Named:id=42:FilterBar"
         _StyleDefs(96)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid JobGridS 
         Height          =   1455
         Left            =   0
         TabIndex        =   61
         Top             =   5040
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2566
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
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
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
         Caption         =   "Sublet"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
      Begin VB.Image btnBatal 
         Height          =   255
         Left            =   10320
         Picture         =   "frmValidasiBilling.frx":31C6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblPDS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9840
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label LblReturn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9840
         TabIndex        =   48
         Top             =   2040
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Wo"
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
         Left            =   195
         TabIndex        =   43
         Top             =   2055
         Width           =   585
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
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
         TabIndex        =   42
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label Label6 
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
         TabIndex        =   34
         Top             =   555
         Width           =   405
      End
      Begin VB.Label Label3 
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
         TabIndex        =   33
         Top             =   855
         Width           =   495
      End
      Begin VB.Label Label4 
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
         TabIndex        =   32
         Top             =   1155
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telpon"
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
         Left            =   300
         TabIndex        =   31
         Top             =   1455
         Width           =   480
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
         Left            =   5955
         TabIndex        =   30
         Top             =   555
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
         Left            =   6180
         TabIndex        =   29
         Top             =   840
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
         Left            =   6120
         TabIndex        =   28
         Top             =   1455
         Width           =   480
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
         Left            =   135
         TabIndex        =   27
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label2 
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
         Left            =   5955
         TabIndex        =   26
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
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
         TabIndex        =   25
         Top             =   1775
         Width           =   780
      End
      Begin VB.Label Label55 
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
         Left            =   8145
         TabIndex        =   24
         Top             =   555
         Width           =   975
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
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
         Left            =   5955
         TabIndex        =   23
         Top             =   2055
         Width           =   645
      End
   End
   Begin TrueOleDBGrid70.TDBGrid listValidasi 
      Height          =   7455
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   13150
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
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
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST VALIDASI"
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   12577002
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
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
      _StyleDefs(77)  =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(78)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(79)  =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(80)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2760
      TabIndex        =   35
      Top             =   195
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   195
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
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
      Mask            =   ">CCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSComCtl2.DTPicker tgl1 
      Height          =   375
      Left            =   3600
      TabIndex        =   39
      Top             =   150
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   316473345
      CurrentDate     =   39859
   End
   Begin MSComCtl2.DTPicker tgl2 
      Height          =   375
      Left            =   5400
      TabIndex        =   40
      Top             =   150
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   316473345
      CurrentDate     =   39859
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   11280
      TabIndex        =   56
      Top             =   8280
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
      MICON           =   "frmValidasiBilling.frx":6AE1
      PICN            =   "frmValidasiBilling.frx":6AFD
      PICH            =   "frmValidasiBilling.frx":6FD1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdProses 
      Height          =   495
      Left            =   7080
      TabIndex        =   58
      Top             =   80
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&v"
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
      MICON           =   "frmValidasiBilling.frx":751C
      PICN            =   "frmValidasiBilling.frx":7538
      PICH            =   "frmValidasiBilling.frx":79FA
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
      Left            =   9600
      TabIndex        =   59
      Top             =   8280
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
      MICON           =   "frmValidasiBilling.frx":7E25
      PICN            =   "frmValidasiBilling.frx":7E41
      PICH            =   "frmValidasiBilling.frx":83C9
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
      Left            =   8400
      TabIndex        =   60
      Top             =   8280
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
      MICON           =   "frmValidasiBilling.frx":8A05
      PICN            =   "frmValidasiBilling.frx":8A21
      PICH            =   "frmValidasiBilling.frx":8ECC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   54
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. WO"
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
      Left            =   600
      TabIndex        =   36
      Top             =   210
      Width           =   690
   End
End
Attribute VB_Name = "frmValidasiBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsJob As New ADODB.Recordset
Dim rsJobS As New ADODB.Recordset
Dim rsGrid As ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Dim objBilling As New BClsBilling

Private Sub hapusForm()
    
    For i = 0 To 21
        n = text(i).Mask
        text(i).Mask = ""
        text(i) = ""
        text(i).Mask = n
    Next
    
    frmeWO.Visible = False
    If rsJob.State <> adStateClosed Then rsJob.Close
    
    CmdOpen.Caption = "v"
    frmeWO.Height = CmdOpen.Top + CmdOpen.Height + 50
End Sub

Private Sub createTblJob(stat As Boolean)
    
    If stat Then
        If rsJob.State Then rsJob.Close
        Set rsJob = Nothing
        With rsJob
            .Fields.Append "GROUP", adChar, 5
            .Fields.Append "JOB", adChar, 50
            .Fields.Append "KET", adChar, 50
            .Fields.Append "RATE", adDouble
            .Fields.Append "HARGA", adDouble
            .Fields.Append "HARGABELI", adDouble
            .Fields.Append "DISC", adDouble
            .Fields.Append "RP", adBoolean
            .Fields.Append "DISCRP", adDouble
            .Fields.Append "STLDISC", adDouble
            .Open
        End With
        
    End If
End Sub

Private Sub createTblJobS(stat As Boolean)
    
    If stat Then
        If rsJobS.State Then rsJobS.Close
        Set rsJobS = Nothing
        With rsJobS
            .Fields.Append "GROUP", adChar, 5
            .Fields.Append "JOB", adChar, 50
            .Fields.Append "KET", adChar, 50
            .Fields.Append "RATE", adDouble
            .Fields.Append "HARGA", adDouble
            .Fields.Append "HARGABELI", adDouble
            .Open
        End With
        
    End If
End Sub

Private Sub headColumns()
    
    With JobGrid
        .Columns(0).Caption = "GROUP"
        .Columns(0).Width = 800
        .Columns(0).DataField = "GROUP"
        .Splits(0).Columns(0).Locked = True
        .Columns(1).Caption = "JOB ORDER"
        .Columns(1).Width = 2500
        .Columns(1).DataField = "JOB"
        .Splits(0).Columns(1).Locked = True
        .Columns(2).Caption = "KETERANGAN JOB"
        .Columns(2).Width = 0
        .Columns(2).DataField = "KET"
        .Splits(0).Columns(2).Visible = False
        .Splits(0).Columns(2).Locked = True
        .Columns(3).Caption = "RATE"
        .Columns(3).Width = 800
        .Columns(3).DataField = "RATE"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Caption = "Harga Beli"
        .Columns(4).Width = 1200
        .Columns(4).DataField = "HARGABELI"
        .Splits(0).Columns(4).Locked = True
        .Columns(5).Caption = "Harga Jual"
        .Columns(5).Width = 1200
        .Columns(5).DataField = "HARGA"
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Caption = "Disc"
        .Columns(6).Width = 800
        .Columns(6).DataField = "DISC"
        .Columns(7).Caption = "Rp"
        .Columns(7).Width = 800
        .Columns(7).DataField = "RP"
        .Columns(7).Alignment = dbgCenter
        .Columns(8).Caption = "Rupiah"
        .Columns(8).Width = 1200
        .Columns(8).DataField = "DISCRP"
        .Columns(9).Caption = "Sub Total"
        .Columns(9).Width = 1500
        .Columns(9).DataField = "STLDISC"
        .Splits(0).Columns(9).Locked = True
        
    End With

    With JobGridS
        .Columns(0).Caption = "GROUP"
        .Columns(0).Width = 800
        .Columns(0).DataField = "GROUP"
        .Splits(0).Columns(0).Locked = True
        .Columns(1).Caption = "JOB ORDER"
        .Columns(1).Width = 2500
        .Columns(1).DataField = "JOB"
        .Splits(0).Columns(1).Locked = True
        .Columns(2).Caption = "KETERANGAN JOB"
        .Columns(2).Width = 0
        .Columns(2).DataField = "KET"
        .Splits(0).Columns(2).Visible = False
        .Splits(0).Columns(2).Locked = True
        .Columns(3).Caption = "RATE"
        .Columns(3).Width = 800
        .Columns(3).DataField = "RATE"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Caption = "Harga Beli"
        .Columns(4).Width = 1200
        .Columns(4).DataField = "HARGABELI"
        .Splits(0).Columns(4).Locked = True
        .Columns(5).Caption = "Harga Jual"
        .Columns(5).Width = 1200
        .Columns(5).DataField = "HARGA"
        .Splits(0).Columns(5).Locked = True
        
        
    End With

End Sub

Private Sub btnBatal_Click()
    frmeWO.Visible = False
End Sub

Private Sub CmdDelete_Click()
    If listValidasi.Columns(1).value <> "" And listValidasi.Columns(4).value = "" Then
        If bolehDiskon(Trim(listValidasi.Columns(1).value)) Then
            If MsgBox("Akan menghapus Pemberian discount untuk No. WO " & Trim(listValidasi.Columns(1).value) & " ?", vbQuestion + vbYesNo, "Hapus Diskon") = vbYes Then
                Call objO.UpdateData("UPDATE VALIDBILLING SET DELLOG = '" & Trim(gSA) & "', TGLDEL = GETDATE() WHERE NO_WO = '" & Trim(listValidasi.Columns(1).value) & "' and dellog = ''")
                Call refreshGrid
            End If
        Else
            MsgBox "Wo sudah ter-invoice atau belum di validasi. Cek status WO !! ", vbInformation + vbOKOnly, "Validasi"
        End If
    Else
        MsgBox "Sudah terhapus tidak bisa di-Ubah. Pilih ulang WO yang akan diproses.", vbInformation + vbOKOnly, "Validasi"
    End If
End Sub

Private Sub cmdEdit_Click()
    If listValidasi.Columns(1).value <> "" And listValidasi.Columns(4).value = "" Then
        If bolehDiskon(Trim(listValidasi.Columns(1).value)) Then
            Dim n As String
            n = Trim(listValidasi.Columns(1).value)
            Call hapusForm
            text(0) = n
            Call panggilWO(Trim(text(0)), 1)
            Call refreshGrid
        Else
            MsgBox "Wo sudah ter-invoice atau belum di validasi. Cek status WO !! ", vbInformation + vbOKOnly, "Validasi"
        End If
    Else
        MsgBox "Sudah terhapus tidak bisa di-Ubah. Pilih ulang WO yang akan diproses.", vbInformation + vbOKOnly, "Validasi"
    End If

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If CmdOpen.Caption = "v" Then
        frmeWO.Height = frmeWO.Height + frmeValid.Height + 50
        CmdOpen.Caption = "^"
    Else
        frmeWO.Height = frmeWO.Height - frmeValid.Height - 50
        CmdOpen.Caption = "v"
    End If
End Sub

Private Sub cmdProses_Click()
    Call refreshGrid
End Sub

Private Sub cmdSimpan_Click()
    If Trim(Text1) <> "" Then
            rsJob.MoveFirst
            Do While Not rsJob.EOF
                If CDbl(rsJob("disc")) > OnDisKBG Then
                    MsgBox "Diskon melebihi batas maksimal (" & OnDisKBG & "%)", vbInformation + vbOKOnly, "Validasi"
                    Exit Sub
                End If
                rsJob.MoveNext
            Loop
            rsJob.MoveFirst
            Do While Not rsJob.EOF
                Dim nilDsc As Double
                If rsJob.Fields(7).value = True Then
                    If Round((rsJob.Fields(8).value / rsJob.Fields(4).value) * 100, 0) > OnDisKBG Then
                        MsgBox "Diskon melebihi batas maksimal (" & OnDisKBG & "%)", vbInformation + vbOKOnly, "Validasi"
                        Exit Sub
                    End If
                End If
                rsJob.MoveNext
            Loop
            
        If rsJob.State <> adStateClosed Then
            Call objO.UpdateData("UPDATE VALIDBILLING SET DELLOG = '" & Trim(gSA) & "', TGLDEL = GETDATE() WHERE NO_WO = '" & Trim(text(0)) & "' and dellog = ''")
            rsJob.MoveFirst
            Do While Not rsJob.EOF
                Call objO.InsertData("INSERT INTO VALIDBILLING VALUES ('" & Trim(text(0)) & "','" & Trim(text(15)) & "','" & Trim(text(14)) & "','" & Trim(rsJob("job")) & "','" & setKoma(CDbl(rsJob("disc"))) & "'," & setKoma(CDbl(rsJob("Harga"))) & "," & setKoma(CDbl(rsJob("StlDisc"))) & ", '" & Trim(Text1) & "', '" & Trim(gSA) & "',getdate(),'','')")
                rsJob.MoveNext
            Loop
            MsgBox "Data Validasi telah disimpan"
        End If
        
        Call hapusForm
        Call refreshGrid
    Else
        MsgBox "Isi alasan perubahan harga Invoice !!!"
    End If
End Sub

Private Sub Form_Load()
    Tgl1 = "01/" & Month(isnow) & "/" & year(isnow)
    Tgl2 = Format(GetTglAkhir(Month(isnow), year(isnow), 0), "DD/MM/YYYY")
    Call refreshGrid
End Sub

Private Sub refreshGrid()
    Set rsGrid = objO.FetchData("SELECT distinct convert(varchar(10),TGLENTRY,101) as TGLENTRY, No_Wo, No_Polisi, Keterangan, case convert(varchar(10),TGLDel,101) When '01/01/1900' then '' else convert(varchar(10),TGLDel,101) End as TGLdel FROM VALIDBILLING WITH(NOLOCK) WHERE TGLENTRY BETWEEN '" & Format(Tgl1, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(Tgl2, "MM/DD/YYYY 23:59:00") & "' AND ENTRYLOG = '" & Trim(gSA) & "' ORDER BY convert(varchar(10),TGLENTRY,101), NO_WO")
    If Not rsGrid.EOF Then
        Set listValidasi.DataSource = rsGrid
        Call headColumnsList
    End If
End Sub

Private Sub headColumnsList()
    
    With listValidasi
        .Columns(0).Caption = "Tgl Setuju"
        .Columns(0).DataField = "TGLENTRY"
        .Columns(0).Width = 1200
        .Columns(0).Locked = True
        .Columns(1).Caption = "No Order"
        .Columns(1).DataField = "NO_WO"
        .Columns(1).Width = 1200
        .Columns(1).Locked = True
        .Columns(2).Caption = "No Polisi"
        .Columns(2).DataField = "NO_polisi"
        .Columns(2).Width = 1200
        .Columns(2).Locked = True
        .Columns(3).Caption = "Keterangan"
        .Columns(3).DataField = "KETERANGAN"
        .Columns(3).Width = 5000
        .Columns(3).Locked = True
        .Columns(4).Caption = "Tgl Hapus"
        .Columns(4).DataField = "TGLdel"
        .Columns(4).Width = 1200
        .Columns(4).Locked = True
    End With
End Sub


Private Sub JobGrid_AfterColUpdate(ByVal ColIndex As Integer)
    
    Select Case ColIndex
        Case 6
            If Val(JobGrid.Columns(6)) > OnDisKBG Then
                MsgBox "Diskon melebihi batas maksimal (" & OnDisKBG & "%)", vbInformation + vbOKOnly, "Validasi"
                Exit Sub
            End If
            JobGrid.Columns(8) = Round((JobGrid.Columns(6) / 100) * JobGrid.Columns(5))
            With JobGrid
                If .Columns(ColIndex) >= 0 And .Columns(ColIndex) <= 100 Then
                    .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(5)))
                Else
                    MsgBox "Discount tidak boleh lebih kecil dari nol atau lebih besar dari 100 !!"
                    rsJob.Find "JOB = '" & Trim(.Columns(1)) & "'", , adSearchForward, 1
                    If Not rsJob.EOF Then
                        rsJob("DISC") = 0
                        rsJob("DISCRP") = 0
                        .Columns(ColIndex) = 0
                        .Columns(ColIndex + 1) = 0
                        rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(5)))
                        .Columns(ColIndex + 3) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(5)))
                    End If
                End If
            End With
        Case 7
            With JobGrid
                If .Columns(ColIndex) Then
                    .Splits(0).Columns(6).Locked = True
                    .Splits(0).Columns(8).Locked = False
                Else
                    .Splits(0).Columns(6).Locked = False
                    .Splits(0).Columns(8).Locked = True
                End If
            End With
        Case 8
            JobGrid.Columns(6) = Round((JobGrid.Columns(ColIndex) / CDbl(JobGrid.Columns(5))) * 100, 4)
            With JobGrid
                If .Columns(6) >= 0 And .Columns(6) <= 100 Then
                    .Columns(9) = Round(((100 - CDbl(.Columns(6))) / 100) * CDbl(.Columns(5)))
                Else
                    MsgBox "Discount tidak boleh lebih kecil dari nol !!"
                    rsJob.Find "JOB = '" & Trim(.Columns(1)) & "'", , adSearchForward, 1
                    If Not rsJob.EOF Then
                        rsJob("DISC") = 0
                        rsJob("DISCRP") = 0
                        .Columns(ColIndex) = 0
                        .Columns(ColIndex + 1) = 0
                        rsJob("STLDISC") = Round(((100 - CDbl(.Columns(6))) / 100) * CDbl(.Columns(5)))
                        .Columns(ColIndex + 1) = Round(((100 - CDbl(.Columns(6))) / 100) * CDbl(.Columns(5)))
                    End If
                End If
            End With
    End Select
    Set JobGrid.DataSource = rsJob
    Call headColumns
End Sub

'Private Sub listValidasi_DblClick()
'Dim jawab As String
'
'ulangi:
'    If listValidasi.Columns(1).value <> "" And listValidasi.Columns(4).value = "01/01/1900" Then
'        If bolehDiskon(Trim(listValidasi.Columns(1).value)) Then
'            jawab = UCase(InputBox("Masukkan Pilihan Proses. Ketik Huruf didalam Kurung. " & vbCr & vbCr & "[U]bah " & vbCr & "[H]apus ", "Pilihan Proses"))
'            Select Case jawab
'                Case "U"
'                    If MsgBox("Akan mengubah Pemberian discount untuk No. WO " & Trim(listValidasi.Columns(1).value) & " ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Ubah Diskon") = vbYes Then
'                        n = Trim(listValidasi.Columns(1).value)
'                        Call hapusForm
'                        text(0) = n
'                        Call panggilWO(Trim(text(0)), 1)
'                        'Call refreshGrid
'                    End If
'                Case "H"
'                    If MsgBox("Akan menghapus Pemberian discount untuk No. WO " & Trim(listValidasi.Columns(1).value) & " ?", vbYesNo + vbDefaultButton1 + vbCritical, "Hapus Diskon") = vbYes Then
'                        Call objO.UpdateData("UPDATE VALIDBILLING SET DELLOG = '" & Trim(gSA) & "', TGLDEL = GETDATE() WHERE NO_WO = '" & Trim(listValidasi.Columns(1).value) & "' and dellog = ''")
'                        'Call refreshGrid
'                    End If
'                Case ""
'                    Exit Sub
'                Case Else
'                    If MsgBox("Input huruf tidak sesuai, akan mengulangi penginputan ?", vbYesNo + vbDefaultButton1 + vbCritical, "Inputan salah") = vbYes Then
'                        GoTo ulangi
'                    Else
'                        Exit Sub
'                    End If
'            End Select
'            Call refreshGrid
'        Else
'            MsgBox "Wo sudah ter-invoice atau belum di validasi. Cek status WO !! ", vbInformation + vbOKOnly, "Validasi"
'        End If
'    Else
'        MsgBox "Sudah terhapus tidak bisa di-Ubah. Pilih ulang WO yang akan diproses.", vbInformation + vbOKOnly, "Validasi"
'    End If
'End Sub

Private Sub listValidasi_FilterChange()
On Error GoTo errHandler
    Set cols = listValidasi.Columns
    Dim C As Integer
    
    C = listValidasi.col
    listValidasi.HoldFields
    rsGrid.filter = getFilter()
    listValidasi.col = C
    listValidasi.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
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

Private Sub Text_GotFocus(Index As Integer)
Dim DiTab As Boolean

    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
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
    Select Case KeyCode
        Case vbKeyF12
            Select Case Index
                Case 0
                    gMainKeyCari = text(Index)
                    gSearchOption = "INV"
                    gBantuan = "AND RPR_COND <> 'K' AND NOINV = '' AND NO_WO NOT IN (SELECT DISTINCT NO_WO FROM VALIDBILLING WITH(NOLOCK) WHERE DELLOG = '')"
                    frmCari.Show (vbModal)
            End Select
        Case vbKeyEscape
            Call hapusForm
    End Select
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsCari As ADODB.Recordset
    Select Case Index
        Case 0
            If Trim(text(0)) = "" Then Exit Sub
            Set rsCari = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(text(Index)) & "' AND NO_WO IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND NO_WO NOT IN (SELECT NO_WO FROM VALIDBILLING WHERE DELLOG = '')")
            If rsCari.EOF Then
                gMainKeyCari = text(Index)
                gSearchOption = "INV"
                gBantuan = "AND RPR_COND <> 'K' AND NOINV = '' AND NO_WO NOT IN (SELECT DISTINCT NO_WO FROM VALIDBILLING WHERE DELLOG = '')"
                frmCari.Show (vbModal)
            Else
                n = Trim(text(0))
                Call hapusForm
                text(0) = Trim(n)
                Call panggilWO(Trim(text(0)))
            End If
    End Select
End Sub

Private Sub panggilWO(noWO As String, Optional tipe As Integer)
Dim rsHlp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
Dim tmp As String
Dim tmp1 As String
Dim tmp2 As String
    
    Set rsHlp = objO.FetchData("SELECT * FROM JOBDTA A INNER JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS WHERE NO_WO = '" & Trim(noWO) & "'")
    If Not rsHlp.EOF Then
        text(1) = Trim(rsHlp("FS_KD_PLG"))
        text(2) = Trim(rsHlp("FS_KD_WIL"))
        text(3) = Trim(rsHlp("CUSTOMER"))
        text(4) = Left(Trim(rsHlp("CUST_ADRES")), 60)
        text(5) = Trim(rsHlp("KOTA"))
        text(6) = Trim(rsHlp("TELPON"))
        text(7) = Trim(rsHlp("GSM"))
        text(8) = Trim(rsHlp("CDMA"))
        text(9) = CDbl(rsHlp("FS_TH_MOBIL"))
        text(10) = Trim(rsHlp("MDL"))
        text(11) = Trim(rsHlp("MODEL"))
        text(12) = Trim(rsHlp("FMDL"))
        text(13) = Trim(rsHlp("COLOR"))
        text(14) = Trim(rsHlp("NO_POLISI"))
        text(15) = Trim(rsHlp("NO_CHASSIS"))
        text(16) = Trim(rsHlp("NO_ENGINE"))
        n = text(17).Mask
        text(17).Mask = ""
        text(17) = Nul(IIf(Replace(rsHlp("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsHlp("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsHlp("DDATE"), "/", "-"))))
        text(17).Mask = n
        text(18) = Trim(rsHlp("MOU"))
        
        ' CHECK DEFAULT DISCOUNT
        Dim statusSBI As Boolean
        Dim nDuoInv As Boolean
        Dim nHarga As Double
        
        statusSBI = False
        Set rsHlp2 = objBilling.takeInvoiceName(rsHlp("NO_CHASSIS"))
        If Not rsHlp2.EOF Then
            nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
        End If
        myChassis = Trim(text(15))
        If Not LblReturn.Visible Then
            If Trim(rsHlp("TPSVC")) = "W" Then          '' WARRANTY
                opTpSvc(2).value = True
                Call giveDiscount("0", text(21), True)
                nHarga = objBilling.takeClassPrice(rsHlp("MODEL"), 3)
                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
            ElseIf Trim(rsHlp("TPSVC")) = "P" Then      '' PROGRAM
                opTpSvc(1).value = True
                Set rsHlp2 = objBilling.searchProgramDetail(rsHlp("KODETIPE"))
                If Not rsHlp2.EOF Then
                    text(19) = Nul(rsHlp2("PROGRAM"))
                    text(19).Visible = True
                End If
                If objBilling.isDuoInvoice(rsHlp("KODETIPE")) Then
                    nDuoInv = True
                    'SBI WO BERKEMBANG
                    If rsHlp("FS_RC") = "R001" Then
                        statusSBI = True
                        Call giveDiscount("100", text(21), False)
                        Select Case nStatusDealer
                            Case 1, 2
                                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            Case 3
                                nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        End Select
                    Else
                        Select Case nStatusDealer
                            Case 1, 2
                                nHarga = objBilling.takeClassPrice(rsHlp("MODEL"), 1)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                            Case 3
                                nHarga = objBilling.takeClassPrice(rsHlp("MODEL"), 2)
                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        End Select
                        Call giveDiscount("100", text(21), False)
                    End If
                Else
                    Call giveDiscount(objBilling.takeDiscountProgram(rsHlp("KODETIPE")), text(21), False)
                    nHarga = objBilling.takeClassPrice(rsHlp("MODEL"), 1)
                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                End If
            ElseIf Trim(rsHlp("TPSVC")) = "" Then
                If lblPDS.Visible Then
                    Call giveDiscount("0", text(21), True)
                    nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                ElseIf rsHlp("FS_RC") = "R001" Then
                    ' HANYA SBI
                    Call giveDiscount("0", text(21), False)
                    Select Case nStatusDealer
                        Case 1, 2
                            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                        Case 3
                            nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                    End Select
                ElseIf objBilling.isMOUCustomer(gKdCbg, rsHlp("MDL"), rsHlp("NO_CHASSIS")) Then
                    Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsHlp("MDL"), rsHlp("NO_CHASSIS")), text(21), False)
                    nHarga = objBilling.takeClassPrice(rsHlp("MODEL"), 1)
                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                Else
                    Call giveDiscount("0", text(21), True)
                    nHarga = objBilling.takeClassPrice(rsHlp("MODEL"), 1)
                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
                End If
            End If
        End If

        
        Dim rsJ As ADODB.Recordset
        For i = 1 To 8
            tmp2 = "KLPKERJA" & i
            tmp = "REP" & i
            tmp1 = "RTE" & i
            If rsHlp(tmp) <> "" Then
               ' If Log.RPLHistori = 0 Then
                    Set rsJ = objO.FetchData("SELECT JOB, FS_ITEM, RATE, PRICE, KETERANGAN FROM HO_FLATE A INNER JOIN MODEL B ON A.MDL = B.MDL " _
                                      & "INNER JOIN CARCLASS C ON C.CLASSCODE = B.MODEL WHERE A.MDL = '" & Trim(rsHlp("MDL")) & "' AND FS_ITEM = '" & Trim(rsHlp(tmp)) & "'")
                'Else
                '    Set rsJ = objO.FetchData("SELECT JOB, FS_ITEM, RATE, PRICE, KETERANGAN FROM [" & dbReplikasiHO & "]..HO_FLATE A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..MODEL B WITH (NOLOCK) ON A.MDL = B.MDL " _
                                      & "INNER JOIN [" & dbReplikasiHO & "]..CARCLASS C WITH (NOLOCK) ON C.CLASSCODE = B.MODEL WHERE A.MDL = '" & Trim(rsHlp("MDL")) & "' AND FS_ITEM = '" & Trim(rsHlp(tmp)) & "'")
                'End If
                If Not rsJ.EOF Then
                    Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                    With rsJob
                        .AddNew
                        !Group = rsJ("JOB")
                        !job = rsJ("FS_ITEM")
                        !Ket = rsJ("KETERANGAN")
                        !Rate = rsHlp(tmp1)
                        !Harga = CDbl(nHarga) * CDbl(!Rate)
                        !hargaBeli = 0
                        If tipe = 1 Then
                            !disc = cariDiskon(Trim(!job), noWO)
                        Else
                            !disc = CDbl(text(21))
                        End If
                        !rp = 0
                        !DISCRP = Round(CDbl(!Harga) * (CDbl(!disc) / 100), 0)
                        !STLDISC = !Harga - !DISCRP
                        .Update
                    End With
                End If
                If rsJ.State <> adStateClosed Then rsJ.Close
            End If
        Next
        
        ' DARI ADDITIONAL JOB
        myChassis = Trim(text(15))
        Dim nDisN As Double
        If CexChassis_Discon(Trim(text(11))) = True Then
            nDisN = (100 - nDiscNRM) / 100
        Else
            nDisN = (100 - nDiscNON) / 100
        End If
        If Log.RPLHistori = 1 Then
            Set rsJ = objO.FetchData("SELECT *, (SELECT PRICE FROM [" & dbReplikasiHO & "]..CARCLASS B INNER JOIN CUSTDAT C ON B.CLASSCODE = C.MODEL WHERE C.NO_CHASSIS = A.NO_CHASSIS)  AS HRG FROM JOBADD2 A WHERE NO_WO = '" & Trim(noWO) & "'")
        Else
            Set rsJ = objO.FetchData("SELECT *, (SELECT PRICE FROM CARCLASS B INNER JOIN CUSTDAT C ON B.CLASSCODE = C.MODEL WHERE C.NO_CHASSIS = A.NO_CHASSIS)  AS HRG FROM JOBADD2 A WHERE NO_WO = '" & Trim(noWO) & "'")
        End If
        ''Set rsJ = objO.FetchData("SELECT *, (SELECT PRICE FROM CARCLASS B INNER JOIN CUSTDAT C ON B.CLASSCODE = C.MODEL WHERE C.NO_CHASSIS = A.NO_CHASSIS)  AS HRG FROM JOBADD2 A WHERE NO_WO = '" & Trim(noWO) & "'")
        If Not rsJ.EOF Then
            Do While Not rsJ.EOF
                Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                With rsJob
                    .AddNew
                    !Group = Trim(rsJ("KLPKERJA"))
                    !job = Trim(rsJ("REP"))
                    !Ket = ""
                    !Rate = CDbl(rsJ("RTE"))
                    If nDisN <> 1 Then
                        !Harga = (CDbl(nHarga) * nDisN) * CDbl(!Rate)
                    Else
                        !Harga = CDbl(nHarga) * CDbl(!Rate)
                    End If
                    !hargaBeli = 0
                    If tipe = 1 Then
                        !disc = cariDiskon(!job, noWO)
                    Else
                        !disc = CDbl(text(21))
                    End If
                    !rp = 0
                    !DISCRP = Round(CDbl(!Harga) * (CDbl(!disc) / 100), 0)
                    !STLDISC = !Harga - !DISCRP
                    .Update
                End With
                rsJ.MoveNext
            Loop
        End If
        
    ' DARI SUBLET
        Set rsJ = objO.FetchData("SELECT * FROM JOBSUBLET WHERE NO_WO = '" & Trim(noWO) & "'")
        Call createTblJobS(IIf(rsJobS.State = adStateClosed, True, False))
        If Not rsJ.EOF Then
            Do While Not rsJ.EOF
                With rsJobS
                    .AddNew
                    !Group = "SUB"
                    !job = Trim(rsJ("NAMABARANG"))
                    !Ket = Trim(rsJ("KODESATUAN"))
                    !Rate = CDbl(rsJ("QTY"))
                    !Harga = CDbl(rsJ("PRICE")) * CDbl(!Rate)
                    !hargaBeli = CDbl(rsJ("HARGA")) * CDbl(!Rate)
                    .Update
                End With
                rsJ.MoveNext
            Loop
        End If

        Set JobGridS.DataSource = rsJobS
        Set JobGrid.DataSource = rsJob
        Call headColumns
        JobGrid.Splits(0).Columns(8).Locked = True
        
        frmeWO.Visible = True
        CmdOpen.SetFocus
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Function giveDiscount(discountValue As Double, discountText As Object, statusText As Boolean)
    
    discountText.Mask = ""
    discountText = Koma2Titik(CDbl(discountValue))
    discountText.Enabled = statusText
    JobGrid.Splits(0).Columns(6).Locked = Not statusText
    JobGrid.Splits(0).Columns(7).Locked = Not statusText
    JobGrid.Splits(0).Columns(8).Locked = Not statusText
    
    cmdSimpan.Enabled = True
    If Not statusText Then
        cmdSimpan.Enabled = False
    End If
End Function

Private Function cariDiskon(deskPekerjaan As String, wo As String) As Double
Dim rsDisc As ADODB.Recordset
    
    cariDiskon = 0
    Set rsDisc = objO.FetchData("Select disc from ValidBilling Where delLog = '' and No_Wo = '" & Trim(wo) & "' and rep = '" & Trim(deskPekerjaan) & "'")
    If Not rsDisc.EOF Then
        cariDiskon = CDbl(rsDisc("Disc"))
    End If
    
End Function

Private Function bolehDiskon(wo As String) As Boolean
    
    bolehDiskon = False
    If Not sudahInvoice(wo) Then
        If sudahValidasi(wo) Then
            bolehDiskon = True
        End If
    End If
    
End Function

Private Function sudahInvoice(wo As String) As Boolean
Dim rsHelp As ADODB.Recordset
    
    sudahInvoice = False
    Set rsHelp = objO.FetchData("Select noinv from jobdta where no_wo = '" & Trim(wo) & "' and deliv <> 'BTL'")
    If Not rsHelp.EOF Then
        If Trim(rsHelp("Noinv")) <> "" Then
            sudahInvoice = True
        End If
    End If
    
End Function

Private Function sudahValidasi(wo As String) As Boolean
Dim rsHelp As ADODB.Recordset
    
    sudahValidasi = False
    Set rsHelp = objO.FetchData("Select userlog from validOrder where no_wo = '" & Trim(wo) & "' and userdel = ''")
    If Not rsHelp.EOF Then
        sudahValidasi = True
    End If
    
End Function


