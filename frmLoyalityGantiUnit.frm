VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLoyalityGantiUnit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganti Pemilik"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   13635
   Begin VB.Frame Frame1 
      Height          =   5445
      Left            =   495
      TabIndex        =   1
      Top             =   615
      Width           =   13035
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H0080C0FF&
         Caption         =   "Tambah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3090
         Width           =   1170
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   735
         Left            =   7425
         Picture         =   "frmLoyalityGantiUnit.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4290
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   8175
         Picture         =   "frmLoyalityGantiUnit.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4305
         Width           =   735
      End
      Begin VB.CommandButton cmdSearch2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Searc&h"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5340
         Picture         =   "frmLoyalityGantiUnit.frx":0D2B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   165
         Width           =   735
      End
      Begin MSMask.MaskEdBox Text 
         DataField       =   "c"
         Height          =   390
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   345
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   25
         Left            =   1020
         TabIndex        =   5
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Left            =   1590
         TabIndex        =   6
         Top             =   1080
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   47
         Left            =   2415
         TabIndex        =   7
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   1020
         TabIndex        =   9
         Top             =   1380
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   23
         Left            =   1020
         TabIndex        =   12
         Top             =   1680
         Width           =   4125
         _ExtentX        =   7276
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
         Index           =   22
         Left            =   4185
         TabIndex        =   11
         Top             =   1365
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   4065
         TabIndex        =   10
         Top             =   1380
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
         Index           =   24
         Left            =   1020
         TabIndex        =   14
         Top             =   1995
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   4065
         TabIndex        =   15
         Top             =   1995
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
         Index           =   20
         Left            =   1020
         TabIndex        =   17
         Top             =   2325
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   6945
         TabIndex        =   21
         Top             =   1380
         Width           =   5985
         _ExtentX        =   10557
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   5
         Left            =   6945
         TabIndex        =   22
         Top             =   1680
         Width           =   5985
         _ExtentX        =   10557
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   8655
         TabIndex        =   23
         Top             =   1080
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
         Left            =   6945
         TabIndex        =   20
         Top             =   1080
         Width           =   1665
         _ExtentX        =   2937
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
         Index           =   56
         Left            =   6945
         TabIndex        =   24
         Top             =   1995
         Width           =   5985
         _ExtentX        =   10557
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   6945
         TabIndex        =   28
         Top             =   2280
         Width           =   3270
         _ExtentX        =   5768
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
         Height          =   270
         Index           =   1
         Left            =   6945
         TabIndex        =   29
         Top             =   2580
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   476
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
         Height          =   270
         Index           =   6
         Left            =   1020
         TabIndex        =   35
         Top             =   3540
         Width           =   5985
         _ExtentX        =   10557
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   7
         Left            =   1020
         TabIndex        =   36
         Top             =   3840
         Width           =   5985
         _ExtentX        =   10557
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   2760
         TabIndex        =   33
         Top             =   3060
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
         Index           =   8
         Left            =   1020
         TabIndex        =   32
         Top             =   3075
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
         Height          =   270
         Index           =   9
         Left            =   1020
         TabIndex        =   37
         Top             =   4155
         Width           =   5985
         _ExtentX        =   10557
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   10
         Left            =   1020
         TabIndex        =   38
         Top             =   4440
         Width           =   3270
         _ExtentX        =   5768
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
         Height          =   270
         Index           =   11
         Left            =   1020
         TabIndex        =   39
         Top             =   4740
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   476
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   345
         Left            =   5205
         TabIndex        =   47
         Top             =   360
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   3465
         TabIndex        =   48
         Top             =   3075
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
         Left            =   2850
         TabIndex        =   49
         Top             =   3120
         Width           =   570
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000000&
         Height          =   1050
         Left            =   7335
         Shape           =   4  'Rounded Rectangle
         Top             =   4140
         Width           =   1680
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H80000000&
         Height          =   510
         Left            =   4155
         Shape           =   4  'Rounded Rectangle
         Top             =   2970
         Width           =   1380
      End
      Begin VB.Label Label12 
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
         Left            =   495
         TabIndex        =   46
         Top             =   3615
         Width           =   405
      End
      Begin VB.Label Label11 
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
         Left            =   405
         TabIndex        =   45
         Top             =   3915
         Width           =   495
      End
      Begin VB.Label Label10 
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
         Left            =   180
         TabIndex        =   44
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   585
         TabIndex        =   43
         Top             =   4785
         Width           =   300
      End
      Begin VB.Label Label7 
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
         Left            =   555
         TabIndex        =   42
         Top             =   4515
         Width           =   330
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   75
         X2              =   12960
         Y1              =   2895
         Y2              =   2895
      End
      Begin VB.Label Label6 
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
         Left            =   6480
         TabIndex        =   31
         Top             =   2355
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   6510
         TabIndex        =   30
         Top             =   2625
         Width           =   300
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
         Left            =   6165
         TabIndex        =   27
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Left            =   6330
         TabIndex        =   26
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label2 
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
         Left            =   6420
         TabIndex        =   25
         Top             =   1455
         Width           =   405
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
         Left            =   255
         TabIndex        =   19
         Top             =   2370
         Width           =   645
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
         Left            =   1650
         TabIndex        =   18
         Top             =   2370
         Width           =   480
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
         Left            =   420
         TabIndex        =   16
         Top             =   2040
         Width           =   480
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
         Left            =   480
         TabIndex        =   13
         Top             =   1425
         Width           =   420
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
         Left            =   255
         TabIndex        =   8
         Top             =   1125
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         X1              =   30
         X2              =   12990
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka : "
         DataField       =   "c"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   585
         TabIndex        =   4
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   ".:: GANTI PEMILIK ::."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   555
      TabIndex        =   0
      Top             =   135
      Width           =   12825
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   12975
   End
   Begin VB.Image Image1 
      Height          =   8205
      Left            =   15
      Picture         =   "frmLoyalityGantiUnit.frx":137A
      Top             =   45
      Width           =   450
   End
End
Attribute VB_Name = "frmLoyalityGantiUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTemp As ADODB.Recordset
Private rsOpsi As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub clearText()
    For i = 0 To 12
        text(i) = ""
    Next
    
    For i = 20 To 25
        text(i) = ""
    Next
    text(46) = ""
    text(47) = ""
    text(56) = ""
End Sub

Private Sub cmdSave_Click()
1        On Error GoTo cmdSave_Click_Error
          Dim isBegin As Boolean
          Dim emailCustomer As String
2         isBegin = False
3         emailCustomer = ""
          
4         If MsgBox("Apakah yakin ingin mengganti kepemilikan unit no rangka " & Trim(text(0)) & " ?", vbYesNo) = vbYes Then
5             Call objHT.OpenTransactionBengkel
              
6             sSQL = "select b.AlamatEmail as email"
7             sSQL = sSQL & " from RplNasApp..NewCustomer a"
8             sSQL = sSQL & " inner join RplNasApp..MasterUser b on a.Email = b.AlamatEmail and b.Verifikasi = 1"
9             sSQL = sSQL & " where a.IDCust = '" & Trim(text(8)) & "'"

10            Set rsTemp = objHT.FetchDataBengkel(sSQL)
11            If Not rsTemp.EOF Then
12                emailCustomer = Nul(rsTemp!Email)
13            End If
              
14            Call objHT.InsertData_Bengkel("UPDATE CUSTDAT SET email = '" & emailCustomer & "',fs_kd_plg = '" & Trim(text(8)) & "',fs_kd_wil = '" & Trim(text(12)) & "',customer = '" & Trim(text(6)) & "',cust_adres = '" & Trim(Nul(text(7)) & " " & Nul(text(9))) & "',telpon = '" & Left(Nul(text(11)), 30) & "' WHERE NO_CHASSIS = '" & Trim(text(0)) & "'")
              
15            Call objHT.InsertData_Bengkel("UPDATE RplNasApp..UnitCustomer SET IDCustService = '" & Trim(text(8)) & "' WHERE NORANGKA = '" & Trim(text(0)) & "'")
              
16            Call objHT.CloseTransactionBengkel
              
17            MsgBox "Data telah diganti", vbInformation
18            Call clearText
19            Exit Sub
20        End If

21       On Error GoTo 0
22       Exit Sub

cmdSave_Click_Error:
23        If isBegin = True Then
24            objHT.CloseTransactionBengkel_err
25        End If
26        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSave_Click of Form frmLoyalityGantiUnit"
End Sub

Private Sub cmdSearch2_Click()
    Call panggilDataNoRangka
End Sub

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
    If KeyCode = 13 Then
        If Index = 0 Then
            Call text_Validate(Index, False)
        End If
    End If
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = text(Index)
                          'gSearchOption = "ORDER"
                gSearchOption = "CUSTOMER2"
                frmCari.Show (vbModal)
            Case 8
                gMainKeyCari = text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If text(Index).BackColor = &HC0E0FF Then KeyAscii = 0
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        Call panggilDataNoRangka
    ElseIf Index = 8 Then
        Call panggilDataLgn
    End If
End Sub

Private Sub panggilDataLgn()
    sSQL = "SELECT * FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & Trim(text(8)) & "' and nonaktifflag = 0"
    Set rsOpsi = objO.FetchDataACC(sSQL)
    If rsOpsi.EOF Then
        MsgBox "Data kodelangganan tidak ditemukan / tidak aktif", vbInformation
        Exit Sub
        
    Else
        text(6) = Replace(Nul(rsOpsi("NAMALGN")), vbNewLine, "")
        text(7) = Nul(rsOpsi("ALAMAT1"))
        text(9) = Nul(rsOpsi("ALAMAT2"))
        text(11) = Trim(Nul(rsOpsi("TELEPON")))
        text(12) = Trim(Nul(rsOpsi!KodeWil))
        If Trim(text(0)) <> "" Then
            cmdSave.Enabled = True
        End If
    End If
End Sub

Private Sub panggilDataNoRangka()
    sSQL = "SELECT a.FS_KD_PLG,'' as NOKTP, a.NPWP ,a.CUSTOMER,'01/01/1900' as TglLahir,a.USEAGAMA as agama,"
    sSQL = sSQL & " a.USEJOBDSC as pekerjaan,'' as propinsi,a.KOTA,a.KECAMATAN,'' as kelurahan,a.KODEPOS,"
    sSQL = sSQL & " '' as rt,'' as rw,a.GSM ,a.USEFS_NM_HOBI as hobi,a.CUST_ADRES,a.USEEMAIL,"
    sSQL = sSQL & " NO_CHASSIS,FS_KD_DEALER,FS_NM_DEALER,ddate,NO_POLISI,MDL,MODEL,FMDL,no_Engine,FS_TH_MOBIL,COLOR,'01/01/1900' as tglawalpolis,'01/01/1900' as tglakhirpolis,'' as namapolis"
    sSQL = sSQL & " FROM custdat a WITH (NOLOCK)"
    sSQL = sSQL & " WHERE a.no_chassis = '" & Trim(text(0)) & "'"
    Set rsTemp = objO.FetchData(sSQL)
    If Not rsTemp.EOF Then
        If Len(Nul(rsTemp("NO_POLISI"))) > 0 Then
            'nPol = Nul(rsCust("NO_POLISI"))
            Dim nDpn1 As String
            Dim nTgh1 As String
            Dim nBlk1 As String
            nDpn1 = ""
            nTgh1 = ""
            nBlk1 = ""
            Call pecahNoPol(Trim(Nul(rsTemp("NO_POLISI"))), nDpn1, nTgh1, nBlk1)
            text(25) = Trim(nDpn1)
            text(46) = Trim(nTgh1)
            text(47) = Trim(nBlk1)
        Else
            text(25) = ""
            text(46) = ""
            text(47) = ""
        End If
        
        text(20) = Nul(rsTemp("FS_TH_MOBIL"))
        text(21) = Nul(rsTemp("MDL"))
        text(22) = Nul(rsTemp("MODEL"))
        text(23) = Nul(rsTemp("FMDL"))
        text(24) = Nul(rsTemp("COLOR"))
        
        text(2) = Nul(rsTemp("FS_KD_PLG"))
        text(4) = Left(Nul(rsTemp("CUSTOMER")), 60)
        text(5) = Left(Nul(rsTemp("CUST_ADRES")), 60)
        If Len(Trim(Nul(rsTemp("CUST_ADRES")))) > 60 Then
            text(56) = Mid(Nul(rsTemp("CUST_ADRES")), 61, Len(Nul(rsTemp("CUST_ADRES"))) - 60)
        End If
        text(3) = Nul(rsTemp("KOTA"))
       ' cmdSave.Enabled = True
    Else
       MsgBox "Tidak ada data no rangka " & Trim(text(0)) & " pada system", vbInformation
       Exit Sub
    End If
End Sub
