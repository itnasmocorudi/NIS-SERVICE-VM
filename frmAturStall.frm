VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmAturStall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Atur Jam Kerja"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmAturStall.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   28
      Top             =   -5640
      Width           =   495
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   915
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   915
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   4
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   6
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   7
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   8
      Left            =   1680
      TabIndex        =   8
      Top             =   2400
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   9
      Left            =   2880
      TabIndex        =   9
      Top             =   2400
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   10
      Left            =   1680
      TabIndex        =   10
      Top             =   2760
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   11
      Left            =   2880
      TabIndex        =   11
      Top             =   2760
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   12
      Left            =   1680
      TabIndex        =   12
      Top             =   3120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   13
      Left            =   2880
      TabIndex        =   13
      Top             =   3120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin ucXPButton.XPButton cmdSimpan 
      Height          =   495
      Left            =   3360
      TabIndex        =   30
      Top             =   3600
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
      MICON           =   "frmAturStall.frx":14D9
      PICN            =   "frmAturStall.frx":14F5
      PICH            =   "frmAturStall.frx":1A08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdBatal 
      Height          =   495
      Left            =   4560
      TabIndex        =   31
      Top             =   3600
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
      MICON           =   "frmAturStall.frx":1FAA
      PICN            =   "frmAturStall.frx":1FC6
      PICH            =   "frmAturStall.frx":249A
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
      Index           =   14
      Left            =   6150
      TabIndex        =   32
      Top             =   840
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   15
      Left            =   7350
      TabIndex        =   33
      Top             =   840
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   16
      Left            =   6150
      TabIndex        =   34
      Top             =   1245
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   17
      Left            =   7350
      TabIndex        =   35
      Top             =   1245
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   18
      Left            =   6150
      TabIndex        =   36
      Top             =   1605
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   19
      Left            =   7350
      TabIndex        =   37
      Top             =   1605
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   20
      Left            =   6150
      TabIndex        =   38
      Top             =   1965
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   21
      Left            =   7350
      TabIndex        =   39
      Top             =   1965
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   22
      Left            =   6150
      TabIndex        =   40
      Top             =   2325
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   23
      Left            =   7350
      TabIndex        =   41
      Top             =   2325
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   24
      Left            =   6150
      TabIndex        =   42
      Top             =   2685
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   25
      Left            =   7350
      TabIndex        =   43
      Top             =   2685
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   26
      Left            =   6150
      TabIndex        =   44
      Top             =   3045
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   27
      Left            =   7350
      TabIndex        =   45
      Top             =   3045
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderStyle     =   6  'Inside Solid
      X1              =   8640
      X2              =   8640
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: ATUR JAM ISTIRAHAT  ::."
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
      Index           =   0
      Left            =   4845
      TabIndex        =   60
      Top             =   315
      Width           =   3645
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Senin"
      Height          =   195
      Index           =   13
      Left            =   5520
      TabIndex        =   59
      Top             =   885
      Width           =   405
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   58
      Top             =   885
      Width           =   240
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   57
      Top             =   1290
      Width           =   240
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   56
      Top             =   1650
      Width           =   240
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   55
      Top             =   2010
      Width           =   240
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   54
      Top             =   2370
      Width           =   240
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   53
      Top             =   2730
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   6990
      TabIndex        =   52
      Top             =   3090
      Width           =   240
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Selasa"
      Height          =   195
      Index           =   12
      Left            =   5445
      TabIndex        =   51
      Top             =   1290
      Width           =   480
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rabu"
      Height          =   195
      Index           =   11
      Left            =   5535
      TabIndex        =   50
      Top             =   1650
      Width           =   390
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kamis"
      Height          =   195
      Index           =   10
      Left            =   5505
      TabIndex        =   49
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumat"
      Height          =   195
      Index           =   9
      Left            =   5505
      TabIndex        =   48
      Top             =   2370
      Width           =   420
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sabtu"
      Height          =   195
      Index           =   8
      Left            =   5505
      TabIndex        =   47
      Top             =   2730
      Width           =   420
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Minggu"
      Height          =   195
      Index           =   7
      Left            =   5400
      TabIndex        =   46
      Top             =   3090
      Width           =   525
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   4560
      X2              =   4560
      Y1              =   240
      Y2              =   3360
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   600
      X2              =   8520
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   600
      X2              =   8520
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: ATUR JAM OPERASIONAL ::."
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
      Index           =   1
      Left            =   885
      TabIndex        =   29
      Top             =   300
      Width           =   3225
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Minggu"
      Height          =   195
      Index           =   6
      Left            =   930
      TabIndex        =   27
      Top             =   3165
      Width           =   525
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sabtu"
      Height          =   195
      Index           =   5
      Left            =   1035
      TabIndex        =   26
      Top             =   2805
      Width           =   420
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumat"
      Height          =   195
      Index           =   4
      Left            =   1035
      TabIndex        =   25
      Top             =   2445
      Width           =   420
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Kamis"
      Height          =   195
      Index           =   3
      Left            =   1035
      TabIndex        =   24
      Top             =   2085
      Width           =   420
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rabu"
      Height          =   195
      Index           =   2
      Left            =   1065
      TabIndex        =   23
      Top             =   1725
      Width           =   390
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Selasa"
      Height          =   195
      Index           =   1
      Left            =   975
      TabIndex        =   22
      Top             =   1365
      Width           =   480
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   21
      Top             =   3165
      Width           =   240
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   20
      Top             =   2805
      Width           =   240
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   19
      Top             =   2445
      Width           =   240
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   18
      Top             =   2085
      Width           =   240
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   17
      Top             =   1725
      Width           =   240
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   1365
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      Height          =   195
      Left            =   2520
      TabIndex        =   15
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Senin"
      Height          =   195
      Index           =   0
      Left            =   1050
      TabIndex        =   14
      Top             =   960
      Width           =   405
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmAturStall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsStall As New ADODB.Recordset

Private Sub cmdBatal_Click()
    Unload Me
    'frmtprofil.Show
End Sub

Private Sub hapusForm()
    For i = 0 To 27
        Text(i).Mask = ""
        Text(i).Text = "00:00"
        Text(i).Mask = "99:99"
    Next
End Sub

Private Sub cmdSimpan_Click()
      Dim i As Integer
      Dim j As Integer
      Dim stall As Integer
      Dim boleh As Boolean

1        On Error GoTo cmdSimpan_Click_Error

2         boleh = True
3         For i = 0 To 27
4             If Trim(Text(i)) = "" Then
5                 boleh = False
6                 Exit For
7             End If
8         Next
          
9         If boleh Then   'And Me.Tag = "Stall"
10            Call objO.UpdateData("UPDATE T_STALL SET WAKTUMULAI = '', WAKTUSELESAI = '', WAKTUMULAIT = '', WAKTUSELESAIT = ''")
11            stall = 0
12            j = 0
13            For i = 0 To 13
14                If i Mod 2 = 0 Then
15                    Call objO.UpdateData("UPDATE T_STALL SET WAKTUMULAI = '" & Trim(Text(i)) & "', WAKTUSELESAI = '" & Trim(Text(i + 1)) & "', WAKTUMULAIT = '" & Trim(Text(i)) & "', WAKTUSELESAIT = '" & Trim(Text(i + 1)) & "' WHERE HARI = '" & j & Trim(Label(j)) & "'")
16                    j = j + 1
17                End If
18            Next
19            stall = 0
20            j = 0
21            If Log.newJpcb = 1 Then
22                For i = 14 To 27
23                    If i Mod 2 = 0 Then
24                        Call objO.UpdateData("UPDATE T_STALL SET WAKTUMULAIistirahat = '" & Trim(Text(i)) & "', WAKTUSELESAIistirahat = '" & Trim(Text(i + 1)) & "' WHERE HARI = '" & j & Trim(Label(j)) & "'")
25                        j = j + 1
26                    End If
27                Next
28            End If
29            Unload Me
30        Else
31            MsgBox "Waktu masih ada yang belum lengkap"
32            Exit Sub
33        End If

34       On Error GoTo 0
35       Exit Sub

cmdSimpan_Click_Error:

36        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmAturStall")
End Sub

Private Sub Form_Load()
    If Log.newJpcb = 0 Then
        Me.Width = 4740
        cmdSimpan.Left = 4740 - cmdSimpan.Width - cmdBatal.Width - 200
        cmdBatal.Left = 4740 - cmdBatal.Width - 200
    End If
    Call hapusForm
    Call isiForm
End Sub

Private Function isiForm()
    Dim i As Integer
    Dim j As Integer
    Dim iM As Integer
    
    i = 0
    j = 0
    
    iM = 14
'    If frmtprofil.Tag = "Stall" Then
    If Log.newJpcb = 1 Then
        frmAturStall.Caption = "Atur Jam Kerja Stall"
        Set rsStall = objO.FetchData("SELECT distinct hari, waktuMulai, waktuSelesai,waktuMulaiIstirahat,waktuSelesaiIstirahat FROM T_STALL WITH (NOLOCK) ORDER BY hari")
        If rsStall.RecordCount > 0 Then
            Do While Not rsStall.EOF
                If Right(Trim(rsStall!hari), Len(Trim(rsStall!hari)) - 1) = Label(j) Then
                    Text(i).Mask = ""
                    Text(i + 1).Mask = ""
                    Text(i) = Trim(rsStall!waktuMulai)
                    Text(i + 1) = Trim(rsStall!waktuSelesai)
                    Text(i).Mask = "##:##"
                    Text(i + 1).Mask = "##:##"
                    
                    Text(iM).Mask = ""
                    Text(iM + 1).Mask = ""
                    Text(iM) = Trim(rsStall!waktumulaiistirahat)
                    Text(iM + 1) = Trim(rsStall!waktuselesaiistirahat)
                    Text(iM).Mask = "##:##"
                    Text(iM + 1).Mask = "##:##"
                                        
                    i = i + 2
                    iM = iM + 2
                    j = j + 1
                End If
                rsStall.MoveNext
            Loop
        End If
    Else
        frmAturStall.Caption = "Atur Jam Kerja Stall"
        Set rsStall = objO.FetchData("SELECT distinct hari, waktuMulai, waktuSelesai FROM T_STALL WITH (NOLOCK) ORDER BY hari")
        If rsStall.RecordCount > 0 Then
            Do While Not rsStall.EOF
                If Right(Trim(rsStall!hari), Len(Trim(rsStall!hari)) - 1) = Label(j) Then
                    Text(i).Mask = ""
                    Text(i + 1).Mask = ""
                    Text(i) = Trim(rsStall!waktuMulai)
                    Text(i + 1) = Trim(rsStall!waktuSelesai)
                    Text(i).Mask = "##:##"
                    Text(i + 1).Mask = "##:##"
                    i = i + 2
                    j = j + 1
                End If
                rsStall.MoveNext
            Loop
        End If
    End If
End Function

Private Sub text_GotFocus(Index As Integer)
    Text(Index).SelLength = Len(Text(Index))
    Text(Index).SetFocus
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
1        On Error GoTo text_Validate_Error

2         If Index Mod 2 = 0 Then
              ' jam mulai
3             If Not CekJam(Text(Index)) Then
4                 Text(Index) = "00:00"
5             ElseIf InStr(Text(Index), " :  ") > 0 Then
6                 Text(Index) = Replace(Replace(Text(Index), " ", ""), ":", "") & "0:00"
7             ElseIf InStr(Text(Index), " : ") > 0 Then
8                 Text(Index) = "0" & Left(Replace(Text(Index), " ", ""), 1) & ":0" & Right(Replace(Text(Index), " ", ""), 1)
9             ElseIf InStr(Text(Index), " :") > 0 Then
10                Text(Index) = "0" & Replace(Text(Index), " ", "")
11            End If
12        Else
              ' jam akhir
13            If Not CekJam(Text(Index)) Then
14                Text(Index) = "23:59"
15            ElseIf InStr(Text(Index), " :  ") > 0 Then
16                Text(Index) = Replace(Replace(Text(Index), " ", ""), ":", "") & "0:00"
17            ElseIf InStr(Text(Index), " : ") > 0 Then
18                Text(Index) = "0" & Left(Replace(Text(Index), " ", ""), 1) & ":0" & Right(Replace(Text(Index), " ", ""), 1)
19            ElseIf InStr(Text(Index), " :") > 0 Then
20                Text(Index) = "0" & Replace(Text(Index), " ", "")
21            End If

              ' jam awal > jam akhir
22            If Replace(Text(Index), " ", "") < Text(Index - 1) Then
23                MsgBox "Waktu akhir tidak bisa lebih kecil dari waktu awal"
24                Text(Index) = "23:59"
25                Text(Index).SelLength = Len(Text(Index))
26                Text(Index).SetFocus
27            End If
28        End If

29       On Error GoTo 0
30       Exit Sub

text_Validate_Error:

31        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmAturStall")
End Sub
