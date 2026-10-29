VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmtAkun 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Setting Kode Jurnal"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9420
      Left            =   480
      TabIndex        =   52
      Top             =   240
      Width           =   13620
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sublet"
         Height          =   855
         Left            =   120
         TabIndex        =   127
         Top             =   8520
         Width           =   13395
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   65
            Left            =   1470
            TabIndex        =   48
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   67
            Left            =   2565
            TabIndex        =   128
            Top             =   180
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   68
            Left            =   8190
            TabIndex        =   49
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   69
            Left            =   9285
            TabIndex        =   129
            Top             =   180
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   100
            Left            =   1470
            TabIndex        =   164
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   101
            Left            =   2565
            TabIndex        =   165
            Top             =   480
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   9
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Left            =   960
            TabIndex        =   166
            Top             =   525
            Width           =   330
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   25
            Left            =   9180
            Top             =   180
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Hutang"
            Height          =   195
            Index           =   16
            Left            =   7470
            TabIndex        =   131
            Top             =   218
            Width           =   525
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PDP"
            Height          =   195
            Left            =   960
            TabIndex        =   130
            Top             =   218
            Width           =   330
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   8
            Left            =   2460
            Top             =   180
            Width           =   60
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Affiliasi PPN"
         Height          =   2085
         Left            =   120
         TabIndex        =   109
         Top             =   6495
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   48
            Left            =   1470
            TabIndex        =   18
            Top             =   165
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   49
            Left            =   1470
            TabIndex        =   20
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   50
            Left            =   2565
            TabIndex        =   110
            Top             =   180
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   51
            Left            =   2565
            TabIndex        =   111
            Top             =   780
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   52
            Left            =   1470
            TabIndex        =   22
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   53
            Left            =   2565
            TabIndex        =   112
            Top             =   1380
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   54
            Left            =   1470
            TabIndex        =   21
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   55
            Left            =   2565
            TabIndex        =   113
            Top             =   1080
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   76
            Left            =   1470
            TabIndex        =   19
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   77
            Left            =   2565
            TabIndex        =   138
            Top             =   480
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   94
            Left            =   1470
            TabIndex        =   23
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   95
            Left            =   2565
            TabIndex        =   158
            Top             =   1695
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   38
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   159
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   29
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   840
            TabIndex        =   139
            Top             =   480
            Width           =   450
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   6
            Left            =   2460
            Top             =   780
            Width           =   60
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   117
            Top             =   780
            Width           =   525
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   20
            Left            =   2460
            Top             =   165
            Width           =   60
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   13
            Left            =   945
            TabIndex        =   115
            Top             =   1080
            Width           =   330
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   19
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   18
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   12
            Left            =   735
            TabIndex        =   114
            Top             =   1380
            Width           =   540
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Affiliasi PPN"
         Height          =   2085
         Left            =   6840
         TabIndex        =   100
         Top             =   6495
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   40
            Left            =   1470
            TabIndex        =   42
            Top             =   165
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   41
            Left            =   1470
            TabIndex        =   44
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   42
            Left            =   2565
            TabIndex        =   101
            Top             =   180
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   43
            Left            =   2565
            TabIndex        =   102
            Top             =   780
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   44
            Left            =   1470
            TabIndex        =   46
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   45
            Left            =   2565
            TabIndex        =   103
            Top             =   1380
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   46
            Left            =   1470
            TabIndex        =   45
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   47
            Left            =   2565
            TabIndex        =   104
            Top             =   1080
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   84
            Left            =   1470
            TabIndex        =   43
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   85
            Left            =   2565
            TabIndex        =   146
            Top             =   480
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   96
            Left            =   1470
            TabIndex        =   47
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   97
            Left            =   2565
            TabIndex        =   160
            Top             =   1695
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   39
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   161
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   33
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   855
            TabIndex        =   147
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   11
            Left            =   765
            TabIndex        =   108
            Top             =   1380
            Width           =   540
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   17
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   16
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   10
            Left            =   975
            TabIndex        =   107
            Top             =   1080
            Width           =   330
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   165
            TabIndex        =   106
            Top             =   240
            Width           =   1140
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   15
            Left            =   2460
            Top             =   165
            Width           =   60
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   780
            TabIndex        =   105
            Top             =   780
            Width           =   525
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   5
            Left            =   2460
            Top             =   780
            Width           =   60
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Affiliasi Non PPN"
         Height          =   2115
         Left            =   120
         TabIndex        =   118
         Top             =   4440
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   56
            Left            =   1470
            TabIndex        =   14
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   57
            Left            =   1470
            TabIndex        =   16
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   58
            Left            =   1470
            TabIndex        =   12
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   59
            Left            =   2565
            TabIndex        =   119
            Top             =   180
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   60
            Left            =   2565
            TabIndex        =   120
            Top             =   780
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   61
            Left            =   2565
            TabIndex        =   121
            Top             =   1380
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   62
            Left            =   1470
            TabIndex        =   15
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   63
            Left            =   2565
            TabIndex        =   122
            Top             =   1080
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   74
            Left            =   1470
            TabIndex        =   13
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   75
            Left            =   2565
            TabIndex        =   136
            Top             =   480
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   92
            Left            =   1470
            TabIndex        =   17
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   93
            Left            =   2565
            TabIndex        =   156
            Top             =   1695
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   37
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   157
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   28
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   840
            TabIndex        =   137
            Top             =   525
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   15
            Left            =   945
            TabIndex        =   126
            Top             =   1080
            Width           =   330
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   23
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   22
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   21
            Left            =   2460
            Top             =   780
            Width           =   60
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   1140
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   7
            Left            =   2460
            Top             =   180
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   14
            Left            =   735
            TabIndex        =   124
            Top             =   1380
            Width           =   540
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   123
            Top             =   780
            Width           =   525
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Affiliasi Non PPN "
         Height          =   2085
         Left            =   6840
         TabIndex        =   91
         Top             =   4440
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   32
            Left            =   1470
            TabIndex        =   38
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   33
            Left            =   1470
            TabIndex        =   40
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   34
            Left            =   1470
            TabIndex        =   36
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   35
            Left            =   2565
            TabIndex        =   92
            Top             =   180
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   36
            Left            =   2565
            TabIndex        =   93
            Top             =   780
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   37
            Left            =   2565
            TabIndex        =   94
            Top             =   1380
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   38
            Left            =   1470
            TabIndex        =   39
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   39
            Left            =   2565
            TabIndex        =   95
            Top             =   1080
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   82
            Left            =   1470
            TabIndex        =   37
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   83
            Left            =   2565
            TabIndex        =   144
            Top             =   480
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   98
            Left            =   1470
            TabIndex        =   41
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   99
            Left            =   2565
            TabIndex        =   162
            Top             =   1695
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   40
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   163
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   32
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   840
            TabIndex        =   145
            Top             =   518
            Width           =   450
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   99
            Top             =   780
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   9
            Left            =   735
            TabIndex        =   98
            Top             =   1380
            Width           =   540
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   4
            Left            =   2460
            Top             =   180
            Width           =   60
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   150
            TabIndex        =   97
            Top             =   255
            Width           =   1140
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   14
            Left            =   2460
            Top             =   780
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   13
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   12
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   8
            Left            =   945
            TabIndex        =   96
            Top             =   1080
            Width           =   330
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Intern"
         Height          =   2085
         Left            =   6840
         TabIndex        =   80
         Top             =   2340
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   24
            Left            =   1470
            TabIndex        =   32
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   25
            Left            =   1470
            TabIndex        =   34
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   26
            Left            =   1470
            TabIndex        =   30
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   27
            Left            =   2565
            TabIndex        =   81
            Top             =   180
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   28
            Left            =   2565
            TabIndex        =   82
            Top             =   780
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   29
            Left            =   2565
            TabIndex        =   83
            Top             =   1380
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   30
            Left            =   1470
            TabIndex        =   33
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   31
            Left            =   2565
            TabIndex        =   84
            Top             =   1080
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   78
            Left            =   1470
            TabIndex        =   31
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   79
            Left            =   2565
            TabIndex        =   140
            Top             =   480
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   90
            Left            =   1470
            TabIndex        =   35
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   91
            Left            =   2565
            TabIndex        =   154
            Top             =   1695
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   36
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   155
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   30
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   840
            TabIndex        =   141
            Top             =   518
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   7
            Left            =   945
            TabIndex        =   88
            Top             =   1080
            Width           =   330
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   11
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   10
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   9
            Left            =   2460
            Top             =   780
            Width           =   60
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   150
            TabIndex        =   87
            Top             =   255
            Width           =   1140
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   3
            Left            =   2460
            Top             =   180
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   6
            Left            =   735
            TabIndex        =   86
            Top             =   1380
            Width           =   540
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   85
            Top             =   780
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Intern"
         Height          =   2085
         Left            =   120
         TabIndex        =   53
         Top             =   2340
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   5
            Left            =   1470
            TabIndex        =   8
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   6
            Left            =   1470
            TabIndex        =   10
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   0
            Left            =   1470
            TabIndex        =   6
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   1
            Left            =   2565
            TabIndex        =   54
            Top             =   180
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   7
            Left            =   2565
            TabIndex        =   55
            Top             =   780
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   9
            Left            =   2565
            TabIndex        =   56
            Top             =   1380
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   10
            Left            =   1470
            TabIndex        =   9
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   11
            Left            =   2565
            TabIndex        =   57
            Top             =   1080
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   70
            Left            =   1470
            TabIndex        =   7
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   71
            Left            =   2565
            TabIndex        =   132
            Top             =   480
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   88
            Left            =   1470
            TabIndex        =   11
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   89
            Left            =   2565
            TabIndex        =   152
            Top             =   1695
            Width           =   3930
            _ExtentX        =   6932
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   35
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   153
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   26
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   795
            TabIndex        =   133
            Top             =   525
            Width           =   450
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   61
            Top             =   825
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   60
            Top             =   1425
            Width           =   540
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   0
            Left            =   2460
            Top             =   180
            Width           =   60
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   150
            TabIndex        =   59
            Top             =   255
            Width           =   1140
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   1
            Left            =   2460
            Top             =   780
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   2
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   3
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   1
            Left            =   945
            TabIndex        =   58
            Top             =   1125
            Width           =   330
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Extern"
         Height          =   2085
         Left            =   6840
         TabIndex        =   71
         Top             =   315
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   16
            Left            =   1470
            TabIndex        =   24
            Top             =   165
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   17
            Left            =   1470
            TabIndex        =   26
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   18
            Left            =   2565
            TabIndex        =   72
            Top             =   180
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   19
            Left            =   2565
            TabIndex        =   73
            Top             =   780
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   20
            Left            =   1470
            TabIndex        =   28
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   21
            Left            =   2565
            TabIndex        =   74
            Top             =   1380
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   22
            Left            =   1470
            TabIndex        =   27
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   23
            Left            =   2565
            TabIndex        =   75
            Top             =   1080
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   80
            Left            =   1470
            TabIndex        =   25
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   81
            Left            =   2565
            TabIndex        =   142
            Top             =   480
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   86
            Left            =   1470
            TabIndex        =   29
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   87
            Left            =   2565
            TabIndex        =   150
            Top             =   1695
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   34
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   480
            TabIndex        =   151
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   31
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   840
            TabIndex        =   143
            Top             =   480
            Width           =   450
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   2
            Left            =   2460
            Top             =   780
            Width           =   60
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   79
            Top             =   780
            Width           =   525
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   8
            Left            =   2460
            Top             =   165
            Width           =   60
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   150
            TabIndex        =   78
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   5
            Left            =   945
            TabIndex        =   77
            Top             =   1080
            Width           =   330
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   7
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   6
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   4
            Left            =   735
            TabIndex        =   76
            Top             =   1380
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kendaraan Extern"
         Height          =   2085
         Left            =   120
         TabIndex        =   62
         Top             =   315
         Width           =   6675
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   2
            Left            =   1470
            TabIndex        =   0
            Top             =   165
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   4
            Left            =   1470
            TabIndex        =   2
            Top             =   780
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   3
            Left            =   2565
            TabIndex        =   63
            Top             =   180
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   8
            Left            =   2565
            TabIndex        =   64
            Top             =   780
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   12
            Left            =   1470
            TabIndex        =   4
            Top             =   1380
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   13
            Left            =   2565
            TabIndex        =   67
            Top             =   1380
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   14
            Left            =   1470
            TabIndex        =   3
            Top             =   1080
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   15
            Left            =   2565
            TabIndex        =   68
            Top             =   1080
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   72
            Left            =   1470
            TabIndex        =   1
            Top             =   480
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   73
            Left            =   2565
            TabIndex        =   134
            Top             =   480
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   64
            Left            =   1470
            TabIndex        =   5
            Top             =   1695
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Mask            =   ">&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   66
            Left            =   2565
            TabIndex        =   148
            Top             =   1695
            Width           =   3915
            _ExtentX        =   6906
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
            Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            PromptChar      =   ">"
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   24
            Left            =   2460
            Top             =   1680
            Width           =   60
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "HPP Sublet"
            Height          =   195
            Left            =   450
            TabIndex        =   149
            Top             =   1755
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   27
            Left            =   2460
            Top             =   480
            Width           =   60
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sublet"
            Height          =   195
            Left            =   840
            TabIndex        =   135
            Top             =   525
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Piutang"
            Height          =   195
            Index           =   3
            Left            =   735
            TabIndex        =   70
            Top             =   1418
            Width           =   540
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   5
            Left            =   2460
            Top             =   1380
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   4
            Left            =   2460
            Top             =   1080
            Width           =   60
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "PPN"
            Height          =   195
            Index           =   2
            Left            =   945
            TabIndex        =   69
            Top             =   1118
            Width           =   330
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour, Material"
            Height          =   195
            Left            =   150
            TabIndex        =   66
            Top             =   240
            Width           =   1140
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   0
            Left            =   2460
            Top             =   165
            Width           =   60
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Meterai"
            Height          =   195
            Left            =   765
            TabIndex        =   65
            Top             =   818
            Width           =   525
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Index           =   1
            Left            =   2460
            Top             =   780
            Width           =   60
         End
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Jurnal Retur Invoice"
         Height          =   255
         Left            =   6840
         TabIndex        =   90
         Top             =   75
         Width           =   2940
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Jurnal Pembuatan Invoice"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   75
         Width           =   2940
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   0
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   11685
      TabIndex        =   51
      Top             =   9735
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   661
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   375
      Left            =   10470
      TabIndex        =   50
      Top             =   9735
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9735
      Left            =   0
      Picture         =   "frmtAkun.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmtAkun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset
    
    Call ClearsSql
    gsSql1(0) = "FS_KD_LABOUR_INT"
    gsSql2(0) = text(0)
    gsSql1(1) = "FS_KD_METERAI_INT"
    gsSql2(1) = text(5)
    gsSql1(2) = "FS_KD_PPN_INT"
    gsSql2(2) = text(10)
    gsSql1(3) = "FS_KD_PIUTANG_INT"
    gsSql2(3) = text(6)
    gsSql1(4) = "FS_KD_LABOUR_EXT"
    gsSql2(4) = text(2)
    gsSql1(5) = "FS_KD_METERAI_EXT"
    gsSql2(5) = text(4)
    gsSql1(6) = "FS_KD_PPN_EXT"
    gsSql2(6) = text(14)
    gsSql1(7) = "FS_KD_PIUTANG_EXT"
    gsSql2(7) = text(12)
    gsSql1(8) = "FS_KD_LABOUR_AFFNONPPN"
    gsSql2(8) = text(58)
    gsSql1(9) = "FS_KD_METERAI_AFFNONPPN"
    gsSql2(9) = text(56)
    gsSql1(10) = "FS_KD_PIUTANG_AFFNONPPN"
    gsSql2(10) = text(57)
    gsSql1(11) = "FS_KD_PPN_AFFNONPPN"
    gsSql2(11) = text(62)
    gsSql1(12) = "FS_KD_LABOUR_AFFPPN"
    gsSql2(12) = text(48)
    gsSql1(13) = "FS_KD_METERAI_AFFPPN"
    gsSql2(13) = text(49)
    gsSql1(14) = "FS_KD_PIUTANG_AFFPPN"
    gsSql2(14) = text(52)
    gsSql1(15) = "FS_KD_PPN_AFFPPN"
    gsSql2(15) = text(54)
    
    ' RETUR
    gsSql1(16) = "FS_KD_RLABOUR_INT"
    gsSql2(16) = text(26)
    gsSql1(17) = "FS_KD_RMETERAI_INT"
    gsSql2(17) = text(24)
    gsSql1(18) = "FS_KD_RPIUTANG_INT"
    gsSql2(18) = text(25)
    gsSql1(19) = "FS_KD_RPPN_INT"
    gsSql2(19) = text(30)
    gsSql1(20) = "FS_KD_RLABOUR_EXT"
    gsSql2(20) = text(16)
    gsSql1(21) = "FS_KD_RMETERAI_EXT"
    gsSql2(21) = text(17)
    gsSql1(22) = "FS_KD_RPIUTANG_EXT"
    gsSql2(22) = text(20)
    gsSql1(23) = "FS_KD_RPPN_EXT"
    gsSql2(23) = text(22)
    
    gsSql1(24) = "FS_KD_RLABOUR_AFFNONPPN"
    gsSql2(24) = text(34)
    gsSql1(25) = "FS_KD_RMETERAI_AFFNONPPN"
    gsSql2(25) = text(32)
    gsSql1(26) = "FS_KD_RPIUTANG_AFFNONPPN"
    gsSql2(26) = text(33)
    gsSql1(27) = "FS_KD_RPPN_AFFNONPPN"
    gsSql2(27) = text(38)
    gsSql1(28) = "FS_KD_RLABOUR_AFFPPN"
    gsSql2(28) = text(40)
    gsSql1(29) = "FS_KD_RMETERAI_AFFPPN"
    gsSql2(29) = text(41)
    gsSql1(30) = "FS_KD_RPIUTANG_AFFPPN"
    gsSql2(30) = text(44)
    gsSql1(31) = "FS_KD_RPPN_AFFPPN"
    gsSql2(31) = text(46)
    
    gsSql1(32) = "FS_KD_PDP"
    gsSql2(32) = text(65)
    gsSql1(33) = "FS_KD_HUTANG"
    gsSql2(33) = text(68)
    
    gsSql1(34) = "FS_KD_SUBLET_INT"
    gsSql2(34) = text(70)
    gsSql1(35) = "FS_KD_SUBLET_EXT"
    gsSql2(35) = text(72)
    gsSql1(36) = "FS_KD_SUBLET_AFFNONPPN"
    gsSql2(36) = text(74)
    gsSql1(37) = "FS_KD_SUBLET_AFFPPN"
    gsSql2(37) = text(76)
    gsSql1(38) = "FS_KD_RSUBLET_INT"
    gsSql2(38) = text(78)
    gsSql1(39) = "FS_KD_RSUBLET_EXT"
    gsSql2(39) = text(80)
    gsSql1(40) = "FS_KD_RSUBLET_AFFNONPPN"
    gsSql2(40) = text(82)
    gsSql1(41) = "FS_KD_RSUBLET_AFFPPN"
    gsSql2(41) = text(84)
    
    gsSql1(42) = "FS_KD_HPP_EXT"
    gsSql2(42) = text(64)
    gsSql1(43) = "FS_KD_HPP_INT"
    gsSql2(43) = text(88)
    gsSql1(44) = "FS_KD_HPP_AFFNONPPN"
    gsSql2(44) = text(92)
    gsSql1(45) = "FS_KD_HPP_AFFPPN"
    gsSql2(45) = text(94)
    
    gsSql1(46) = "FS_KD_RHPP_EXT"
    gsSql2(46) = text(86)
    gsSql1(47) = "FS_KD_RHPP_INT"
    gsSql2(47) = text(90)
    gsSql1(48) = "FS_KD_RHPP_AFFNONPPN"
    gsSql2(48) = text(98)
    gsSql1(49) = "FS_KD_RHPP_AFFPPN"
    gsSql2(49) = text(96)
    
    gsSql1(50) = "FS_KD_PPN"
    gsSql2(50) = text(100)
    
    Set rsSearch = objO.FetchData("Select * from T_AKUN")
    If Not rsSearch.BOF Then
        sSQL = GensSql("T_AKUN", True)
    Else
        sSQL = GensSql("T_AKUN", False)
    End If
    Call objO.UpdateData(sSQL)
    
    gsSql1(51) = "userEntry"
    gsSql2(51) = gSA
    gsSql1(52) = "tglEntry"
    gsSql2(52) = Format(isnow, "MM/DD/YYYY HH:NN:SS")
    gsSql1(53) = "clientName"
    gsSql2(53) = ComptName
    
    Call objO.UpdateData(sSQL)
    text(0).SetFocus
End Sub

Private Sub Form_Load()
Dim rsSearch As ADODB.Recordset

    Call hapusForm
    
    Set rsSearch = objO.FetchData("Select * from T_AKUN")
    If Not rsSearch.BOF Then
        text(0) = rsSearch("FS_KD_LABOUR_INT")
        text(70) = rsSearch("FS_KD_SUBLET_INT")
        text(5) = rsSearch("FS_KD_METERAI_INT")
        text(10) = rsSearch("FS_KD_PPN_INT")
        text(6) = rsSearch("FS_KD_PIUTANG_INT")
        text(88) = rsSearch("FS_KD_HPP_INT")
        
        text(2) = rsSearch("FS_KD_LABOUR_EXT")
        text(72) = rsSearch("FS_KD_SUBLET_EXT")
        text(4) = rsSearch("FS_KD_METERAI_EXT")
        text(14) = rsSearch("FS_KD_PPN_EXT")
        text(12) = rsSearch("FS_KD_PIUTANG_EXT")
        text(64) = rsSearch("FS_KD_HPP_EXT")
        
        text(58) = rsSearch("FS_KD_LABOUR_AFFNONPPN")
        text(74) = rsSearch("FS_KD_SUBLET_AFFNONPPN")
        text(56) = rsSearch("FS_KD_METERAI_AFFNONPPN")
        text(62) = rsSearch("FS_KD_PPN_AFFNONPPN")
        text(57) = rsSearch("FS_KD_PIUTANG_AFFNONPPN")
        text(92) = rsSearch("FS_KD_HPP_AFFNONPPN")
        
        text(48) = rsSearch("FS_KD_LABOUR_AFFPPN")
        text(76) = rsSearch("FS_KD_SUBLET_AFFPPN")
        text(49) = rsSearch("FS_KD_METERAI_AFFPPN")
        text(54) = rsSearch("FS_KD_PPN_AFFPPN")
        text(52) = rsSearch("FS_KD_PIUTANG_AFFPPN")
        text(94) = rsSearch("FS_KD_HPP_AFFPPN")
        
        
        ''' RETUR
        text(26) = rsSearch("FS_KD_RLABOUR_INT")
        text(78) = rsSearch("FS_KD_RSUBLET_INT")
        text(24) = rsSearch("FS_KD_RMETERAI_INT")
        text(25) = rsSearch("FS_KD_RPIUTANG_INT")
        text(30) = rsSearch("FS_KD_RPPN_INT")
        text(90) = rsSearch("FS_KD_RHPP_INT")
        
        text(16) = rsSearch("FS_KD_RLABOUR_EXT")
        text(80) = rsSearch("FS_KD_RSUBLET_EXT")
        text(17) = rsSearch("FS_KD_RMETERAI_EXT")
        text(20) = rsSearch("FS_KD_RPIUTANG_EXT")
        text(22) = rsSearch("FS_KD_RPPN_EXT")
        text(86) = rsSearch("FS_KD_RHPP_EXT")
        
        text(34) = rsSearch("FS_KD_RLABOUR_AFFNONPPN")
        text(82) = rsSearch("FS_KD_RSUBLET_AFFNONPPN")
        text(32) = rsSearch("FS_KD_RMETERAI_AFFNONPPN")
        text(38) = rsSearch("FS_KD_RPPN_AFFNONPPN")
        text(33) = rsSearch("FS_KD_RPIUTANG_AFFNONPPN")
        text(98) = rsSearch("FS_KD_RHPP_AFFNONPPN")
        
        text(40) = rsSearch("FS_KD_RLABOUR_AFFPPN")
        text(84) = rsSearch("FS_KD_RSUBLET_AFFPPN")
        text(41) = rsSearch("FS_KD_RMETERAI_AFFPPN")
        text(46) = rsSearch("FS_KD_RPPN_AFFPPN")
        text(44) = rsSearch("FS_KD_RPIUTANG_AFFPPN")
        text(96) = rsSearch("FS_KD_RHPP_AFFPPN")
        
        text(65) = rsSearch("FS_KD_PDP")
        text(68) = rsSearch("FS_KD_HUTANG")
        text(100) = rsSearch("FS_KD_PPN")
    End If
    
    
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(0)) & "'")
        If Not rsSearch.BOF Then
            text(1) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(5)) & "'")
        If Not rsSearch.BOF Then
            text(7) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(10)) & "'")
        If Not rsSearch.BOF Then
            text(11) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(6)) & "'")
        If Not rsSearch.BOF Then
            text(9) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(2)) & "'")
        If Not rsSearch.BOF Then
            text(3) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(4)) & "'")
        If Not rsSearch.BOF Then
            text(8) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(14)) & "'")
        If Not rsSearch.BOF Then
            text(15) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(12)) & "'")
        If Not rsSearch.BOF Then
            text(13) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(58)) & "'")
        If Not rsSearch.BOF Then
            text(59) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(56)) & "'")
        If Not rsSearch.BOF Then
            text(60) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(62)) & "'")
        If Not rsSearch.BOF Then
            text(63) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(57)) & "'")
        If Not rsSearch.BOF Then
            text(61) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(48)) & "'")
        If Not rsSearch.BOF Then
            text(50) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(49)) & "'")
        If Not rsSearch.BOF Then
            text(51) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(54)) & "'")
        If Not rsSearch.BOF Then
            text(55) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(52)) & "'")
        If Not rsSearch.BOF Then
            text(53) = rsSearch("namagl")
        End If
        
        ''RETUR
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(26)) & "'")
        If Not rsSearch.BOF Then
            text(27) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(24)) & "'")
        If Not rsSearch.BOF Then
            text(28) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(30)) & "'")
        If Not rsSearch.BOF Then
            text(31) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(25)) & "'")
        If Not rsSearch.BOF Then
            text(29) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(16)) & "'")
        If Not rsSearch.BOF Then
            text(18) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(17)) & "'")
        If Not rsSearch.BOF Then
            text(19) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(22)) & "'")
        If Not rsSearch.BOF Then
            text(23) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(20)) & "'")
        If Not rsSearch.BOF Then
            text(21) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(34)) & "'")
        If Not rsSearch.BOF Then
            text(35) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(32)) & "'")
        If Not rsSearch.BOF Then
            text(36) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(38)) & "'")
        If Not rsSearch.BOF Then
            text(39) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(33)) & "'")
        If Not rsSearch.BOF Then
            text(37) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(40)) & "'")
        If Not rsSearch.BOF Then
            text(42) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(41)) & "'")
        If Not rsSearch.BOF Then
            text(43) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(46)) & "'")
        If Not rsSearch.BOF Then
            text(47) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(44)) & "'")
        If Not rsSearch.BOF Then
            text(45) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(65)) & "'")
        If Not rsSearch.BOF Then
            text(67) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(68)) & "'")
        If Not rsSearch.BOF Then
            text(69) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(70)) & "'")
        If Not rsSearch.BOF Then
            text(71) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(72)) & "'")
        If Not rsSearch.BOF Then
            text(73) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(74)) & "'")
        If Not rsSearch.BOF Then
            text(75) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(76)) & "'")
        If Not rsSearch.BOF Then
            text(77) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(78)) & "'")
        If Not rsSearch.BOF Then
            text(79) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(80)) & "'")
        If Not rsSearch.BOF Then
            text(81) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(82)) & "'")
        If Not rsSearch.BOF Then
            text(83) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(84)) & "'")
        If Not rsSearch.BOF Then
            text(85) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(64)) & "'")
        If Not rsSearch.BOF Then
            text(66) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(88)) & "'")
        If Not rsSearch.BOF Then
            text(89) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(92)) & "'")
        If Not rsSearch.BOF Then
            text(93) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(94)) & "'")
        If Not rsSearch.BOF Then
            text(95) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(86)) & "'")
        If Not rsSearch.BOF Then
            text(87) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(90)) & "'")
        If Not rsSearch.BOF Then
            text(91) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(98)) & "'")
        If Not rsSearch.BOF Then
            text(99) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(96)) & "'")
        If Not rsSearch.BOF Then
            text(97) = rsSearch("namagl")
        End If
        
        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(100)) & "'")
        If Not rsSearch.BOF Then
            text(101) = rsSearch("namagl")
        End If
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0, 2, 4, 5, 6, 10, 12, 14, 16, 17, 20, 22, 24, 25, 26, 30, 58, 56, 62, 57, 48, 49, 54, 52, 34, 32, 38, 33, 40, 41, 46, 44, 64, 65, 68, 70, 72, 74, 76, 78, 80, 82, 84, 88, 92, 94, 86, 90, 98, 96
                gMainKeyCari = text(Index)
                gSearchOption = "KODEJNRL"
                frmCariAcess.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    
    Select Case Index
        Case 0, 2, 4, 5, 6, 10, 12, 14, 16, 17, 20, 22, 24, 25, 26, 30, 58, 56, 62, 57, 48, 49, 54, 52, 34, 32, 38, 33, 40, 41, 46, 44, 64, 65, 68, 70, 72, 74, 76, 78, 80, 82, 84, 88, 92, 94, 86, 90, 98, 96, 100
            If Trim(text(Index)) <> "" Then
                
                    Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(text(Index)) & "'")
                    If rsSearch.BOF Then
                        Cancel = True
                        gMainKeyCari = text(Index)
                        gSearchOption = "KODEJNRL"
                        frmCariAcess.Show (vbModal)
                    Else
                        Select Case Index
                            Case 0
                                text(1) = rsSearch("namagl")
                            Case 2
                                text(3) = rsSearch("namagl")
                            Case 4
                                text(8) = rsSearch("namagl")
                            Case 5
                                text(7) = rsSearch("namagl")
                            Case 6
                                text(9) = rsSearch("namagl")
                            Case 10
                                text(11) = rsSearch("namagl")
                            Case 12
                                text(13) = rsSearch("namagl")
                            Case 14
                                text(15) = rsSearch("namagl")
                            Case 16
                                text(18) = rsSearch("namagl")
                            Case 17
                                text(19) = rsSearch("namagl")
                            Case 20
                                text(21) = rsSearch("namagl")
                            Case 22
                                text(23) = rsSearch("namagl")
                            Case 24
                                text(28) = rsSearch("namagl")
                            Case 25
                                text(29) = rsSearch("namagl")
                            Case 26
                                text(27) = rsSearch("namagl")
                            Case 30
                                text(31) = rsSearch("namagl")
                            Case 58
                                text(59) = rsSearch("namagl")
                            Case 56
                                text(60) = rsSearch("namagl")
                            Case 62
                                text(63) = rsSearch("namagl")
                            Case 57
                                text(61) = rsSearch("namagl")
                            Case 48
                                text(50) = rsSearch("namagl")
                            Case 49
                                text(51) = rsSearch("namagl")
                            Case 54
                                text(55) = rsSearch("namagl")
                            Case 52
                                text(53) = rsSearch("namagl")
                            Case 34
                                text(35) = rsSearch("namagl")
                            Case 32
                                text(36) = rsSearch("namagl")
                            Case 38
                                text(39) = rsSearch("namagl")
                            Case 33
                                text(37) = rsSearch("namagl")
                            Case 40
                                text(42) = rsSearch("namagl")
                            Case 41
                                text(43) = rsSearch("namagl")
                            Case 46
                                text(47) = rsSearch("namagl")
                            Case 44
                                text(45) = rsSearch("namagl")
                            Case 58
                                text(59) = rsSearch("namagl")
                            Case 56
                                text(60) = rsSearch("namagl")
                            Case 62
                                text(63) = rsSearch("namagl")
                            Case 57
                                text(61) = rsSearch("namagl")
                            Case 48
                                text(50) = rsSearch("namagl")
                            Case 49
                                text(51) = rsSearch("namagl")
                            Case 54
                                text(55) = rsSearch("namagl")
                            Case 52
                                text(53) = rsSearch("namagl")
                            Case 34
                                text(35) = rsSearch("namagl")
                            Case 32
                                text(36) = rsSearch("namagl")
                            Case 38
                                text(39) = rsSearch("namagl")
                            Case 33
                                text(37) = rsSearch("namagl")
                            Case 40
                                text(42) = rsSearch("namagl")
                            Case 41
                                text(43) = rsSearch("namagl")
                            Case 46
                                text(47) = rsSearch("namagl")
                            Case 44
                                text(45) = rsSearch("namagl")
                            
                            Case 64
                                text(66) = rsSearch("namagl")
                            Case 65
                                text(67) = rsSearch("namagl")
                            Case 68
                                text(69) = rsSearch("namagl")
                            
                            Case 70
                                text(Index + 1) = rsSearch("namagl")
                            Case 72
                                text(Index + 1) = rsSearch("namagl")
                            Case 74
                                text(Index + 1) = rsSearch("namagl")
                            Case 76
                                text(Index + 1) = rsSearch("namagl")
                            Case 78
                                text(Index + 1) = rsSearch("namagl")
                            Case 80
                                text(Index + 1) = rsSearch("namagl")
                            Case 82
                                text(Index + 1) = rsSearch("namagl")
                            Case 84
                                text(Index + 1) = rsSearch("namagl")
                            
                            Case 88
                                text(Index + 1) = rsSearch("namagl")
                            Case 92
                                text(Index + 1) = rsSearch("namagl")
                            Case 94
                                text(Index + 1) = rsSearch("namagl")
                                
                            Case 86
                                text(Index + 1) = rsSearch("namagl")
                            Case 90
                                text(Index + 1) = rsSearch("namagl")
                            Case 98
                                text(Index + 1) = rsSearch("namagl")
                            Case 96
                                text(Index + 1) = rsSearch("namagl")
                            Case 100
                                text(Index + 1) = rsSearch("namagl")
                        End Select
                    End If
                
            Else
                Select Case Index
                    Case 0
                        text(1) = ""
                    Case 2
                        text(3) = ""
                    Case 4
                        text(8) = ""
                    Case 5
                        text(7) = ""
                    Case 6
                        text(9) = ""
                    Case 10
                        text(11) = ""
                    Case 12
                        text(13) = ""
                    Case 14
                        text(15) = ""
                    Case 16
                        text(18) = ""
                    Case 17
                        text(19) = ""
                    Case 20
                        text(21) = ""
                    Case 22
                        text(23) = ""
                    Case 24
                        text(28) = ""
                    Case 25
                        text(29) = ""
                    Case 26
                        text(27) = ""
                    Case 30
                        text(31) = ""
                    Case 58
                        text(59) = ""
                    Case 56
                        text(60) = ""
                    Case 62
                        text(63) = ""
                    Case 57
                        text(61) = ""
                    Case 48
                        text(50) = ""
                    Case 49
                        text(51) = ""
                    Case 54
                        text(55) = ""
                    Case 52
                        text(53) = ""
                    Case 34
                        text(35) = ""
                    Case 32
                        text(36) = ""
                    Case 38
                        text(39) = ""
                    Case 33
                        text(37) = ""
                    Case 40
                        text(42) = ""
                    Case 41
                        text(43) = ""
                    Case 46
                        text(47) = ""
                    Case 44
                        text(45) = ""
                        
                    Case 64
                        text(66) = ""
                    Case 65
                        text(67) = ""
                    Case 68
                        text(69) = ""
                    Case 70
                        text(Index + 1) = ""
                    Case 72
                        text(Index + 1) = ""
                    Case 74
                        text(Index + 1) = ""
                    Case 76
                        text(Index + 1) = ""
                    Case 78
                        text(Index + 1) = ""
                    Case 80
                        text(Index + 1) = ""
                    Case 82
                        text(Index + 1) = ""
                    Case 84
                        text(Index + 1) = ""
                    
                    Case 88
                        text(Index + 1) = ""
                    Case 92
                        text(Index + 1) = ""
                    Case 94
                        text(Index + 1) = ""
                        
                    Case 86
                        text(Index + 1) = ""
                    Case 90
                        text(Index + 1) = ""
                    Case 98
                        text(Index + 1) = ""
                    Case 96
                        text(Index + 1) = ""
                    Case 100
                        text(Index + 1) = ""
                End Select
            End If
    End Select
    
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 0 To 101
        text(i) = ""
    Next
End Sub
