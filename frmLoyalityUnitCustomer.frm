VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmLoyalityUnitCustomer 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15690
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   15690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5640
      Left            =   495
      TabIndex        =   1
      Top             =   615
      Width           =   10320
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   735
         Left            =   1005
         Picture         =   "frmLoyalityUnitCustomer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4485
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   1755
         Picture         =   "frmLoyalityUnitCustomer.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4500
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
         Picture         =   "frmLoyalityUnitCustomer.frx":0D2B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   735
      End
      Begin MSMask.MaskEdBox Text 
         DataField       =   "c"
         Height          =   390
         Index           =   0
         Left            =   2040
         TabIndex        =   3
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
         Index           =   1
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
         Index           =   2
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
         Index           =   3
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
         Index           =   4
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
         Index           =   6
         Left            =   1020
         TabIndex        =   11
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
         Index           =   5
         Left            =   4185
         TabIndex        =   10
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
         TabIndex        =   22
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
         Index           =   7
         Left            =   1020
         TabIndex        =   12
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
         TabIndex        =   24
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
         Index           =   8
         Left            =   1020
         TabIndex        =   13
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
         Index           =   10
         Left            =   1020
         TabIndex        =   15
         Top             =   3000
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
         Index           =   11
         Left            =   1020
         TabIndex        =   16
         Top             =   3300
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
         Left            =   2730
         TabIndex        =   28
         Top             =   2700
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
         Index           =   9
         Left            =   1020
         TabIndex        =   14
         Top             =   2700
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
         Index           =   12
         Left            =   1020
         TabIndex        =   17
         Top             =   3615
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
         Index           =   13
         Left            =   1020
         TabIndex        =   18
         Top             =   3900
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
         Index           =   14
         Left            =   1020
         TabIndex        =   19
         Top             =   4200
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   405
         Left            =   5190
         TabIndex        =   34
         Top             =   345
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   75
         X2              =   10230
         Y1              =   2640
         Y2              =   2640
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
         Left            =   555
         TabIndex        =   33
         Top             =   3975
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
         Left            =   585
         TabIndex        =   32
         Top             =   4245
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
         Left            =   240
         TabIndex        =   31
         Top             =   2745
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
         Left            =   405
         TabIndex        =   30
         Top             =   3375
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
         Left            =   495
         TabIndex        =   29
         Top             =   3075
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   23
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
         X2              =   10260
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
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   ".:: UNIT CUSTOMER ::."
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
      Width           =   10500
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   10605
   End
   Begin VB.Image Image1 
      Height          =   8190
      Left            =   15
      Picture         =   "frmLoyalityUnitCustomer.frx":137A
      Top             =   45
      Width           =   450
   End
End
Attribute VB_Name = "frmLoyalityUnitCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTemp As ADODB.Recordset
Private rsHlp As ADODB.Recordset
Public emailCustomer As String

Private Sub clearText()
    For i = 0 To 14
        Text(i) = ""
    Next
    cmdSave.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
   On Error GoTo cmdSave_Click_Error

    If Trim(Text(0)) = "" Then
        MsgBox "No Rangka harus diisi", vbInformation
        Exit Sub
    End If
    
    Set rsHlp = objO.FetchData("SELECT b.tglEntry,b.userEntry FROM RPLNASAPP..MASTERUSER a INNER JOIN RPLNASAPP..NewCustomer b ON a.alamatemail = b.email WHERE a.VERIFIKASI = 1 AND a.ALAMATEMAIL = '" & Trim(emailCustomer) & "'")
    If Not rsHlp.EOF Then
        MsgBox "Telah diverifikasi oleh user " & Nul(rsHlp!userentry) & " pada tanggal " & TglDDMMMMYYYY(Nul(rsHlp!tglentry), 1), vbInformation
        Exit Sub
    End If
    Dim namaCust() As String
    Dim isBegin As Boolean
    
    isBegin = False
    
    If Not rsTemp.EOF Then
        Dim nama1 As String
        Dim nama2 As String
        Dim nama3 As String
        
        namaCust() = Split(Trim(rsTemp!CUSTOMER), " ")
        For i = 0 To UBound(namaCust)
            If i = 0 Then
                nama1 = UCase(namaCust(i))
            ElseIf i = 1 Then
                nama2 = UCase(namaCust(i))
            ElseIf i >= 2 Then
                nama3 = nama3 & " " & UCase(namaCust(i))
            End If
        Next
        
        Call objHT.OpenTransactionAll
        isBegin = True
        
        sSQL = "INSERT INTO [RplNasApp]..NewCustomer(topVhc,BatasVhc,topGrp,BatasGrp,topBp,BatasBp,topPart,BatasPart,nonPPn,fpgab,typeCustomer,KodeDealer,IDCust,noktp , nonpwp ,title , namacust1 , namacust2 ,namacust3 , tgllahir ,agama ,pekerjaan , propinsi ,kota ,kecamatan , kelurahan ,kodepos , rt ,rw ,idtelp ,telp,hp1,hobi, alamat,email,userentry,tglentry)"
        sSQL = sSQL & " VALUES (0,0,0,0,0,0,0,0,0,0,'Perorangan','" & gKdCbg & "','" & Nul(rsTemp!FS_KD_PLG) & "','" & Nul(rsTemp!noktp) & "','" & Nul(rsTemp!npwp) & "','','" & Nul(nama1) & "','" & Nul(nama2) & "','" & Nul(nama3) & "',"
        sSQL = sSQL & " '" & Nul(rsTemp!tgllahir) & "','" & Nul(rsTemp!Agama) & "','" & Nul(rsTemp!Pekerjaan) & "','" & Nul(rsTemp!Propinsi) & "','" & Nul(rsTemp!Kota) & "','" & Nul(rsTemp!kecamatan) & "','" & Nul(rsTemp!kelurahan) & "',"
        sSQL = sSQL & " '" & Nul(rsTemp!KodePos) & "','" & Nul(rsTemp!rt) & "','" & Nul(rsTemp!rw) & "','','" & Nul(rsTemp!GSM) & "','" & Nul(rsTemp!GSM) & "','" & Nul(rsTemp!hobi) & "','" & Nul(rsTemp!cust_adres) & "','" & emailCustomer & "','" & gSA & "',getdate())"
        Call objHT.InsertData_Bengkel(sSQL)
        
        sSQL = "INSERT INTO [RplNasApp]..UnitCustomer(ddate,fmdl,noEngine,IDCustService,NoRangka,KdDlrPenjual,nmDlrPenjual,noPolisi,mdl,thmobil,userentry,tglentry)"
        sSQL = sSQL & " VALUES ('" & Nul(rsTemp!dDate) & "','" & Nul(rsTemp!Fmdl) & "','" & Nul(rsTemp!No_Engine) & "','" & Nul(rsTemp!FS_KD_PLG) & "','" & Nul(rsTemp!No_Chassis) & "','" & Nul(rsTemp!FS_KD_DEALER) & "','" & Nul(rsTemp!FS_NM_DEALER) & "','" & Nul(rsTemp!NO_POLISI) & "','" & Nul(rsTemp!MDL) & "','" & Nul(rsTemp!FS_TH_MOBIL) & "','" & gSA & "',getdate())"
        Call objHT.InsertData_Bengkel(sSQL)
         
        
        Call objHT.InsertData_Bengkel("UPDATE CUSTDAT SET useemail = '" & emailCustomer & "' WHERE NO_CHASSIS = '" & Nul(rsTemp!No_Chassis) & "'")
        
        Call objHT.InsertData_Bengkel("UPDATE [RplNasApp]..MASTERUSER SET verifikasi = '1' WHERE AlamatEmail = '" & emailCustomer & "' ")
        
        Call objHT.CloseTransactionAll
        MsgBox "Sukses menambahkan unit", vbInformation
        Unload Me
    Else
        MsgBox "Tidak ada data detail untuk norangka " & Trim(Text(0)), vbInformation
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:
    If isBegin = True Then
        Call objHT.CloseTransactionAll_Err
    End If
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSave_Click of Form frmLoyalityUnitCustomer"
End Sub

Private Sub cmdSearch2_Click()
    Call panggilDataNoRangka
End Sub

Private Sub Form_Load()
    Call clearText
    
End Sub

Private Sub panggilDataNoRangka()

    If Trim(Text(0)) = "" Then
        MsgBox "No Rangka harus diisi.", vbInformation
        Exit Sub
    End If
    sSQL = "SELECT a.FS_KD_PLG,'' as NOKTP, a.NPWP ,a.CUSTOMER,'01/01/1900' as TglLahir,a.USEAGAMA as agama,"
    sSQL = sSQL & " a.USEJOBDSC as pekerjaan,'' as propinsi,a.KOTA,a.KECAMATAN,'' as kelurahan,a.KODEPOS,"
    sSQL = sSQL & " '' as rt,'' as rw,a.GSM ,a.USEFS_NM_HOBI as hobi,a.CUST_ADRES,a.USEEMAIL,"
    sSQL = sSQL & " NO_CHASSIS,FS_KD_DEALER,FS_NM_DEALER,ddate,NO_POLISI,MDL,MODEL,FMDL,no_Engine,FS_TH_MOBIL,COLOR,'01/01/1900' as tglawalpolis,'01/01/1900' as tglakhirpolis,'' as namapolis"
    sSQL = sSQL & " FROM custdat a WITH (NOLOCK)"
    sSQL = sSQL & " WHERE a.no_chassis = '" & Trim(Text(0)) & "'"
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
            Text(1) = Trim(nDpn1)
            Text(2) = Trim(nTgh1)
            Text(3) = Trim(nBlk1)
        Else
            Text(1) = ""
            Text(2) = ""
            Text(3) = ""
        End If
        
        Text(8) = Nul(rsTemp("FS_TH_MOBIL"))
        Text(4) = Nul(rsTemp("MDL"))
        Text(5) = Nul(rsTemp("MODEL"))
        Text(6) = Nul(rsTemp("FMDL"))
        Text(7) = Nul(rsTemp("COLOR"))
        
        Text(9) = Nul(rsTemp("FS_KD_PLG"))
        Text(10) = Left(Nul(rsTemp("CUSTOMER")), 60)
        Text(11) = Left(Nul(rsTemp("CUST_ADRES")), 60)
        If Len(Trim(Nul(rsTemp("CUST_ADRES")))) > 60 Then
            Text(12) = Mid(Nul(rsTemp("CUST_ADRES")), 61, Len(Nul(rsTemp("CUST_ADRES"))) - 60)
        End If
        Text(13) = Nul(rsTemp("KOTA"))
        cmdSave.Enabled = True
    Else
        Dim noRangka As String
        noRangka = Trim(Text(0))
        Call clearText
        
        Text(0) = noRangka
        
        With frmtCustomer
            .Text(50) = Left(Text(0), 3)
            If Len(Text(0)) > 4 Then
                .Text(1) = Mid(Text(0), 4, 6)
            Else
                .Text(1) = Mid(Text(0), 4, Len(Text(0)) - 4)
            End If
            If Len(Text(Index)) > 9 Then
                .Text(0) = Right(Text(0), Len(Text(0)) - 9)
            Else
                .Text(0) = ""
            End If
            .Show (vbModal)
            Text(1).SelLength = Len(Text(1))
            Text(1).SetFocus
        End With
    End If
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
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
                gMainKeyCari = Text(Index)
                          'gSearchOption = "ORDER"
                gSearchOption = "CUSTOMER2"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If Text(Index).BackColor = &HC0E0FF Then KeyAscii = 0
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        Call panggilDataNoRangka
        
    End If
End Sub
