VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmtAddPlgn2 
   Caption         =   "Tambah Pelanggan"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12075
   ControlBox      =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmePelanggan 
      BackColor       =   &H00E0E0E0&
      Height          =   3180
      Left            =   5280
      TabIndex        =   23
      Top             =   3960
      Width           =   5820
      Begin VB.ComboBox CboStatus 
         Height          =   315
         ItemData        =   "frmtAddPlgn.frx":0000
         Left            =   1275
         List            =   "frmtAddPlgn.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2625
         Width           =   1935
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   6
         Left            =   1275
         TabIndex        =   7
         Top             =   465
         Width           =   3270
         _ExtentX        =   5768
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
         Index           =   7
         Left            =   1275
         TabIndex        =   8
         Top             =   765
         Width           =   4305
         _ExtentX        =   7594
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
         Height          =   270
         Index           =   8
         Left            =   1275
         TabIndex        =   9
         Top             =   1065
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
         Height          =   285
         Index           =   10
         Left            =   1275
         TabIndex        =   11
         Top             =   1665
         Width           =   1695
         _ExtentX        =   2990
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
         Index           =   11
         Left            =   1275
         TabIndex        =   12
         Top             =   1980
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Index           =   12
         Left            =   1275
         TabIndex        =   13
         Top             =   2295
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   9
         Left            =   1275
         TabIndex        =   10
         Top             =   1365
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Index           =   0
         Left            =   2040
         TabIndex        =   24
         Top             =   2310
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
         Index           =   5
         Left            =   4080
         TabIndex        =   6
         Top             =   180
         Width           =   465
         _ExtentX        =   820
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
         Left            =   1275
         TabIndex        =   5
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
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
         Height          =   285
         Index           =   13
         Left            =   2130
         TabIndex        =   14
         Top             =   2295
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   735
         TabIndex        =   34
         Top             =   540
         Width           =   405
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   645
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   810
         TabIndex        =   32
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Perush."
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
         Left            =   90
         TabIndex        =   31
         Top             =   2655
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Syarat Bayar"
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
         Left            =   195
         TabIndex        =   30
         Top             =   2340
         Width           =   945
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
         Left            =   705
         TabIndex        =   29
         Top             =   2025
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Left            =   885
         TabIndex        =   28
         Top             =   1710
         Width           =   270
      End
      Begin VB.Label Label4 
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
         Left            =   840
         TabIndex        =   27
         Top             =   1410
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
         Left            =   480
         TabIndex        =   26
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Wilayah"
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
         Left            =   3015
         TabIndex        =   25
         Top             =   225
         Width           =   975
      End
   End
   Begin VB.Frame frmeNama 
      Height          =   855
      Left            =   4800
      TabIndex        =   21
      Top             =   1320
      Width           =   6015
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   270
         Width           =   4380
         _ExtentX        =   7726
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
         Height          =   270
         Index           =   3
         Left            =   1440
         TabIndex        =   3
         Top             =   555
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Customer"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   315
         Width           =   1125
      End
   End
   Begin VB.Frame frmeWilayah 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6855
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   277
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   6720
         TabIndex        =   20
         Top             =   277
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
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   550
         Visible         =   0   'False
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Wilayah"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   315
         Width           =   1035
      End
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Batal"
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
      Height          =   360
      Left            =   2520
      TabIndex        =   18
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Simpan"
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
   Begin VB.Label lblKet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tekan Enter untuk melanjutkan"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   3360
   End
End
Attribute VB_Name = "frmtAddPlgn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flagUnload As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim vStatus As String
Dim rsSimpan As New ADODB.Recordset

    If text(5) <> "" Then
        If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
            If UCase(CboStatus.text) = "SWASTA" Then
                vStatus = "S"
            Else
                vStatus = "P"
            End If
            rsSimpan.Open "select KodeLgn from ARMST where KodeLgn = '" & Trim(text(4)) & "'", _
                     oConnNis, adOpenStatic
        
            If rsSimpan.BOF Then
                sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2" _
                     & " ,Telepon,Nofax,Npwp,SyaratBayar,StatusPersh,KodeUser)" _
                    & " VALUES ('" & text(4) & "','" & text(6) & "','" & text(5) & "'" _
                    & ",'" & text(7) & "','" & text(8) & "','" & Replace(Replace(text(9), ">", ""), " ", "") & "','" & Replace(Replace(text(10), ">", ""), " ", "") & "'" _
                    & ",'" & text(11) & "','" & text(12) & "','" & vStatus & "','" & gSA & "')"
            Else
                sSQL = "UPDATE ARMST SET KodeLgn = '" & text(4) & "', NamaLgn = '" & text(6) & "'," _
                     & " KodeWil = '" & text(5) & "',Alamat1 = '" & text(7) & "',Alamat2 = '" & text(8) & "'," _
                    & " Telepon = '" & Replace(Replace(text(9), ">", ""), " ", "") & "',Nofax = '" & Replace(Replace(text(10), ">", ""), " ", "") & "',Npwp = '" & text(11) & "'," _
                    & " SyaratBayar = '" & text(12) & "',StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "'" _
                    & " where KodeLgn = '" & Trim(text(4)) & "'"
            End If
            oConnNis.Execute sSQL
            rsSimpan.Close
            
            gBantuan = Trim(text(4))
            'Call hapusForm
            'Text(57) = ""
            'text(57).SetFocus
            
            With frmtCustomer
                If .Visible Then
                    .text(2) = Trim(gBantuan)
                    '.text(2).SetFocus
                    Unload Me
                End If
            End With
        End If
        If oConnNis.State <> adStateClosed Then oConnNis.Close
    End If
End Sub

Private Sub Form_Load()
    'If Me.Tag = "wilayah" Then
        With frmeWilayah
            .Visible = True
            .Top = 0
            .Left = 0
        End With
        lblKet.Top = frmeWilayah.Top + frmeWilayah.Height + 50
        lblKet.Left = frmeWilayah.Left
        MsgBox "Input wilayah customer berdasarkan KTP ."
        cmdSimpan.Visible = False
        cmdExit.Visible = False
        frmeNama.Visible = False
        frmePelanggan.Visible = False
        flagUnload = False
    'End If
End Sub

Private Sub text_GotFocus(Index As Integer)
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

Private Sub text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = text(Index)
                gSearchOption = "KODEWIL"
                frmCariAcess.Show (vbModal)
            Case 12
                gMainKeyCari = text(Index)
                gSearchOption = "SYARAT"
                frmCariAcess.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Dim rsSearch As New ADODB.Recordset
Dim clear As Boolean
    
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
    
    If KeyAscii = 13 Then
        If rsSearch.State Then rsSearch.Close
        If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
            Select Case Index
                Case 0
                    If text(Index) <> "" Then
                        sSQL = "SELECT * FROM ARWILAYAH WHERE NAMAWIL LIKE '%" & Trim(text(Index)) & "%'"
                        rsSearch.Open sSQL, oConnNis, adOpenStatic, adLockReadOnly
                        If Not rsSearch.EOF Then
                            clear = False
                            Do While Not rsSearch.EOF
                                If MsgBox("Wilayah : " & rsSearch("NAMAWIL") & " ?", vbYesNo, "WILAYAH") = vbYes Then
                                    ' LANJUT STEP BERIKUT
                                    text(1) = Trim(rsSearch("KODEWIL"))
                                    text(0) = Trim(rsSearch("NAMAWIL"))
                                    With frmeNama
                                        .Visible = True
                                        .Left = 0
                                        .Top = 0
                                    End With
                                    lblKet.Top = frmeNama.Top + frmeNama.Height
                                    lblKet.Left = frmeNama.Left
                                    text(2).SetFocus
                                    clear = True
                                    frmeWilayah.Visible = False
                                    Exit Sub
                                End If
                                rsSearch.MoveNext
                            Loop
                            rsSearch.Close
                            If Not clear Then
                                gMainKeyCari = text(Index)
                                gSearchOption = "KODEWIL"
                                frmCariAcess.Show (vbModal)
                            End If
                        Else
                            If MsgBox("Tambah Wilayah baru ?", vbYesNo, "WILAYAH") = vbYes Then
                                frmtWil.Show vbModal
                            Else
                                gMainKeyCari = text(Index)
                                gSearchOption = "KODEWIL"
                                frmCariAcess.Show (vbModal)
                            End If
                        End If
                    End If
                Case 2
                    If text(Index) <> "" Then
                        flagUnload = False
                        sSQL = "SELECT * FROM ARMST WHERE NAMALGN LIKE '%" & Trim(text(Index)) & "%' AND KODEWIL = '" & Trim(text(1)) & "'"
                        rsSearch.Open sSQL, oConnNis, adOpenStatic, adLockReadOnly
                        If Not rsSearch.EOF Then
                            ' CUSTOMER ADA DI DBASE
                            Do While Not rsSearch.EOF
                                If MsgBox("Customer yang dimaksud : " & Trim(rsSearch("NAMALGN")) & " ?", vbYesNo, "CUSTOMER") = vbYes Then
                                    text(3) = Trim(rsSearch("KODELGN"))
                                    text(2) = Trim(rsSearch("NAMALGN"))
                                    MsgBox "Kode Pelanggan : " & Trim(rsSearch("KODELGN"))
                                    With frmtCustomer
                                        If .Visible Then
                                            .text(2) = Trim(rsSearch("KODELGN"))
                                            .text(3) = Trim(rsSearch("KODEWIL"))
                                            .text(4) = Trim(rsSearch("NAMALGN"))
                                            .text(5) = Trim(rsSearch("ALAMAT1"))
                                            .text(6) = Trim(rsSearch("ALAMAT2"))
                                        End If
                                    End With
                                    flagUnload = True
                                    Exit Do
                                End If
                                rsSearch.MoveNext
                            Loop
                            rsSearch.Close
                            If flagUnload Then
                                Unload Me
                            Else
                                ' BETUL2 CUSTOMER BARU
                                With frmePelanggan
                                    .Visible = True
                                    .Left = 0
                                    .Top = 0
                                End With
                                text(4) = Trim(text(1)) & Trim(Left(text(2), 3)) & "-" & cariCounter(Trim(text(1)) & Trim(Left(text(2), 3)))
                                text(5) = Trim(text(1))
                                text(6) = Trim(text(2))
                                text(8) = Trim(text(0))
                                cmdSimpan.Visible = True
                                cmdExit.Visible = True
                                frmeNama.Visible = False
                                CboStatus.AddItem "Pemerintah", 0
                                CboStatus.AddItem "Swasta", 1
                            End If
                        Else
                            ' BETUL2 CUSTOMER BARU
                            With frmePelanggan
                                .Visible = True
                                .Left = 0
                                .Top = 0
                            End With
                            text(4) = Trim(text(1)) & Trim(Left(text(2), 3)) & "-" & cariCounter(Trim(text(1)) & Trim(Left(text(2), 3)))
                            text(5) = Trim(text(1))
                            text(6) = Trim(text(2))
                            text(8) = Trim(text(0))
                            cmdSimpan.Visible = True
                            cmdExit.Visible = True
                            frmeNama.Visible = False
                            CboStatus.AddItem "Pemerintah", 0
                            CboStatus.AddItem "Swasta", 1
                        End If
                    End If
                Case 12
                    If text(Index) <> "" Then
                        rsSearch.Open "select kodesyarat,KetSyaratBayar from arsyarat where " & _
                                 "kodesyarat = '" & text(Index) & "'", oConnNis, adOpenStatic
                        If rsSearch.BOF Then
                            Cancel = True
                            gMainKeyCari = text(Index)
                            gSearchOption = "SYARAT"
                            frmCariAcess.Show (vbModal)
                        Else
                            text(0) = rsSearch("KetSyaratBayar")
                        End If
                        rsSearch.Close
                    End If
            End Select
        End If
        If oConnNis.State <> adStateClosed Then oConnNis.Close
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
'Dim rsSearch As New ADODB.Recordset
'Dim clear As Boolean
'
'    If rsSearch.State Then rsSearch.Close
'    If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
'        Select Case Index
'            Case 0
'                If Text(Index) <> "" Then
'                    sSQL = "SELECT * FROM ARWILAYAH WHERE NAMAWIL LIKE '%" & Trim(Text(Index)) & "%'"
'                    rsSearch.Open sSQL, oConnNis, adOpenStatic, adLockReadOnly
'                    If Not rsSearch.EOF Then
'                        clear = False
'                        Do While Not rsSearch.EOF
'                            If MsgBox("Wilayah : " & rsSearch("NAMAWIL") & " ?", vbYesNo, "WILAYAH") = vbYes Then
'                                ' LANJUT STEP BERIKUT
'                                Text(1) = Trim(rsSearch("KODEWIL"))
'                                Text(0) = Trim(rsSearch("NAMAWIL"))
'                                With frmeNama
'                                    .Visible = True
'                                    .Left = 0
'                                    .Top = 0
'                                End With
'                                lblKet.Top = frmeNama.Top + frmeNama.Height
'                                lblKet.Left = frmeNama.Left
'                                Text(2).SetFocus
'                                clear = True
'                                frmeWilayah.Visible = False
'                                Exit Sub
'                            End If
'                            rsSearch.MoveNext
'                        Loop
'                        If Not clear Then
'                            gMainKeyCari = Text(Index)
'                            gSearchOption = "KODEWIL"
'                            frmCariAcess.Show (vbModal)
'                        End If
'                    Else
'                        If MsgBox("Tambah Wilayah baru ?", vbYesNo, "WILAYAH") = vbYes Then
'                            frmtWil.Show vbModal
'                        Else
'                            gMainKeyCari = Text(Index)
'                            gSearchOption = "KODEWIL"
'                            frmCariAcess.Show (vbModal)
'                        End If
'                    End If
'                    rsSearch.Close
'
'                End If
'            Case 2
'                If Text(Index) <> "" Then
'                    flagUnload = False
'                    sSQL = "SELECT * FROM ARMST WHERE NAMALGN LIKE '%" & Trim(Text(Index)) & "%' AND KODEWIL = '" & Trim(Text(1)) & "'"
'                    rsSearch.Open sSQL, oConnNis, adOpenStatic, adLockReadOnly
'                    If Not rsSearch.EOF Then
'                        ' CUSTOMER ADA DI DBASE
'                        Do While Not rsSearch.EOF
'                            If MsgBox("Customer yang dimaksud : " & Trim(rsSearch("NAMALGN")) & " ?", vbYesNo, "CUSTOMER") = vbYes Then
'                                Text(3) = Trim(rsSearch("KODELGN"))
'                                Text(2) = Trim(rsSearch("NAMALGN"))
'                                MsgBox "Kode Pelanggan : " & Trim(rsSearch("KODELGN"))
'                                With frmtCustomer
'                                    If .Visible Then
'                                        .Text(2) = Trim(rsSearch("KODELGN"))
'                                        .Text(3) = Trim(rsSearch("KODEWIL"))
'                                        .Text(4) = Trim(rsSearch("NAMALGN"))
'                                        .Text(5) = Trim(rsSearch("ALAMAT1"))
'                                        .Text(6) = Trim(rsSearch("ALAMAT2"))
'                                    End If
'                                End With
'                                flagUnload = True
'                                Exit Do
'                            End If
'                            rsSearch.MoveNext
'                        Loop
'                        rsSearch.Close
'                    Else
'                        ' BETUL2 CUSTOMER BARU
'                        With frmePelanggan
'                            .Visible = True
'                            .Left = 0
'                            .Top = 0
'                        End With
'                        Text(4) = Trim(Text(1)) & Trim(Left(Text(2), 3)) & "-" & cariCounter(Trim(Text(1)) & Trim(Left(Text(2), 3)))
'                        Text(5) = Trim(Text(1))
'                        Text(6) = Trim(Text(2))
'                        Text(8) = Trim(Text(0))
'                        cmdSimpan.Visible = True
'                        cmdExit.Visible = True
'                        frmeNama.Visible = False
'                        CboStatus.AddItem "Pemerintah", 0
'                        CboStatus.AddItem "Swasta", 1
'                    End If
'                End If
'            Case 12
'                If Text(Index) <> "" Then
'                    rsSearch.Open "select kodesyarat,KetSyaratBayar from arsyarat where " & _
'                             "kodesyarat = '" & Text(Index) & "'", oConnNis, adOpenStatic
'                    If rsSearch.BOF Then
'                        Cancel = True
'                        gMainKeyCari = Text(Index)
'                        gSearchOption = "SYARAT"
'                        frmCariAcess.Show (vbModal)
'                    Else
'                        Text(0) = rsSearch("KetSyaratBayar")
'                    End If
'                    rsSearch.Close
'                End If
'        End Select
'    End If
'    If oConnNis.State Then oConnNis.Close
End Sub

Private Function cariCounter(cust As String) As String
Dim rsCari As New ADODB.Recordset

    If Get_ConnectionNis(NisDb, NisUSER, NisPass, NisKomp) Then
        If rsCari.State Then rsCari.Close
        Set rsCari = oConnNis.Execute("SELECT SUBSTRING(KODELGN, 8,4) AS COUNTER, * FROM ARMST WHERE LEFT(KODELGN,6) = '" & Trim(cust) & "' ORDER BY CAST(SUBSTRING(KODELGN, 8,4) AS INTEGER) DESC")
        If Not rsCari.EOF Then
            cariCounter = CDbl(Trim(rsCari("COUNTER"))) + 1
        Else
            cariCounter = 1
        End If
        rsCari.Close
        cariCounter = Format(cariCounter, "0000")
    End If
    If oConnNis.State <> adStateClosed Then oConnNis.Close
End Function
