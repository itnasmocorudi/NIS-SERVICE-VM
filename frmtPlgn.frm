VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtPlgn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "`"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8640
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmtPlgn.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   25
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   480
      TabIndex        =   10
      Top             =   -360
      Width           =   7140
      Begin VB.ComboBox CboStatus 
         Height          =   315
         ItemData        =   "frmtPlgn.frx":2259
         Left            =   1755
         List            =   "frmtPlgn.frx":225B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3225
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1755
         TabIndex        =   1
         Top             =   1065
         Width           =   4425
         _ExtentX        =   7805
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1755
         TabIndex        =   2
         Top             =   1365
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
         Height          =   270
         Index           =   3
         Left            =   1755
         TabIndex        =   3
         Top             =   1665
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
         Height          =   285
         Index           =   5
         Left            =   1755
         TabIndex        =   5
         Top             =   2265
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   1755
         TabIndex        =   6
         Top             =   2580
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1755
         TabIndex        =   7
         Top             =   2895
         Visible         =   0   'False
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
         Index           =   4
         Left            =   1755
         TabIndex        =   4
         Top             =   1965
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
         Left            =   2520
         TabIndex        =   19
         Top             =   2910
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
         Index           =   56
         Left            =   4560
         TabIndex        =   9
         Top             =   750
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
         Index           =   57
         Left            =   1755
         TabIndex        =   0
         Top             =   743
         Width           =   1440
         _ExtentX        =   2540
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   3225
         TabIndex        =   20
         Top             =   750
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
         Index           =   0
         Left            =   2680
         TabIndex        =   23
         Top             =   2895
         Visible         =   0   'False
         Width           =   3500
         _ExtentX        =   6165
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   5040
         TabIndex        =   24
         Top             =   1665
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   4510
         TabIndex        =   26
         Top             =   3960
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
         MICON           =   "frmtPlgn.frx":225D
         PICN            =   "frmtPlgn.frx":2279
         PICH            =   "frmtPlgn.frx":278C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   5880
         TabIndex        =   27
         Top             =   3960
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
         MICON           =   "frmtPlgn.frx":2D2E
         PICN            =   "frmtPlgn.frx":2D4A
         PICH            =   "frmtPlgn.frx":321E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdtWilayah 
         Height          =   495
         Left            =   2640
         TabIndex        =   28
         Top             =   3960
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&K"
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
         MICON           =   "frmtPlgn.frx":3769
         PICN            =   "frmtPlgn.frx":3785
         PICH            =   "frmtPlgn.frx":3E60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   3375
         Left            =   120
         Top             =   480
         Width           =   6975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   7080
         Y1              =   4560
         Y2              =   4560
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
         Left            =   3495
         TabIndex        =   22
         Top             =   810
         Width           =   975
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
         Left            =   960
         TabIndex        =   21
         Top             =   825
         Width           =   660
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
         Left            =   1320
         TabIndex        =   18
         Top             =   2010
         Width           =   300
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
         Left            =   1365
         TabIndex        =   17
         Top             =   2310
         Width           =   270
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
         Left            =   1185
         TabIndex        =   16
         Top             =   2625
         Width           =   435
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
         Left            =   675
         TabIndex        =   15
         Top             =   2940
         Visible         =   0   'False
         Width           =   945
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
         Left            =   570
         TabIndex        =   14
         Top             =   3255
         Visible         =   0   'False
         Width           =   1065
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
         Left            =   1290
         TabIndex        =   13
         Top             =   1740
         Width           =   330
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
         Left            =   1125
         TabIndex        =   12
         Top             =   1440
         Width           =   495
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
         Left            =   1215
         TabIndex        =   11
         Top             =   1140
         Width           =   405
      End
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   5415
      Left            =   0
      Picture         =   "frmtPlgn.frx":4696
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmtPlgn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gInNama As String
Dim rsTemp As ADODB.Recordset

Private Function cariCounter(str As String) As String
1        On Error GoTo cariCounter_Error

2         Set rsTemp = objO.FetchDataACC("SELECT * FROM armst WITH(NOLOCK) where left(KodeLgn,6) = '" & str & "' ORDER BY kodelgn DESC")
3         If Not rsTemp.EOF Then
4             cariCounter = Format(Right(rsTemp("KodeLgn"), Len(rsTemp("kodelgn")) - 6), "00000")
5         Else
6             cariCounter = Format(1, "00000")
7         End If

8        On Error GoTo 0
9        Exit Function

cariCounter_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariCounter of Form frmtPlgn")
End Function

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombolCombo(KeyCode)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          Dim vStatus As String
          Dim rsSearch As ADODB.Recordset
          
1        On Error GoTo cmdSimpan_Click_Error

2         If Text(57) <> "" Then
              'If Get_ConnectionAcces(NisDb, NisUSER, NisPass, NisKomp) Then
                  
                  '20/12/2012  Cust Non Axtif
                  Dim rsNon As ADODB.Recordset
3                 Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & Text(57) & "' and nonaktifflag  = 1")
4                 If rsNon.RecordCount > 0 Then
5                     MsgBox "Kode Langgan : " & Text(57) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut", vbCritical + vbOKOnly, "Master Pelanggan Non Aktif"
6                     pelangganAda = False
7                     Exit Sub
8                 End If
                  
9                 If Left(Trim(Text(57)), 3) = "AFL" Then
10                    MsgBox "Tidak diperkenankan membuat Kode Pelanggan Afiliasi.", vbCritical + vbOKOnly, "Master Pelanggan Salah"
11                    pelangganAda = False
12                    Exit Sub
13                End If
                          
14                Call objHT.OpenTransactionAll
15                If UCase(cboStatus.Text) = "SWASTA" Then
16                    vStatus = "S"
17                Else
18                    vStatus = "P"
19                End If
20                Set rsSearch = objO.FetchDataACC("select KodeLgn from ARMST WITH(NOLOCK) where KodeLgn = '" & Trim(Text(57)) & "'")
21                If rsSearch.BOF Then
22                    If OnTopSV = 0 Then
23                    sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2" _
                           & " ,Telepon,Nofax,Npwp,SyaratBayar,StatusPersh,KodeUser,TglEntry)" _
                          & " VALUES ('" & Text(57) & "','" & Text(1) & "','" & Text(56) & "'" _
                          & ",'" & Text(2) & "','" & Text(3) & "','" & Replace(Replace(Text(4), ">", ""), " ", "") & "','" & Replace(Replace(Text(5), ">", ""), " ", "") & "'" _
                          & ",'" & Text(6) & "','" & Text(7) & "','" & vStatus & "','" & gSA & "','" & Format(isnow, "MM/DD/YYYY") & "')"
24                    Else
                      'UNTUK FEBUARI
25                    sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2" _
                           & " ,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry)" _
                          & " VALUES ('" & Text(57) & "','" & Text(1) & "','" & Text(56) & "'" _
                          & ",'" & Text(2) & "','" & Text(3) & "','" & Replace(Replace(Text(4), ">", ""), " ", "") & "','" & Replace(Replace(Text(5), ">", ""), " ", "") & "'" _
                          & ",'" & Text(6) & "','D0','D0','" & Text(7) & "','D0','D0','" & vStatus & "','" & gSA & "','" & Format(isnow, "MM/DD/YYYY") & "')"
26                    End If
27                Else
28                    If OnTopSV = 0 Then
29                    sSQL = "UPDATE ARMST SET KodeLgn = '" & Text(57) & "', NamaLgn = '" & Text(1) & "'," _
                           & " KodeWil = '" & Text(56) & "',Alamat1 = '" & Text(2) & "',Alamat2 = '" & Text(3) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                          & " Telepon = '" & Replace(Replace(Text(4), ">", ""), " ", "") & "',Nofax = '" & Replace(Replace(Text(5), ">", ""), " ", "") & "',Npwp = '" & Text(6) & "'," _
                          & " SyaratBayar = '" & Text(7) & "',StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "'" _
                          & " where KodeLgn = '" & Trim(Text(57)) & "'"
30                    Else
                      'UNTUK FEBUARI
31                    sSQL = "UPDATE ARMST SET KodeLgn = '" & Text(57) & "', NamaLgn = '" & Text(1) & "'," _
                           & " KodeWil = '" & Text(56) & "',Alamat1 = '" & Text(2) & "',Alamat2 = '" & Text(3) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                          & " Telepon = '" & Replace(Replace(Text(4), ">", ""), " ", "") & "',Nofax = '" & Replace(Replace(Text(5), ">", ""), " ", "") & "',Npwp = '" & Text(6) & "'," _
                          & " TOPGRP = '" & Text(7) & "',StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "'" _
                          & " where KodeLgn = '" & Trim(Text(57)) & "'"
32                    End If
33                End If
34                Call objHT.InsertData_ACC(sSQL)
35                Call objHT.CloseTransactionAll
                  
36                gBantuan = Trim(Text(57))
37                Call hapusForm
38                Text(57) = ""
                  
39                With frmtCustomer
40                    If .Visible Then
41                        .Text(2) = Trim(gBantuan)
42                        Unload Me
43                    End If
44                End With
                  
45                With frmBooking
46                    If .Visible Then
47                        .Text(2) = Trim(gBantuan)
48                        Unload Me
49                    End If
50                End With
              'End If
51        Else
52            MsgBox "Masukkan kode pelanggan Kode lokasi 3 digit (con: SMG=Semarang), 1 digit nama depan customer (con: B=Budi) kemudian beri tanda (-) No urut bila no urutnya 2 maka Hasilnya : SMGB-002 !"
53        End If

54       On Error GoTo 0
55       Exit Sub

cmdSimpan_Click_Error:

56        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmtPlgn")
End Sub

Private Sub cmdtWilayah_Click()
    frmtWil.Show vbModal, Me
End Sub

Private Sub Form_Load()
    Text(57) = ""
    cboStatus.AddItem "Pemerintah", 0
    cboStatus.AddItem "Swasta", 1
    Call hapusForm
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
            Case 3
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEWIL"
                frmCariAcess.Show (vbModal)
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "SYARAT"
                frmCariAcess.Show (vbModal)
            Case 57
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim pelangganAda As Boolean
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 1
3                 If Trim(Text(Index)) <> "" Then
4                     gInNama = AmbilInisial(Text(Index), 2)
5                 Else
6                     Text(Index).SetFocus
7                 End If
8             Case 7
9                     Set rsSearch = objO.FetchDataACC("select kodesyarat,KetSyaratBayar from arsyarat WITH(NOLOCK) where kodesyarat = '" & Text(Index) & "'")
10                    If rsSearch.BOF Then
11                        Cancel = True
12                        gMainKeyCari = Text(Index)
13                        gSearchOption = "SYARAT"
14                        frmCariAcess.Show (vbModal)
15                    Else
16                        Text(0) = rsSearch("KetSyaratBayar")
17                    End If
18            Case 57
19                If Trim(Text(Index)) <> "" Then
20                    pelangganAda = False
21                    Text(56) = Left(Text(57), 3)
                          
                          '20/12/2012  Cust Non Axtif
                          Dim rsNon As ADODB.Recordset
22                        Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & Text(Index) & "' and nonaktifflag  = 1")
23                        If rsNon.RecordCount > 0 Then
24                            MsgBox "Kode Langgan : " & Text(Index) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut", vbCritical + vbOKOnly, "Master Pelanggan Non Aktif"
25                            pelangganAda = False
26                            Exit Sub
27                        End If
                          
28                        If OnTopSV = 0 Then
                              'Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp, Nofax, SyaratBayar, StatusPersh, KodeSales from Armst where KodeLgn = '" & text(Index) & "'")
                              '20/12/2012  Cust Non Axtif
29                            Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp, Nofax, SyaratBayar, StatusPersh, KodeSales from Armst WITH(NOLOCK) where KodeLgn = '" & Text(Index) & "' and nonaktifflag  = 0")
30                        Else
                              'UNTUK FEBUARI - TOP PELANGGAN
31                            Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp, Nofax, TOPGRP as SyaratBayar, StatusPersh, KodeSales from Armst WITH(NOLOCK) where KodeLgn = '" & Text(Index) & "' and nonaktifflag  = 0")
32                        End If
                          
33                        If rsSearch.BOF Then
34                            If gBantuan2 <> "kosong" Then
35                                Call hapusForm
36                            End If
37                            Text(56) = Left(Text(57), 3)
38                            pelangganAda = False
39                        Else
40                            pelangganAda = True
41                            If IsNull(rsSearch("NamaLgn")) Then
42                                Text(1) = " "
43                            Else
44                                Text(1) = Replace(Trim(rsSearch("NamaLgn")), vbNewLine, "")
45                            End If
46                            Text(2) = Nul(rsSearch("Alamat1")) & " " & Nul(rsSearch("Alamat2"))
47                            If IsNull(rsSearch("Alamat3")) Then
48                                Text(3) = " "
49                            Else
50                                Text(3) = rsSearch("Alamat3")
51                            End If
52                            If IsNull(rsSearch("Telepon")) Then
53                                Text(4) = " "
54                            Else
55                                Text(4) = rsSearch("Telepon")
56                            End If
57                            If IsNull(rsSearch("Npwp")) Then
58                                Text(6) = " "
59                            Else
60                                Text(6) = rsSearch("Npwp")
61                            End If
62                            If IsNull(rsSearch("KodeWil")) Then
63                                Text(56) = " "
64                            Else
65                                Text(56) = rsSearch("KodeWil")
66                            End If
67                            If IsNull(rsSearch("Nofax")) Then
68                                Text(5) = " "
69                            Else
70                                Text(5) = rsSearch("Nofax")
71                            End If
72                            If IsNull(rsSearch("SyaratBayar")) Or Trim(rsSearch("SyaratBayar")) = "" Then
73                                Text(7) = "D0"
74                            Else
75                                Text(7) = rsSearch("SyaratBayar")
76                            End If
77                            If IsNull(rsSearch("StatusPersh")) Then
78                                cboStatus.Text = ""
79                            Else
80                                If UCase(Trim(rsSearch("StatusPersh"))) = "P" Then
81                                    cboStatus.ListIndex = 0
82                                Else
83                                    cboStatus.ListIndex = 1
84                                End If
85                            End If
86                            Text(1).BackColor = &HC0E0FF
87                        End If
88                        Set rsSearch = objO.FetchDataACC("select KodeWil,NamaWil from Arwilayah WITH(NOLOCK) where KodeWil = '" & Left(Text(57), 3) & "'")
89                        If rsSearch.BOF Then
90                            MsgBox "Kode Wilayah Salah !", vbOKOnly, "Peringatan"
91                            MsgBox "Masukkan kode pelanggan Kode lokasi 3 digit (con: SMG=Semarang), 1 digit nama depan customer (con: B=Budi) kemudian beri tanda (-) No urut bila no urutnya 2 maka Hasilnya : SMGB-002 !"
92                            Text(57) = ""
93                        Else
94                            Text(3) = rsSearch("NamaWil")
95                        End If
                          
96                        Set rsSearch = objO.FetchDataACC("select KodeSyarat,KetSyaratBayar from ARSYARAT WITH(NOLOCK) where KodeSyarat = '" & Text(7) & "'")
97                        If Not rsSearch.BOF Then
98                            Text(0) = rsSearch("KetSyaratBayar")
99                        End If
                      ' CEK BULAN LALU
100                   If Not pelangganAda Then
101                           If OnTopSV = 0 Then
102                               Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp, Nofax, SyaratBayar, StatusPersh, KodeSales from Armst WITH(NOLOCK) where KodeLgn = '" & Text(Index) & "' and nonaktifflag  = 0")
103                           Else
                                  'UNTUK FEBUARI - TOP PELANGGAN
104                               Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp, Nofax, TOPGRP as SyaratBayar, StatusPersh, KodeSales from Armst WITH(NOLOCK) where KodeLgn = '" & Text(Index) & "' and nonaktifflag  = 0")
105                           End If
                              
106                           If rsSearch.BOF Then
107                               If gBantuan2 <> "kosong" Then
108                                   Call hapusForm
109                               End If
110                               Text(56) = Left(Text(57), 3)
111                           Else
112                               MsgBox "Kode Tersebut dipakai di bulan lalu !", vbOKOnly, "Duplikasi Kode Langgan"
113                               If IsNull(rsSearch("NamaLgn")) Then
114                                   Text(1) = " "
115                               Else
116                                   Text(1) = rsSearch("NamaLgn")
117                               End If

118                               Text(2) = Nul(rsSearch("Alamat1")) & " " & Nul(rsSearch("Alamat2"))

119                               If IsNull(rsSearch("Alamat3")) Then
120                                   Text(3) = " "
121                               Else
122                                   Text(3) = rsSearch("Alamat3")
123                               End If
124                               If IsNull(rsSearch("Telepon")) Then
125                                   Text(4) = " "
126                               Else
127                                   Text(4) = rsSearch("Telepon")
128                               End If
129                               If IsNull(rsSearch("Npwp")) Then
130                                   Text(6) = " "
131                               Else
132                                   Text(6) = rsSearch("Npwp")
133                               End If
134                               If IsNull(rsSearch("KodeWil")) Then
135                                   Text(56) = " "
136                               Else
137                                   Text(56) = rsSearch("KodeWil")
138                               End If
139                               If IsNull(rsSearch("Nofax")) Then
140                                   Text(5) = " "
141                               Else
142                                   Text(5) = rsSearch("Nofax")
143                               End If
144                               If IsNull(rsSearch("SyaratBayar")) Or Trim(rsSearch("SyaratBayar")) = "" Then
145                                   Text(7) = "D0"
146                               Else
147                                   Text(7) = rsSearch("SyaratBayar")
148                               End If
149                               If IsNull(rsSearch("StatusPersh")) Then
150                                   cboStatus.Text = ""
151                               Else
152                                   If UCase(Trim(rsSearch("StatusPersh"))) = "P" Then
153                                       cboStatus.ListIndex = 0
154                                   Else
155                                       cboStatus.ListIndex = 1
156                                   End If
157                               End If
158                               Text(1).BackColor = &HC0E0FF
159                           End If
                              
160                           Set rsSearch = objO.FetchDataACC("select KodeWil,NamaWil from Arwilayah WITH(NOLOCK) where KodeWil = '" & Left(Text(57), 3) & "'")
161                           If rsSearch.BOF Then
162                               MsgBox "Kode Wilayah Salah !", vbOKOnly, "Peringatan"
163                               MsgBox "Masukkan kode pelanggan Kode lokasi 3 digit (con: SMG=Semarang), 1 digit nama depan customer (con: B=Budi) kemudian beri tanda (-) No urut bila no urutnya 2 maka Hasilnya : SMGB-002 !"
164                               Text(57) = ""
165                           Else
166                               Text(3) = rsSearch("NamaWil")
167                           End If
                              
168                           Set rsSearch = objO.FetchDataACC("select KodeSyarat,KetSyaratBayar from ARSYARAT WITH(NOLOCK) where KodeSyarat = '" & Text(7) & "'")
                              
169                           If Not rsSearch.BOF Then
170                               Text(0) = rsSearch("KetSyaratBayar")
171                           End If
172                   End If
173               End If
174       End Select
175      On Error GoTo 0
176      Exit Sub

text_Validate_Error:

177       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmtPlgn")
End Sub

Private Sub hapusForm()
          Dim i As Integer
          
1        On Error GoTo hapusForm_Error

2         For i = 0 To 7
3             Select Case i
                  Case 4, 5
4                     Text(i) = ">>>>->>>>>>>>>>"
5                 Case Else
6                     Text(i) = ""
7             End Select
8         Next
9         Text(56) = " "
10        cboStatus = "Swasta"
          
          
          Dim rsHlp As New ADODB.Recordset
           
11            Set rsHlp = objO.FetchDataACC("SELECT KODESYARAT, KETSYARATBAYAR FROM ArSyarat WITH(NOLOCK) WHERE KODESYARAT = 'D0'")
12            If Not rsHlp.EOF Then
13                Text(7) = Trim(rsHlp("KodeSyarat"))
14                Text(0) = Trim(rsHlp("KetSyaratBayar"))
15            End If

16       On Error GoTo 0
17       Exit Sub

hapusForm_Error:

18        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmtPlgn")

End Sub

