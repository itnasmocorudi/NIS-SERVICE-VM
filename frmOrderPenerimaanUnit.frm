VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form FrmOrderPenerimaanUnit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERBAIKAN PANGGILAN SELESAI SERVICE"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   12600
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4920
      TabIndex        =   11
      Top             =   3960
      Width           =   810
      _ExtentX        =   1429
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3840
      TabIndex        =   10
      Top             =   3960
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
      Height          =   5205
      Left            =   -600
      TabIndex        =   0
      Top             =   360
      Width           =   11800
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   3735
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   4750
         TabIndex        =   15
         Top             =   1320
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
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   "??"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   56
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   57
         Left            =   3600
         TabIndex        =   3
         Top             =   720
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   "CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   1020
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
         Left            =   2640
         TabIndex        =   6
         Top             =   1600
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
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
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
         Left            =   2640
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   15
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
         Index           =   2
         Left            =   4560
         TabIndex        =   8
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin AIFCmp1.asxPowerButton asxPowerButton2 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   1320
         TabIndex        =   12
         Top             =   3600
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Ganti"
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   1500
         TabIndex        =   21
         Top             =   2280
         Width           =   990
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
         Left            =   1545
         TabIndex        =   20
         Top             =   1920
         Width           =   930
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
         Left            =   1080
         TabIndex        =   19
         Top             =   1035
         Width           =   1425
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
         Left            =   1695
         TabIndex        =   18
         Top             =   1605
         Width           =   750
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
         Left            =   1110
         TabIndex        =   17
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENERIMAAN UNIT"
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
         Left            =   4300
         TabIndex        =   14
         Top             =   120
         Width           =   2600
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BorderColor     =   &H80000002&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   15
         Width           =   11200
      End
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   13
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
Attribute VB_Name = "FrmOrderPenerimaanUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private noPol As String
Private isBoleh As Boolean
Private rsCustomer As ADODB.Recordset
Private pesanError As String

Private isDealerNewEnkrip As Boolean

Private Sub asxPowerButton2_Click()
    Call enableText
    text(3).SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo vb_error
    isBoleh = True
    pesanError = ""
    If Trim(text(3)) = "" Then
        isBoleh = False
        pesanError = pesanError & " Nama Customer harus diisi !" & vbCrLf
    End If
    If Trim(text(10)) = "" Then
        isBoleh = False
        pesanError = pesanError & " Model Kendaraan harus diisi !" & vbCrLf
    End If
    If Trim(text(9)) = "" Then
        isBoleh = False
        pesanError = pesanError & " Tahun Kendaraan harus diisi!" & vbCrLf
    End If
    If Trim(text(1)) = "" And Trim(text(2)) = "" Then
        isBoleh = False
        pesanError = pesanError & " No Telepon harus diisi !" & vbCrLf
    End If
    If Text1 = "" Then
        isBoleh = False
        pesanError = pesanError & " Keterangan harus diisi!" & vbCrLf
    End If
    If isBoleh Then
        Call objO.InsertData("INSERT INTO UNITCOME (NamaCustomer,JenisKendaraan,TahunKendaraan,NoPolisi,NoTelp1,NoTelp2,TanggalIn,UserIn,Keterangan) VALUES ('" & Trim(text(3)) & "','" & Trim(text(10)) & "','" & Trim(text(9)) & "','" & noPol & "','" & Trim(text(1)) & "','" & Trim(text(2)) & "',getdate(),'" & gSA & "','" & Text1.text & "')")
        MsgBox "Data berhasil disimpan!", vbInformation, "Informasi Penambahan"
        Call hapusForm
        text(0) = ""
        text(0).Mask = "??"
    
        text(56) = ""
        text(56).Mask = "9999"
        
        text(57) = ""
        text(57).Mask = "CCC"
    Else
        MsgBox pesanError, vbCritical, "Kesalahan Menyimpan"
        text(3).SetFocus
    End If
    Exit Sub
vb_error:
    MsgBox "Terjadi kesalahan menyimpan data", vbCritical, "Kesalahan Menyimpan"
    'Text(3).SetFocus
End Sub

Private Sub Form_Load()
    isBoleh = True
    
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

    Select Case Index
        Case 31
            VScroll1.value = 2500
        Case 32
            VScroll1.value = 2700
        Case 28
            VScroll1.value = 2000
    End Select
        
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        If Index = 10 And text(10).Enabled = True Then
            gMainKeyCari = text(Index)
            gSearchOption = "MDL"
            frmCari.Show (vbModal)
        End If
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii = 13 Then
        Call text_Validate(Index, True)
        Sendkeys ("{TAB}")
    End If
    
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 57
            If text(0) <> "" And text(56) Then
                Call hapusForm
                noPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
                If noPol <> "" Then
                    Set rsCustomer = objO.FetchData("SELECT * FROM CUSTDAT WHERE NO_POLISI = '" & Trim(noPol) & "'")
                    If Not rsCustomer.EOF Then
                        text(3) = Trim(rsCustomer("CUSTOMER"))
                        text(10) = Trim(rsCustomer("MDL"))
                        text(9) = Trim(rsCustomer("FS_TH_MOBIL"))
                        text(1) = Trim(rsCustomer("GSM"))
                        text(2) = Trim(rsCustomer("CDMA"))
                    Else
                        Call enableText
                    End If
                End If
            End If
    End Select
End Sub

Private Sub enableText()
    text(3).Enabled = True
    text(3).BackColor = vbWhite
    
    text(10).Enabled = True
    text(10).BackColor = vbWhite
    
    text(9).Enabled = True
    text(9).BackColor = vbWhite
    
    text(1).Enabled = True
    text(1).BackColor = vbWhite
    
    text(2).Enabled = True
    text(2).BackColor = vbWhite
    
End Sub

Private Sub hapusForm()
'   text(0) = ""
'   text(0).Mask = "??"
'
'   text(56) = ""
'   text(56).Mask = "9999"
'
'   text(57) = ""
'   text(57).Mask = "CCC"
   
   For i = 1 To 3
        text(1) = ""
        Select Case i
            Case 1, 2
                text(i) = ""
                text(i).Mask = "999999999999999"
                
            Case 3
                text(i) = ""
                text(i).Mask = ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
        End Select
        text(i).BackColor = &HC0E0FF
        text(i).Enabled = False
   Next
   
   text(9) = ""
   text(9).Mask = "9999"
   text(9).BackColor = &HC0E0FF
   text(9).Enabled = False
   
   text(10) = ""
   text(10).Mask = ">CCCCCCCCCCCCCCCCCCCC"
   text(10).BackColor = &HC0E0FF
   text(10).Enabled = False
   
   Text1.text = ""
End Sub
