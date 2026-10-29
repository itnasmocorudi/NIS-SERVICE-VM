VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtWil 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Wilayah"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmtWil.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   -480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6660
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   11340
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1755
         TabIndex        =   1
         Top             =   1065
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   56
         Left            =   1755
         TabIndex        =   0
         Top             =   750
         Width           =   435
         _ExtentX        =   767
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2220
         TabIndex        =   4
         Top             =   750
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   6615
         Left            =   0
         Top             =   0
         Width           =   11295
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Wilayah"
         Height          =   255
         Left            =   555
         TabIndex        =   5
         Top             =   825
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Wilayah"
         Height          =   195
         Left            =   585
         TabIndex        =   3
         Top             =   1140
         Width           =   1035
      End
   End
   Begin ucXPButton.XPButton cmdSimpan 
      Height          =   495
      Left            =   9360
      TabIndex        =   7
      Top             =   6840
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
      MICON           =   "frmtWil.frx":2197
      PICN            =   "frmtWil.frx":21B3
      PICH            =   "frmtWil.frx":26C6
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
      Left            =   10680
      TabIndex        =   8
      Top             =   6840
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
      MICON           =   "frmtWil.frx":2C68
      PICN            =   "frmtWil.frx":2C84
      PICH            =   "frmtWil.frx":3158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   11880
      X2              =   600
      Y1              =   7440
      Y2              =   7440
   End
End
Attribute VB_Name = "frmtWil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

    If text(56) <> "" Then
            Set rsSearch = objO.FetchDataACC("select NamaWil from arwilayah WITH(NOLOCK) where KodeWil = '" & Trim(text(56)) & "'")
            If rsSearch.BOF Then
                Call objO.InsertDataACC("INSERT INTO arwilayah(KodeWil,NamaWil)  " & _
                                    "VALUES ('" & text(56) & "','" & text(1) & "')")
            Else
                MsgBox "Kode wilayah sudah ada !"
            End If
        
        text(56) = ""
        text(1) = ""
        text(56).SetFocus
    End If
End Sub

Private Sub Form_Load()
    text(56) = ""
    text(1) = ""
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 56
                gMainKeyCari = text(Index)
                gSearchOption = "KODEWIL"
                frmCariAcess.Show (vbModal)
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
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 56
            If Trim(text(Index)) <> "" Then
                    Set rsSearch = objO.FetchDataACC("select KodeWil, NamaWil from arwilayah WITH(NOLOCK) where KodeWil = '" & Trim(text(56)) & "'")
                    If Not rsSearch.BOF Then
                        If IsNull(rsSearch("NamaWil")) Then
                            text(1) = " "
                        Else
                            text(1) = rsSearch("NamaWil")
                        End If
                    End If
                
            End If
    End Select
End Sub
