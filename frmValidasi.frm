VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmValidasi 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Validasi"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameValid 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdValid 
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&O"
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
         MICON           =   "frmValidasi.frx":0000
         PICN            =   "frmValidasi.frx":001C
         PICH            =   "frmValidasi.frx":04BD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCancel 
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   1920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&c"
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
         MICON           =   "frmValidasi.frx":09D7
         PICN            =   "frmValidasi.frx":09F3
         PICH            =   "frmValidasi.frx":0F37
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Validasi Berwenang ::."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   615
         TabIndex        =   7
         Top             =   120
         Width           =   2910
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   360
         TabIndex        =   6
         Top             =   645
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   360
         TabIndex        =   5
         Top             =   1365
         Width           =   885
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   360
         TabIndex        =   4
         Top             =   1005
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   2175
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   3855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   2775
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmValidasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fPassword As String
Dim nTipe As String

Private Sub cmdCancel_Click()
    frmBilling.Tag = ""
    Unload Me
End Sub

Private Sub cmdValid_Click()
Dim nPass As String
        
    nPass = Replace(txtPass, "'", "")
    If (UCase(nPass) = fPassword And Text(1) <> "" And (nTipe = "KACAB" Or nTipe = "KABENG")) Then
        gVerifikator = Trim(Text(0))
        If Me.Tag = "ADJ" Then
            'frmBilling.Tag = "valid"
            gBantuan = "valid"
        ElseIf Me.Tag = "DISC" Then
            'frmBilling.Tag = "validDisc"
            gBantuan = "validDisc"
        End If
        Unload Me
    Else
        MsgBox "Password Tidak sesuai !!!", vbCritical, "SALAH PASSWORD"
        Text(0) = ""
        Text(1) = ""
        txtPass = ""
        Text(0).SetFocus
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE FS_KD_PETUGAS = '" & Trim(Text(Index)) & "' AND (TIPE = 'KABENG' OR TIPE = 'KACAB')")
                If Not rsSearch.EOF Then
                    Text(Index) = RTrim(rsSearch("fs_kd_petugas"))
                    Text(1) = RTrim(rsSearch("fs_nm_petugas"))
                    nTipe = rsSearch("tipe")
                    fPassword = RTrim(Mid(rsSearch("fs_kd_pass"), 3, 10))
                Else
                    MsgBox "User tidak valid !!!"
                    Text(Index) = ""
                    Text(1) = ""
                    frmBilling.Tag = ""
                End If
                rsSearch.Close
            End If
    End Select
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub
