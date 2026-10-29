VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmConfirmPass 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame KabPass 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdBtlPass 
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2040
         TabIndex        =   4
         Top             =   1440
         Width           =   1050
      End
      Begin VB.CommandButton tblOkPass 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1215
         TabIndex        =   0
         Top             =   240
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   2400
         _ExtentX        =   4233
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   8
         Top             =   1005
         Width           =   690
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama "
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
         Left            =   615
         TabIndex        =   7
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Petugas"
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
         TabIndex        =   6
         Top             =   285
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmConfirmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fPassword As String
Dim nTipe As String

Private Sub cmdBtlPass_Click()
    Unload Me
    frmOrder.opByr(1).value = True
End Sub

Private Sub tblOkPass_Click()
    Dim nPass As String
    Dim rsHlp As ADODB.Recordset
        
    nPass = Replace(txtPass, "'", "")
    If (UCase(nPass) = fPassword And text(1) <> "" And (nTipe = "KACAB" Or nTipe = "KABENG")) Then
        frmOrder.opByr(2).value = True
        gVerifikator = Trim(text(0))
        
        ' CEK SISA LIMIT
        If nTipe = "KACAB" Then
            Set rsHlp = objO.FetchData("SELECT LIMITKACAB AS LIMIT, HUTANGKACAB - BAYARKACAB AS SISA FROM T_PROFIL WITH(NOLOCK)")
        ElseIf nTipe = "KABENG" Then
            Set rsHlp = objO.FetchData("SELECT LIMITKABENG AS LIMIT, HUTANGKABENG - BAYARKABENG AS SISA FROM T_PROFIL WITH(NOLOCK)")
        End If
        If Not rsHlp.EOF Then
            If CDbl(rsHlp("LIMIT")) <= CDbl(rsHlp("SISA")) + CDbl(frmOrder.text(61)) Then
                MsgBox "Limit tagihan " & nTipe & " sudah terlampaui. " & vbCr & "Harap melakukan pelunasan terlebih dahulu. "
                frmOrder.opByr(0).value = True
                gVerifikator = ""
            End If
        End If
        Unload Me
    Else
        MsgBox "Password Tidak sesuai !!!", vbCritical, "SALAH PASSWORD"
        text(0) = ""
        text(1) = ""
        txtPass = ""
        text(0).SetFocus
        frmOrder.opByr(1).value = True
    End If
End Sub

Private Sub text_GotFocus(Index As Integer)
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0
            sSQL = "SELECT * from t_petugas WITH(NOLOCK) where fs_kd_petugas ='" & text(Index) & "' AND (TIPE = 'KABENG' OR TIPE = 'KACAB')"
            Set rsSearch = objO.FetchData(sSQL)
            If Not rsSearch.BOF Then
                text(Index) = RTrim(rsSearch("fs_kd_petugas"))
                text(1) = RTrim(rsSearch("fs_nm_petugas"))
                nTipe = rsSearch("tipe")
                fPassword = RTrim(Mid(rsSearch("fs_kd_pass"), 3, 10))
            Else
                text(1) = ""
                fPassword = ""
            End If
    End Select
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub
