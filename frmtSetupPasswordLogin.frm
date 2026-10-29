VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmSetupPasswordLogin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Setup Password Login"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   120
      TabIndex        =   5
      Top             =   345
      Width           =   5460
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2310
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   975
         Width           =   1815
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2310
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox Text 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2310
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1350
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   1
         Left            =   2310
         TabIndex        =   8
         Top             =   645
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   0
         Left            =   2310
         TabIndex        =   9
         Top             =   315
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   2
         Left            =   6480
         TabIndex        =   3
         Top             =   315
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   8400
         TabIndex        =   14
         Top             =   315
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Printer"
         Height          =   195
         Index           =   3
         Left            =   5580
         TabIndex        =   13
         Top             =   353
         Width           =   810
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Lama"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Petugas/ Group "
         Height          =   195
         Left            =   645
         TabIndex        =   11
         Top             =   353
         Width           =   1605
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Petugas/ Group"
         Height          =   195
         Left            =   615
         TabIndex        =   10
         Top             =   683
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Verifikasi Password Baru"
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   7
         Top             =   1770
         Width           =   1740
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password Baru"
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   6
         Top             =   1410
         Width           =   1065
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   0
      Top             =   30
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
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
   Begin AIFCmp1.asxPowerButton btnBatal 
      Height          =   360
      Left            =   4320
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Batal"
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
End
Attribute VB_Name = "frmSetupPasswordLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PwdLama As String

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

Private Sub btnBatal_Click()
End
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset

    If Len(Trim(Text(0))) > 5 Then
        If UCase(Text(2)) = Trim(Mid(PwdLama, 3, 10)) Then
            If UCase(Text(0)) = UCase(Text(1)) Then
                If UCase(Text(0)) <> UCase(Text(2)) Then
                    Set rsSearch = objO.FetchData("Select * from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS='" & gSA & "'")
                    
                    Call ClearsSql
                    gsSql1(0) = "FS_KD_PASS"
                    gsSql2(0) = Left(PwdLama, 2) + UCase(Text(0))
                    gsSql1(1) = "TGLUPDATE"
                    gsSql2(1) = Format(isnow, "YYYY-MM-DD HH:MM:SS")
                                        
                    If Not rsSearch.BOF Then
                        sSQL = GensSql("T_PETUGAS", True) & " where FS_KD_PETUGAS='" & gSA & "'"
                    Else
                        sSQL = GensSql("T_PETUGAS", False)
                    End If
                    Call objO.UpdateData(sSQL)
                    If Log.RPLHistori = 1 Then
                        sSQL = "UPDATE [" & dbReplikasiHO & "]..T_PETUGAS SET FS_KD_PASS = '" & Left(PwdLama, 2) + UCase(Text(0)) & "',TGLUPDATE = '" & Format(isnow, "YYYY-MM-DD HH:MM:SS") & "' WHERE FS_KD_PETUGAS='" & gSA & "' AND dealer = '" & gDb & "' "
                        Call objO.UpdateData(sSQL)
                    End If
                    
                    MsgBox "Password Anda telah berhasil diganti, silahkan login kembali !", vbOKOnly, "Informasi"
                    End
                Else
                    MsgBox "Password Baru Tidak boleh sama dengan Password Lama !", vbOKOnly, "Kesalahan"
                End If
            Else
                MsgBox "Password dan Verifikasi tidak sama !", vbOKOnly, "Kesalahan"
            End If
        Else
            MsgBox "Password Lama tidak sama !", vbOKOnly, "Kesalahan"
        End If
    Else
        MsgBox "Password Baru minimal 6 karakter !", vbOKOnly, "Kesalahan"
    End If
    
End Sub

Private Sub Form_Activate()
Dim rsSearch As ADODB.Recordset

    txtEntry(0) = gSA
    Set rsSearch = objO.FetchData("Select * from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS='" & gSA & "'")
    If Not rsSearch.BOF Then
        txtEntry(1) = rsSearch("FS_NM_PETUGAS")
        PwdLama = RTrim(rsSearch("fs_kd_pass"))
        txtEntry(2) = Trim(rsSearch("TPPRINT"))
    End If
     
End Sub

Private Sub Form_Load()
    Call hapusForm
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
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

Private Sub hapusForm()
    Text(0) = ""
    Text(1) = ""
    txtEntry(0) = ""
    txtEntry(1) = ""
End Sub


Private Sub txtEntry_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If txtEntry(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            txtEntry(Index) = gMainKeyCari
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

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "TPPRINTER"
                frmCari.show (vbModal)
        End Select
    End If
End Sub
