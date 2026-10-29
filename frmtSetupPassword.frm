VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmSetupPassword 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Setup Password"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6660
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   11340
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
         Index           =   3
         Left            =   2310
         MaxLength       =   15
         TabIndex        =   16
         Top             =   960
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
         Index           =   2
         Left            =   2310
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1335
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
         Top             =   2070
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
         Top             =   1710
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   1
         Left            =   2310
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   15
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
         Caption         =   "No HP"
         Height          =   195
         Index           =   4
         Left            =   1725
         TabIndex        =   17
         Top             =   1020
         Width           =   480
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1395
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   683
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Verifikasi"
         Height          =   195
         Index           =   1
         Left            =   1575
         TabIndex        =   8
         Top             =   2130
         Width           =   630
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Index           =   0
         Left            =   1515
         TabIndex        =   7
         Top             =   1770
         Width           =   690
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   30
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9405
      TabIndex        =   5
      Top             =   7095
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
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
      Height          =   360
      Left            =   8190
      TabIndex        =   4
      Top             =   7095
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
   Begin VB.Image asxPowerBanner1 
      Height          =   7695
      Left            =   0
      Picture         =   "frmtSetupPassword.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmSetupPassword"
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
1         If Len(Trim(Text(0))) > 5 Then
2             If UCase(Text(2)) = Trim(Mid(PwdLama, 3, 10)) Then
3                 If UCase(Text(0)) = UCase(Text(1)) Then
4                     If UCase(Text(0)) <> UCase(Text(2)) Then
5                         Set rsSearch = objO.FetchData("Select * from T_PETUGAS WITH(NOLOCK) where FS_KD_PETUGAS='" & gSA & "'")
                          
6                         Call ClearsSql
7                         gsSql1(0) = "FS_KD_PASS"
8                         gsSql2(0) = Left(PwdLama, 2) + UCase(Text(0))
9                         gsSql1(1) = "TPPRINT"
10                        gsSql2(1) = Trim(txtEntry(2))
11                        gsSql1(2) = "TGLUPDATE"
12                        gsSql2(2) = Format(isnow, "YYYY-MM-DD HH:MM:SS")
13                        gsSql1(3) = "NOHP"
14                        gsSql2(3) = NulSID(Text(3))
                          
15                        If Not rsSearch.BOF Then
16                            sSQL = GensSql("T_PETUGAS", True) & " where FS_KD_PETUGAS='" & gSA & "'"
                              
17                        Else
18                            sSQL = GensSql("T_PETUGAS", False)
19                        End If
20                        Call objO.UpdateData(sSQL)
                          
21                        If Log.RPLHistori = 1 And gUstat = 0 Then
22                            Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..T_PETUGAS WITH(NOLOCK) where FS_KD_PETUGAS='" & gSA & "' and dealer = '" & gDb & "'")
23                            Call ClearsSql
24                            gsSql1(0) = "FS_KD_PASS"
25                            gsSql2(0) = Left(PwdLama, 2) + UCase(Text(0))
26                            gsSql1(1) = "TPPRINT"
27                            gsSql2(1) = Trim(txtEntry(2))
28                            gsSql1(2) = "TGLUPDATE"
29                            gsSql2(2) = Format(isnow, "YYYY-MM-DD HH:MM:SS")
30                            gsSql1(3) = "FS_KD_PETUGAS"
31                            gsSql2(3) = NulSID(gSA)
32                            gsSql1(4) = "dealer"
33                            gsSql2(4) = NulSID(gDb)
34                            gsSql1(5) = "NOHP"
35                            gsSql2(5) = NulSID(Text(3))
                              
36                            If Not rsSearch.BOF Then
37                                sSQL = GensSql("[" & dbReplikasiHO & "]..T_PETUGAS", True) & " where FS_KD_PETUGAS='" & gSA & "' and dealer = '" & gDb & "'"
                                  
38                            Else
39                                sSQL = GensSql("[" & dbReplikasiHO & "]..T_PETUGAS", False)
40                            End If
41                        End If
                           
42                        Text(0) = ""
43                        Text(1) = ""
44                        Text(2) = ""
                          
45                        Set rsSearch = objO.FetchData("Select * from T_PETUGAS WITH(NOLOCK) where FS_KD_PETUGAS='" & gSA & "'")
46                        If Not rsSearch.BOF Then
47                            txtEntry(1) = rsSearch("FS_NM_PETUGAS")
48                            PwdLama = RTrim(rsSearch("fs_kd_pass"))
49                        End If
50                        MsgBox "Password Anda telah berhasil diganti !", vbOKOnly, "Informasi"
51                    Else
52                        MsgBox "Password Baru Tidak boleh sama dengan Password Lama !", vbOKOnly, "Kesalahan"
53                    End If
54                Else
55                    MsgBox "Password dan Verifikasi tidak sama !", vbOKOnly, "Kesalahan"
56                End If
57            Else
58                MsgBox "Password Lama tidak sama !", vbOKOnly, "Kesalahan"
59            End If
60        Else
61            MsgBox "Password Baru minimal 6 karakter !", vbOKOnly, "Kesalahan"
62        End If
End Sub

Private Sub Form_Activate()
Dim rsSearch As ADODB.Recordset

    txtEntry(0) = gSA
    Set rsSearch = objO.FetchData("Select * from T_PETUGAS WITH(NOLOCK) where FS_KD_PETUGAS='" & gSA & "'")
    If Not rsSearch.BOF Then
        txtEntry(1) = rsSearch("FS_NM_PETUGAS")
        Text(3) = NulSID(rsSearch("nohp"))
        PwdLama = RTrim(rsSearch("fs_kd_pass"))
        txtEntry(2) = Trim(rsSearch("TPPRINT"))
    End If
     
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
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
