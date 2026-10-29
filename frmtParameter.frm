VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFCE1.OCX"
Begin VB.Form frmtparameter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5505
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9990
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4635
      Left            =   450
      TabIndex        =   3
      Top             =   345
      Width           =   9525
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   2
         Left            =   2040
         TabIndex        =   0
         Top             =   540
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   4
         Left            =   2040
         TabIndex        =   6
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   5
         Left            =   2040
         TabIndex        =   8
         Top             =   1140
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   6
         Left            =   2040
         TabIndex        =   13
         Top             =   2040
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   11
         Left            =   2040
         TabIndex        =   15
         Top             =   1740
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox txtEntry 
         Height          =   270
         Index           =   12
         Left            =   2040
         TabIndex        =   17
         Top             =   1440
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nomor Surat Ijin"
         Height          =   195
         Left            =   855
         TabIndex        =   18
         Top             =   1515
         Width           =   1125
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nomor Memo"
         Height          =   195
         Left            =   1035
         TabIndex        =   16
         Top             =   1815
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setting Default Printer"
         Height          =   195
         Left            =   420
         TabIndex        =   14
         Top             =   2115
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Batas Maks. Masuk Kerja"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "( Jam : Menit )"
         Height          =   195
         Left            =   2715
         TabIndex        =   11
         Top             =   285
         Width           =   990
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nomor Surat Cuti"
         Height          =   195
         Left            =   780
         TabIndex        =   9
         Top             =   1215
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nomor Surat Lembur"
         Height          =   195
         Left            =   525
         TabIndex        =   7
         Top             =   915
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "( Jam : Menit )"
         Height          =   195
         Left            =   2715
         TabIndex        =   5
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Toleransi Masuk Kerja"
         Height          =   195
         Left            =   405
         TabIndex        =   4
         Top             =   615
         Width           =   1575
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   5565
      Left            =   0
      Top             =   -60
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   9816
      StartColor      =   16711680
      EndColor        =   0
      FormatString    =   "Tabel Parameter"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   65535
   End
   Begin AIFCmp1.asxSysButton BtnTutup 
      Height          =   285
      Left            =   9615
      Top             =   15
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      ButtonStyle     =   0
      ButtonFlags     =   16384
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   30
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   6345
      TabIndex        =   1
      Top             =   5070
      Width           =   1215
      _ExtentX        =   2143
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
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7560
      TabIndex        =   2
      Top             =   5070
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12632256
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
Attribute VB_Name = "frmtparameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub BtnTutup_Click()
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    sSQL = "Select * from T_PROFIL"
    oRs.Open sSQL, oConn, adOpenStatic
    
    Call ClearsSql
    gsSql1(0) = "FD_TOLERANSI"
    gsSql2(0) = "3000-01-01 " & Format(txtEntry(2), "HH:MM")
    gsSql1(1) = "FS_NO_LEMBUR"
    gsSql2(1) = Format(txtEntry(4), "0000000000")
    gsSql1(2) = "FS_NO_CUTI"
    gsSql2(2) = Format(txtEntry(5), "0000000000")
    gsSql1(3) = "FD_MAX_MASUK"
    gsSql2(3) = "3000-01-01 " & Format(txtEntry(1), "HH:MM")
    gsSql1(4) = "FS_PRINTER"
    gsSql2(4) = txtEntry(6)
    gsSql1(5) = "FN_LBR1B"
    gsSql2(5) = Val(txtEntry(0))
    gsSql1(6) = "FN_LBR2B"
    gsSql2(6) = Val(txtEntry(3))
    gsSql1(7) = "FN_LBR3B"
    gsSql2(7) = Val(txtEntry(7))
    gsSql1(8) = "FN_LBR1L"
    gsSql2(8) = Val(txtEntry(8))
    gsSql1(9) = "FN_LBR2L"
    gsSql2(9) = Val(txtEntry(9))
    gsSql1(10) = "FN_LBR3L"
    gsSql2(10) = Val(txtEntry(10))
    gsSql1(11) = "FS_NO_MEMO"
    gsSql2(11) = Format(txtEntry(11), "0000000000")
    gsSql1(12) = "FS_NO_IJIN"
    gsSql2(12) = Format(txtEntry(12), "0000000000")
    
    gsSql1(13) = "FS_LBRJB1"
    gsSql2(13) = txtEntry(13)
    gsSql1(14) = "FS_LBRJB2"
    gsSql2(14) = txtEntry(15)
    gsSql1(15) = "FS_LBRJB3"
    gsSql2(15) = txtEntry(17)
    gsSql1(16) = "FS_LBRTD1"
    gsSql2(16) = txtEntry(14)
    gsSql1(17) = "FS_LBRTD2"
    gsSql2(17) = txtEntry(16)
    gsSql1(18) = "FS_LBRTD3"
    gsSql2(18) = txtEntry(18)
    
    gsSql1(19) = "FS_CTJB1"
    gsSql2(19) = txtEntry(19)
    gsSql1(20) = "FS_CTJB2"
    gsSql2(20) = txtEntry(21)
    gsSql1(21) = "FS_CTJB3"
    gsSql2(21) = txtEntry(23)
    gsSql1(22) = "FS_CTTD1"
    gsSql2(22) = txtEntry(20)
    gsSql1(23) = "FS_CTTD2"
    gsSql2(23) = txtEntry(22)
    gsSql1(24) = "FS_CTTD3"
    gsSql2(24) = txtEntry(24)
    
    gsSql1(25) = "FS_IJJB1"
    gsSql2(25) = txtEntry(25)
    gsSql1(26) = "FS_IJJB2"
    gsSql2(26) = txtEntry(27)
    gsSql1(27) = "FS_IJJB3"
    gsSql2(27) = txtEntry(29)
    gsSql1(28) = "FS_IJTD1"
    gsSql2(28) = txtEntry(26)
    gsSql1(29) = "FS_IJTD2"
    gsSql2(29) = txtEntry(28)
    gsSql1(30) = "FS_IJTD3"
    gsSql2(30) = txtEntry(30)
    
    If Not oRs.BOF Then
        sSQL = GensSql("T_PROFIL", True)
    Else
        sSQL = GensSql("T_PROFIL", False)
    End If
    oConn.Execute sSQL
    oRs.Close
    
End Sub


Private Sub Form_Load()
    Call CenterForm(Me, True)
    
    sSQL = "Select * from T_PROFIL"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        txtEntry(1) = Format(oRs("FD_MAX_MASUK"), "HH:MM")
        txtEntry(2) = Format(oRs("FD_TOLERANSI"), "HH:MM")
        txtEntry(4) = Format(oRs("FS_NO_LEMBUR"), "0000000000")
        txtEntry(5) = Format(oRs("FS_NO_CUTI"), "0000000000")
        txtEntry(6) = oRs("FS_PRINTER")
        txtEntry(0) = Format(oRs("FN_LBR1B"), "0.000000")
        txtEntry(3) = Format(oRs("FN_LBR2B"), "0.000000")
        txtEntry(7) = Format(oRs("FN_LBR3B"), "0.000000")
        txtEntry(8) = Format(oRs("FN_LBR1L"), "0.000000")
        txtEntry(9) = Format(oRs("FN_LBR2L"), "0.000000")
        txtEntry(10) = Format(oRs("FN_LBR3L"), "0.000000")
        txtEntry(11) = Format(oRs("FS_NO_MEMO"), "0000000000")
        txtEntry(12) = Format(oRs("FS_NO_IJIN"), "0000000000")
    
        txtEntry(13) = oRs("FS_LBRJB1")
        txtEntry(14) = oRs("FS_LBRTD1")
        txtEntry(15) = oRs("FS_LBRJB2")
        txtEntry(16) = oRs("FS_LBRTD2")
        txtEntry(17) = oRs("FS_LBRJB3")
        txtEntry(18) = oRs("FS_LBRTD3")
        
        txtEntry(19) = oRs("FS_CTJB1")
        txtEntry(20) = oRs("FS_CTTD1")
        txtEntry(21) = oRs("FS_CTJB2")
        txtEntry(22) = oRs("FS_CTTD2")
        txtEntry(23) = oRs("FS_CTJB3")
        txtEntry(24) = oRs("FS_CTTD3")
    
        txtEntry(25) = oRs("FS_IJJB1")
        txtEntry(26) = oRs("FS_IJTD1")
        txtEntry(27) = oRs("FS_IJJB2")
        txtEntry(28) = oRs("FS_IJTD2")
        txtEntry(29) = oRs("FS_IJJB3")
        txtEntry(30) = oRs("FS_IJTD3")
    End If
    oRs.Close

End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If txtEntry(Index).BackColor = &H808000 Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            txtEntry(Index) = gMainKeyCari
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

Private Sub txtEntry_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    Dim i As Integer
    
    Select Case Index
        Case 1, 2
            txtEntry(Index) = CekJam(txtEntry(Index))

        Case 4, 5
            txtEntry(Index) = Format(Trim(txtEntry(Index)), "0000000000")
'        Case 0, 3, 7, 8, 9, 10
'            txtEntry(Index) = txtEntry(Index)
    End Select
End Sub
