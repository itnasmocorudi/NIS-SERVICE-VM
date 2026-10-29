VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmtModelCbg 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Master Model"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11880
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6660
      Left            =   450
      TabIndex        =   6
      Top             =   405
      Width           =   11340
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2190
         TabIndex        =   0
         Top             =   465
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   2190
         TabIndex        =   1
         Top             =   765
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   2190
         TabIndex        =   3
         Top             =   1770
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   2190
         TabIndex        =   4
         Top             =   2070
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   2190
         TabIndex        =   2
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   2190
         TabIndex        =   7
         Top             =   1400
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "MDL"
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
         Left            =   1710
         TabIndex        =   12
         Top             =   540
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Keterangan"
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
         Left            =   1170
         TabIndex        =   11
         Top             =   840
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4320
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Kelas"
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
         Left            =   1635
         TabIndex        =   10
         Top             =   1140
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4440
         Top             =   1770
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Group Model"
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
         Left            =   1110
         TabIndex        =   9
         Top             =   1830
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frame"
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
         Left            =   1560
         TabIndex        =   8
         Top             =   2130
         Width           =   450
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3120
         Top             =   1080
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7395
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13044
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Tabel Model Kendaraan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   8190
      TabIndex        =   13
      Top             =   7155
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Hapus"
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
      Left            =   9405
      TabIndex        =   14
      Top             =   7155
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   6960
      TabIndex        =   5
      Top             =   7155
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
End
Attribute VB_Name = "frmtModelCbg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub BtnTutup_Click()
    Unload Me
End Sub

Private Sub CboModel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombolCombo(KeyCode)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        sSQL = "UPDATE model SET FD_TGL_DEL = '" & Format(gTglServer, "MM/DD/YYYY HH:MM:SS") & "', FS_LOG_DEL = '" & gSA & "' where mdl='" & text(0) & "' and FS_LOG_DEL = ''"
        oConn.Execute sSQL
        text(0) = ""
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As New ADODB.Recordset

    If text(2) <> "" Then
        If rsSearch.State Then rsSearch.Close
        sSQL = "Select * from model where mdl='" & text(0) & "'"
        rsSearch.Open sSQL, oConn, adOpenStatic
        
        Call ClearsSql
        gsSql1(0) = "MDL"
        gsSql2(0) = text(0)
        gsSql1(1) = "[Desc]"
        gsSql2(1) = text(1)
        gsSql1(2) = "MODEL"
        gsSql2(2) = text(2)
        gsSql1(3) = "FS_GRP_MDL"
        gsSql2(3) = text(3)
        gsSql1(4) = "NO_CHASSIS"
        gsSql2(4) = text(4)
        gsSql1(5) = "ISICABANG"
        gsSql2(5) = 1
        
        If Not rsSearch.BOF Then
            sSQL = GensSql("model", True) & " where mdl='" & text(0) & "'"
        Else
            sSQL = GensSql("model", False)
        End If
        oConn.Execute sSQL
        rsSearch.Close
        text(0) = ""
        Call hapusForm
        text(0).SetFocus
    Else
        MsgBox "Masukkan Model !"
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
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
                gMainKeyCari = text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
            Case 2
                gMainKeyCari = text(Index)
                gSearchOption = "KELAS_MOBIL"
                frmCari.Show (vbModal)
            Case 3
                gMainKeyCari = text(Index)
                gSearchOption = "GROUP_MODEL"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As New ADODB.Recordset

    If rsSearch.State Then rsSearch.Close
    Select Case Index
        Case 0
            If text(Index) <> "" Then
                sSQL = "Select * from model a left join carclass b on a.model = b.classcode and B.FS_LOG_DEL = '' where MDL='" & text(0) & "'"
                rsSearch.Open sSQL, oConn, adOpenStatic
                If Not rsSearch.EOF Then
                    text(1) = rsSearch("DESC")
                    text(2) = rsSearch("MODEL")
                    text(3) = rsSearch("FS_GRP_MDL")
                    text(4) = rsSearch("NO_CHASSIS")
                    text(5) = Nul(rsSearch("ClassName"))
                Else
                    Call hapusForm
                End If
                rsSearch.Close
            End If
        Case 2
            If text(Index) <> "" Then
                sSQL = "select * from carclass WHERE classcode = '" & text(Index) & "' and FS_LOG_DEL = ''"
                rsSearch.Open sSQL, oConn, adOpenStatic
                If rsSearch.EOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "KELAS_MOBIL"
                    frmCari.Show (vbModal)
                Else
                    text(5) = rsSearch("className")
                End If
                rsSearch.Close
            End If
        Case 3
            If text(2) <> "" Then
                sSQL = "select distinct FS_GRP_MDL from MODEL where FS_GRP_MDL = '" & text(3) & "'"
                rsSearch.Open sSQL, oConn, adOpenStatic
                If rsSearch.BOF Then
'                    Cancel = True
'                    gMainKeyCari = Text(Index)
'                    gSearchOption = "GROUP_MODEL"
'                    frmCariHO.Show (vbModal)
                End If
                rsSearch.Close
            End If
    End Select
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 1 To 5
        text(i) = ""
    Next
End Sub

