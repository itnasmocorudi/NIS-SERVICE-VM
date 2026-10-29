VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtMeterai 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Meterai"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11940
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmtMeterai.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6660
      Left            =   435
      TabIndex        =   2
      Top             =   345
      Width           =   11340
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2190
         TabIndex        =   0
         Top             =   465
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   2190
         TabIndex        =   1
         Top             =   765
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   2160
         TabIndex        =   6
         Top             =   1800
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
         MICON           =   "frmtMeterai.frx":211C
         PICN            =   "frmtMeterai.frx":2138
         PICH            =   "frmtMeterai.frx":264B
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
         Left            =   4725
         TabIndex        =   7
         Top             =   1800
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
         MICON           =   "frmtMeterai.frx":2BED
         PICN            =   "frmtMeterai.frx":2C09
         PICH            =   "frmtMeterai.frx":30DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   1800
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
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
         MICON           =   "frmtMeterai.frx":3628
         PICN            =   "frmtMeterai.frx":3644
         PICH            =   "frmtMeterai.frx":3BCC
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
         X1              =   5880
         X2              =   240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000080FF&
         Height          =   1575
         Left            =   240
         Top             =   120
         Width           =   5655
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4380
         Top             =   465
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nilai Meterai"
         Height          =   195
         Left            =   1260
         TabIndex        =   4
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Batas Minimal"
         Height          =   195
         Left            =   1155
         TabIndex        =   3
         Top             =   540
         Width           =   975
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   30
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7575
      Left            =   0
      Picture         =   "frmtMeterai.frx":4208
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmtMeterai"
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
            Case "PPN", "DISC", "FN_BATAS", "FN_NILAI"
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
            Case "PPN", "DISC", "FN_BATAS", "FN_NILAI"
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

Private Sub cmdHapus_Click()
    If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        Call objO.DeleteData("delete T_METERAI where FN_BATAS=" & text(0) & "")
        text(0) = ""
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset

    If text(0) <> "" And text(1) <> "" Then
        Set rsSearch = objO.FetchData("Select * from T_METERAI WITH(NOLOCK) where FN_BATAS=" & text(0) & "")
        
        Call ClearsSql
        gsSql1(0) = "FN_BATAS"
        gsSql2(0) = SimpanKoma2Titik(text(0))
        gsSql1(1) = "FN_NILAI"
        gsSql2(1) = SimpanKoma2Titik(text(1))
        
        If Not rsSearch.BOF Then
            sSQL = GensSql("T_METERAI", True) & " where FN_BATAS=" & text(0) & ""
        Else
            sSQL = GensSql("T_METERAI", False)
        End If
        Call objO.UpdateData(sSQL)
        
        Call hapusForm
        text(0) = ""
        text(0).SetFocus
    Else
        MsgBox "Masukkan Batas Minimal & Nilai Meterainya !"
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
            Case 0
                gMainKeyCari = text(Index)
                gSearchOption = "METERAI"
                frmCari.Show (vbModal)
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
        Case 0
            text(Index) = Round(Val(SimpanKoma2Titik(text(Index))), 2)
            Set rsSearch = objO.FetchData("Select * from T_METERAI WITH(NOLOCK) where FN_BATAS=" & text(0) & "")
            If rsSearch.BOF Then
                Call hapusForm
            Else
                text(1) = rsSearch("FN_NILAI")
            End If
        Case 1
            text(Index) = Round(Val(SimpanKoma2Titik(text(Index))), 2)
    End Select
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 1 To 1
        text(i) = ""
    Next
End Sub
