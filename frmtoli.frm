VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtoli 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Kode Oli"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmtoli.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3420
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   7620
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3090
         TabIndex        =   11
         Top             =   1440
         Width           =   2775
         Begin VB.OptionButton opGen 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Genuine"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opGen 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Non Genuine"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   3090
         TabIndex        =   0
         Top             =   750
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">C"
         PromptChar      =   ">"
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3090
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Oli"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   2
            Top             =   120
            Width           =   495
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Part"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   2400
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
         MICON           =   "frmtoli.frx":2477
         PICN            =   "frmtoli.frx":2493
         PICH            =   "frmtoli.frx":29A6
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
         Left            =   5325
         TabIndex        =   14
         Top             =   2400
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
         MICON           =   "frmtoli.frx":2F48
         PICN            =   "frmtoli.frx":2F64
         PICH            =   "frmtoli.frx":3438
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
         Left            =   3960
         TabIndex        =   15
         Top             =   2400
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
         MICON           =   "frmtoli.frx":3983
         PICN            =   "frmtoli.frx":399F
         PICH            =   "frmtoli.frx":3F27
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdInfoPart 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1730
         _ExtentX        =   3043
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   ""
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
         MICON           =   "frmtoli.frx":4563
         PICN            =   "frmtoli.frx":457F
         PICH            =   "frmtoli.frx":4B71
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
         X1              =   6480
         X2              =   120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000080FF&
         Height          =   2175
         Left            =   120
         Top             =   120
         Width           =   6375
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3435
         Top             =   750
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Genuine / Non Genuine"
         Height          =   195
         Left            =   1245
         TabIndex        =   9
         Top             =   1530
         Width           =   1710
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
         Height          =   195
         Left            =   2640
         TabIndex        =   8
         Top             =   1170
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Untuk memisahkan antara part dan oli"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   2685
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Digit Pertama"
         Height          =   195
         Left            =   1590
         TabIndex        =   6
         Top             =   795
         Width           =   1365
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
      Height          =   7335
      Left            =   0
      Picture         =   "frmtoli.frx":5266
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmtoli"
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
    If MsgBox("Ingin menghapus Kode ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        objO.FetchData ("delete T_OLI where FS_KD_OLI='" & Text(0) & "'")
        Text(0) = ""
        Text(0).SetFocus
    End If
End Sub

Private Sub cmdInfoPart_Click()
    frmtoli_Info.Show
    frmtoli_Info.WindowState = vbMaximized
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset

    If Text(0) <> "" Then
        Set rsSearch = objO.FetchData("Select * from T_OLI WITH (NOLOCK) where FS_KD_OLI='" & Text(0) & "'")
        
        Call ClearsSql
        gsSql1(0) = "FS_KD_OLI"
        gsSql2(0) = Text(0)
        gsSql1(1) = "TIPE"
        gsSql2(1) = IIf(opTipe(0).value, "P", "O")
        gsSql1(2) = "GENUINE"
        gsSql2(2) = IIf(opGen(0).value, 1, 0)
        
        If Not rsSearch.BOF Then
            sSQL = GensSql("T_OLI", True) & " where FS_KD_OLI='" & Text(0) & "'"
        Else
            sSQL = GensSql("T_OLI", False)
        End If
        Call objO.UpdateData(sSQL)
        
        Text(0) = ""
        Text(0).SetFocus
    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub opGen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub opTipe_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
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
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "OLI"
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
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Text(Index) & "'")
                If Not rsSearch.EOF Then
                    If rsSearch("TIPE") = "P" Then
                        opTipe(0) = True
                    Else
                        opTipe(1) = True
                    End If
                    If rsSearch("GENUINE") = 1 Then
                        opGen(0) = True
                    Else
                        opGen(1) = True
                    End If
                End If
            End If
    End Select
     
End Sub
