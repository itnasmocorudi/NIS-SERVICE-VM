VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmTShift 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Master Shift"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chAktif 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktif"
         Height          =   255
         Left            =   1545
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   405
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   2625
         TabIndex        =   5
         Top             =   720
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   2240
         TabIndex        =   8
         Top             =   720
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   1920
         TabIndex        =   10
         Top             =   1560
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
         MICON           =   "frmTShift.frx":0000
         PICN            =   "frmTShift.frx":001C
         PICH            =   "frmTShift.frx":052F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton tbnAdd 
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   1560
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmTShift.frx":0AD1
         PICN            =   "frmTShift.frx":0AED
         PICH            =   "frmTShift.frx":0F39
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H000080FF&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   3855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   3720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2415
         TabIndex        =   6
         Top             =   645
         Width           =   120
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Left            =   975
         TabIndex        =   3
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Operasional"
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
         Left            =   120
         TabIndex        =   1
         Top             =   765
         Width           =   1185
      End
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
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
      MICON           =   "frmTShift.frx":1401
      PICN            =   "frmTShift.frx":141D
      PICH            =   "frmTShift.frx":18F1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   3960
      X2              =   120
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmTShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub hapusForm()
    Text(0) = ""
    Text(1) = ">>:>>"
    Text(2) = ">>:>>"
    chAktif = Checked
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    
    If (Trim(Text(1)) <> "" Or Trim(Text(1)) <> ":" Or Trim(Text(1)) <> ">>:>>") And (Trim(Text(2)) <> "" Or Trim(Text(2)) <> ":" Or Trim(Text(2)) <> ">>:>>") Then
        If Trim(Text(0)) <> "" Then
            Call objO.UpdateData("Update DealerShift set jammulai = '" & Trim(Text(1)) & "', jamselesai = '" & Trim(Text(2)) & "', berlaku = " & IIf(chAktif = Checked, 1, 0) & " where Shift = '" & Trim(Text(0)) & "'")
        Else
            Call objO.InsertData("Insert into DealerShift values ('" & Trim(Text(1)) & "','" & Trim(Text(2)) & "','" & Format(isnow, "mm/dd/YYYY") & "'," & IIf(chAktif = Checked, 1, 0) & ")")
        End If
    End If
    Call hapusForm
End Sub

Private Sub tbnAdd_Click()
    Call hapusForm
End Sub

Private Sub Text_Change(Index As Integer)
    Dim rsSearch As ADODB.Recordset
    
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                Set rsSearch = objO.FetchData("SELECT DISTINCT SHIFT, JAMMULAI, JAMSELESAI, BERLAKU FROM DEALERSHIFT WITH (NOLOCK) WHERE SHIFT = '" & Trim(Text(Index)) & "' ORDER BY SHIFT")
                If Not rsSearch.EOF Then
                    Text(1) = Trim(rsSearch("JAMMULAI"))
                    Text(2) = Trim(rsSearch("JAMSELESAI"))
                    chAktif = IIf(Trim(rsSearch("BERLAKU")) = 1, Checked, Unchecked)
                End If
            End If
    End Select
End Sub

Private Sub text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then DiTab = True
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
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "DLRSHIFT"
                gBantuan = ""
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)

    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                Dim rsSearch As ADODB.Recordset
                Set rsSearch = objO.FetchData("SELECT DISTINCT SHIFT, JAMMULAI, JAMSELESAI, BERLAKU FROM DEALERSHIFT WITH (NOLOCK) WHERE SHIFT = '" & Trim(Text(Index)) & "' ORDER BY SHIFT")
                If Not rsSearch.EOF Then
                    Text(1) = Trim(rsSearch("JAMMULAI"))
                    Text(2) = Trim(rsSearch("JAMSELESAI"))
                    chAktif = IIf(Trim(rsSearch("BERLAKU")) = 1, Checked, Unchecked)
                End If
            End If
        Case 1
            If Trim(Text(Index)) <> "" Then
                
            End If
    End Select
End Sub
