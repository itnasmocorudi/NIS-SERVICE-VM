VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmTukarPartOil 
   Caption         =   "Tukar Part dengan Oli"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblSimpan 
      Caption         =   "&Tukar"
      Height          =   390
      Left            =   7935
      TabIndex        =   1
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   2
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   525
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCC"
      PromptChar      =   ">"
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Tukar Part dengan Oli"
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Index           =   0
      Left            =   3180
      TabIndex        =   4
      Top             =   540
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   7
      Top             =   885
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   8
      Top             =   1215
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      PromptChar      =   ">"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Oli"
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   1275
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Part"
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   930
      Width           =   750
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Invoice"
      Height          =   255
      Left            =   660
      TabIndex        =   3
      Top             =   585
      Width           =   1365
   End
End
Attribute VB_Name = "frmTukarPartOil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HapusForm()
    Text(0) = ""
End Sub

Private Sub Form_Load()
    Call HapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblSimpan_Click()
    If Text(0) = "" Then
        MsgBox "Masukkan No. Invoice!", vbOKOnly, "Perhatian"
    Else
        If MsgBox("Yakin Menukar Part dengan Oli No Invoice :" & Text(0) & " !", vbYesNo, "konfirmasi") = vbYes Then
            
            sSQL = "UPDATE T_INVOICE set PART =" & SimpanKoma2Titik(Text(2)) _
                 & " ,OIL =" & SimpanKoma2Titik(Text(1)) & "" _
                 & " where NOINV='" & Trim(Text(0)) & "'"
            oConn.Execute sSQL
            
            sSQL = "UPDATE JOBDTA set PART =" & SimpanKoma2Titik(Text(2)) _
                 & " ,OIL =" & SimpanKoma2Titik(Text(1)) & "" _
                 & " where NOINV='" & Trim(Text(0)) & "'"
            oConn.Execute sSQL
            
            MsgBox "Proses Selesai !", vbOKOnly, ""
        End If
    End If
    Text(0).SetFocus
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
                gMainKeyCari = Text(Index)
                gSearchOption = "PARTOLI"
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
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                oRs.Open "select * from T_INVOICE where NOINV='" & Trim(Text(0)) & "'", _
                         oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "PARTOLI"
                    frmCari.Show (vbModal)
                Else
                    Text(1) = Format(oRs("PART"), "#,##0.00")
                    Text(2) = Format(oRs("OIL"), "#,##0.00")
                End If
                oRs.Close
            End If
    End Select
End Sub
