VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmBtlOrder 
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Alasan Pembatalan : "
      Height          =   2085
      Left            =   375
      TabIndex        =   2
      Top             =   375
      Width           =   4080
      Begin VB.CommandButton TblProses 
         Caption         =   "&Proses"
         Height          =   390
         Left            =   1830
         TabIndex        =   1
         Top             =   1620
         Width           =   2145
      End
      Begin VB.TextBox Text1 
         Height          =   1260
         Left            =   120
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   3840
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   4471
      StartColor      =   0
      EndColor        =   12632064
      FormatString    =   "Pembatalan Order"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Left            =   345
      Top             =   45
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxSysButton BtnTutup 
      Height          =   285
      Left            =   4140
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      ButtonStyle     =   0
      ButtonFlags     =   16384
   End
End
Attribute VB_Name = "frmBtlOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterForm(Me, True)
End Sub

Private Sub TblProses_Click()
    If Trim(Text1.Text) = "" Then
        MsgBox "Masukkan Alasan Pembatalannya !", vbOKOnly, "Peringatan"
    Else
        sSQL = "update jobdta set DELIV='BTL', ALASAN='" & Left(Text1.Text, 30) & "' " _
             & " ,FS_LOG_DEL='" & gSA & "', FD_TGL_DEL='" & TglYYYYMMDD(gTanggal) & "' " _
             & " where NO_POLISI ='" & gNoPolis & "' and NO_WO ='" & frmBatalOrder.Text(26) & "'"
        oConn.Execute sSQL
        Unload Me
    End If
End Sub

