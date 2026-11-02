VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00BFE8EA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1875
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1107.812
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0E0FF&
      Height          =   345
      Left            =   1290
      TabIndex        =   6
      Top             =   508
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      MaxLength       =   5
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DCFCFB&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1380
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DCFCFB&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1380
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   885
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Name:"
      Height          =   270
      Index           =   2
      Left            =   105
      TabIndex        =   7
      Top             =   545
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   900
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fPassword As String
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If UCase(txtPassword) = fPassword Then
        LoginSucceeded = True
        gSA = txtUserName
        gNama_SA = txtUserName
        Unload Me
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub txtUserName_Validate(Cancel As Boolean)
Dim rsSearch As ADODB.Recordset
    
    If Trim(txtUserName) <> "" Then
         
        Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WHERE FS_KD_PETUGAS = '" & Trim(txtUserName) & "'")
        If Not rsSearch.EOF Then
            txtName = Trim(rsSearch("FS_NM_PETUGAS"))
            fPassword = RTrim(Mid(rsSearch("fs_kd_pass"), 3, 10))
            gTipeUser = Trim(rsSearch("TIPE"))
        End If
         
    End If
End Sub
