VERSION 5.00
Begin VB.Form FormLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: login ::. "
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormLogin.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Login"
      Height          =   732
      Left            =   4080
      Picture         =   "FormLogin.frx":3CC5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   732
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   2520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   1560
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   1920
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   2280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
   If NullValue = True Then
      MsgBox "User ID / Password Masih Kosong !", vbInformation + vbOKOnly
      Exit Sub
   End If
   
   Dim objLogin As New ClsLogin
   Dim rsLogin As ADODB.Recordset
   Set rsLogin = objLogin.FetchItemPassword(txtUser.Text)
   
   Dim mPass As String
   mPass = Crypt(rsLogin.Fields(0).Value, "antztein", False)
      If Trim(txtPass.Text) <> Crypt(mPass, "tandayu", False) Then
         MsgBox "User ID / Password Tidak Valid !", vbExclamation + vbOKOnly
         Exit Sub
      End If
   Unload Me
   Dim f As New FormMainMRS
   f.Show
End Sub

Private Function NullValue() As Boolean
   If txtUser.Text = "" Or txtPass.Text = "" Then
      NullValue = True
   Else
      NullValue = False
   End If
End Function

Private Sub txtPass_KeyPress(Key As Integer)
      If Key = 13 Then SendKeys "{tab}"
End Sub

Private Sub txtUser_KeyPress(Key As Integer)
      If Key = 13 Then SendKeys "{tab}"
End Sub
