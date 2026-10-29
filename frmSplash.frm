VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4665
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4665
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   360
      Top             =   2880
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   6000
      TabIndex        =   0
      Top             =   720
      Width           =   1665
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    Timer1.Enabled = False
    frmPassword.Show
End Sub

Private Sub Form_Load()
    lblVersion.Caption = gVersi
    Call CenterForm(Me, True)
End Sub

Private Sub Frame1_Click()
    Unload Me
    Timer1.Enabled = False
    frmPassword.Show
End Sub

Private Sub Timer1_Timer()
    Unload Me
    Timer1.Enabled = False
    If allowIn Then
        frmPassword.Show
    Else
        MsgBox "Lengkapi Profil. "
        frmSetup.Show
        Exit Sub
    End If
End Sub
