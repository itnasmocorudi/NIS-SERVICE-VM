VERSION 5.00
Begin VB.Form frmtmAlasan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alasan Perubahan"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmAlasan 
      BackColor       =   &H00FFC0C0&
      Caption         =   "                                       Alasan Perubahan                                        "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdSimpan 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1920
         Width           =   5055
      End
      Begin VB.TextBox txtAlasan 
         Height          =   1695
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmtmAlasan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpan_Click()
    If Trim(txtAlasan) <> "" And Len(Trim(txtAlasan)) >= 10 Then
        frmtmaintenance.Tag = "valid"
        gBantuan4 = Trim(txtAlasan)
        Unload Me
    Else
        MsgBox ("Harus diisi, minimal 10 huruf")
        txtAlasan.SelLength = Len(txtAlasan)
        txtAlasan.SetFocus
    End If
End Sub
