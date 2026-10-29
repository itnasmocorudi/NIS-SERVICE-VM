VERSION 5.00
Begin VB.Form frmStallDisplay 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DISPLAY BOARD"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15270
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmDisplayStall.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   15270
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Keluar"
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmDisplayStall.frx":0342
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmStallDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

