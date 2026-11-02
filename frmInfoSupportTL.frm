VERSION 5.00
Begin VB.Form frmInfoSupportTL 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Support TL"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11415
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInfoSupportTL.frx":0000
         Left            =   1800
         List            =   "frmInfoSupportTL.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kategori Problem :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton TblSimpan 
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
      Height          =   390
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9825
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7065
      Width           =   1890
   End
End
Attribute VB_Name = "frmInfoSupportTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInfoSupportTL
' DateTime  : 15/05/2015 08:47
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------

