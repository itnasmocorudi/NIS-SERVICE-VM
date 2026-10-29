VERSION 5.00
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form ErrInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: SMS ::."
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6135
   ControlBox      =   0   'False
   Icon            =   "ErrInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "ErrInfo.frx":29C12
   ScaleHeight     =   1860
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPesan3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1695
      MaxLength       =   60
      TabIndex        =   4
      Top             =   3825
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.TextBox txtPesan2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1695
      MaxLength       =   61
      TabIndex        =   3
      Top             =   3495
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.TextBox txtPesan1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1695
      MaxLength       =   61
      TabIndex        =   2
      Top             =   3135
      Visible         =   0   'False
      Width           =   9015
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1095
      Left            =   2040
      TabIndex        =   0
      Top             =   40
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1931
      _Version        =   393217
      TextRTF         =   $"ErrInfo.frx":2B15B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ucXPButton.XPButton cmdOK 
      Height          =   375
      Left            =   3330
      TabIndex        =   1
      Top             =   1320
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "ErrInfo.frx":2B1DF
      PICN            =   "ErrInfo.frx":2B1FB
      PICH            =   "ErrInfo.frx":2B68F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line5 
      BorderStyle     =   4  'Dash-Dot
      Visible         =   0   'False
      X1              =   1695
      X2              =   10695
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line Line4 
      BorderStyle     =   4  'Dash-Dot
      Visible         =   0   'False
      X1              =   1695
      X2              =   10695
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Shape Shape1 
      Height          =   1065
      Left            =   1680
      Top             =   3120
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2040
      X2              =   5880
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "ErrInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private HeadCaption As String
Private InfoCaption As String


Public Property Let HeadCap(value As String)
    HeadCaption = value
End Property

Public Property Let InfoCap(value As String)
    InfoCaption = value
End Property

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = HeadCaption
    Me.Text1.text = InfoCaption
End Sub

Private Sub Text1_GotFocus()
    SendKeys ("{TAB}")
End Sub
