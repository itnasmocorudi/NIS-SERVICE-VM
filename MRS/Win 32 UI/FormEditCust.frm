VERSION 5.00
Begin VB.Form FormEditCust 
   BackColor       =   &H00BFE8EA&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ch 
      BackColor       =   &H00BFE8EA&
      Caption         =   "Simpan data lama ke Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2520
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.TextBox txtNoPolisi 
      Appearance      =   0  'Flat
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
      Left            =   9120
      MaxLength       =   25
      TabIndex        =   30
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtPropinsi 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   27
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtKodePos 
      Appearance      =   0  'Flat
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
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtKota 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   10
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtAlamat 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   200
      TabIndex        =   9
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
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
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   8
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtGSM 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9120
      MaxLength       =   25
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtRumah 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   9120
      MaxLength       =   40
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtCDMA 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9120
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Picture         =   "FormEditCust.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   735
      Left            =   9120
      Picture         =   "FormEditCust.frx":0790
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8400
      Picture         =   "FormEditCust.frx":0DEF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      Picture         =   "FormEditCust.frx":14BB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: CUSTOMER ::."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3360
      TabIndex        =   33
      Top             =   520
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   7800
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Pol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8160
      TabIndex        =   32
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9000
      TabIndex        =   31
      Top             =   2040
      Width           =   75
   End
   Begin VB.Label Propinsi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Propinsi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label Propins 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   28
      Top             =   1920
      Width           =   75
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5160
      TabIndex        =   26
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6120
      TabIndex        =   25
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   23
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama  Cust"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat Cust"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   20
      Top             =   840
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   19
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9000
      TabIndex        =   18
      Top             =   1200
      Width           =   75
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9000
      TabIndex        =   17
      Top             =   840
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GSM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   16
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PSTN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   15
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9000
      TabIndex        =   14
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CDMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   13
      Top             =   1560
      Width           =   600
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000FF&
      X1              =   8160
      X2              =   10920
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: TELPON ::."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8160
      TabIndex        =   12
      Top             =   520
      Width           =   2775
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   11040
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   8160
      X2              =   10920
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   3960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1305
      X2              =   3960
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   3960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATA SURAT MENYURAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   705
      TabIndex        =   4
      Top             =   120
      Width           =   10335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   11280
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   11160
      X2              =   0
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11035
   End
End
Attribute VB_Name = "FormEditCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit


Private xMe As Boolean

Private noRangka As String
Private objCu As New ClsDataPelanggan
Private rsCu As ADODB.Recordset
Private NoSPK As String

Private Nama As String
Private Alamat As String
Private KodePos As String
Private Kota As String
Private Prop As String
Private Rumah As String
Private GSM As String
Private CDMA As String
Private NoPolisi As String

Public Property Let isMe(value As Boolean)
   xMe = value
End Property

Public Property Let NomorRangka(value As String)
   noRangka = value
End Property

Private Sub Form_Load()
    Call OnLoad
    Call cmdEdit_Click
End Sub

Private Sub OnLoad()
    If xMe = True Then
        Call ListDgCustomer
        Call LoadControl
    Else
        Call ListDgCustomerBengkel
        Call LoadControl
        txtPropinsi.Visible = False
    End If
End Sub

Private Sub ListDgCustomer()
   Set rsCu = objCu.FetchItemCustomer(noRangka)
   If rsCu.RecordCount = 0 Then
      Unload Me
   End If
   Call BindToControlCust(rsCu)
End Sub

Private Sub ListDgCustomerBengkel()
   Set rsCu = objCu.FetchItemCustomerBengkel(noRangka)
   If rsCu.RecordCount = 0 Then
      Unload Me
   End If
   Call BindToControlCust(rsCu)
End Sub

Private Sub BindToControlCust(rs As ADODB.Recordset)
On Error Resume Next
   rs.MoveFirst
   With Me
      .txtNama.Text = Nul(rs.Fields(1).value)
      .txtAlamat.Text = Nul(rs.Fields(2).value)
      .txtKodePos.Text = Nul(rs.Fields(3).value)
      .txtKota.Text = Nul(rs.Fields(4).value)
      .txtPropinsi.Text = Nul(rs.Fields(5).value)
      .txtRumah.Text = Nul(rs.Fields(6).value)
      .txtGSM.Text = Nul(rs.Fields(7).value)
      .txtCDMA.Text = Nul(rs.Fields(8).value)
      .txtNoPolisi.Text = Nul(rs.Fields(9).value)
      NoSPK = Nul(rs.Fields(10).value)

      Nama = Nul(rs.Fields(1).value)
      Alamat = Nul(rs.Fields(2).value)
      KodePos = Nul(rs.Fields(3).value)
      Kota = Nul(rs.Fields(4).value)
      Prop = Nul(rs.Fields(5).value)
      Rumah = Nul(rs.Fields(6).value)
      GSM = Nul(rs.Fields(7).value)
      CDMA = Nul(rs.Fields(8).value)
      NoPolisi = Nul(rs.Fields(9).value)
   End With
End Sub

Private Sub cmdCancel_Click()
    Call OnLoad
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub LoadControl()
With Me
      .txtNama.Enabled = False
      .txtAlamat.Enabled = False
      .txtKodePos.Enabled = False
      .txtKota.Enabled = False
      .txtPropinsi.Enabled = False
      .txtRumah.Enabled = False
      .txtGSM.Enabled = False
      .txtCDMA.Enabled = False
      .txtNoPolisi.Enabled = False
      .cmdEdit.Enabled = True
      .cmdSave.Enabled = False
      .cmdCancel.Enabled = False
      .cmdExit.Enabled = True
End With
If xMe = False Then
    Propins.Visible = False
    Propinsi.Visible = False
End If
End Sub

Private Sub LoadControl_Edit()
With Me
      .txtNama.Enabled = True
      .txtAlamat.Enabled = True
      .txtKodePos.Enabled = True
      .txtKota.Enabled = True
      .txtPropinsi.Enabled = True
      .txtRumah.Enabled = True
      .txtGSM.Enabled = True
      .txtCDMA.Enabled = True
      .txtNoPolisi.Enabled = True
      .cmdEdit.Enabled = False
      .cmdSave.Enabled = True
      .cmdCancel.Enabled = True
      .cmdExit.Enabled = False
End With
End Sub

Private Sub cmdEdit_Click()
   Call LoadControl_Edit
End Sub

Private Sub cmdSave_Click()
          Dim n As Integer
          Dim boleh As Boolean
          
1        On Error GoTo cmdSave_Click_Error

2         boleh = True
3         If Trim(txtRumah) = "" And Trim(txtGSM) = "" And Trim(txtCDMA) = "" Then
4             boleh = False
5             MsgBox "Data No Telpon belum ada !", vbCritical, "Kelengkapan"
6         End If
         
7         If boleh Then
8         n = MsgBox("Simpan Perubahan Data Surat Menyurat ?", vbQuestion + vbYesNoCancel)
9             If n = 6 Then
10              With objCu
11                  .Nama = Trim(txtNama.Text)
12                  .Alamat = Trim(txtAlamat.Text)
13                  .KodePos = Trim(txtKodePos.Text)
14                  .Kota = Trim(txtKota.Text)
15                  .Propinsi = Trim(txtPropinsi.Text)
16                  .NoPolisi = Trim(txtNoPolisi.Text)
17                  .Rumah = Trim(txtRumah.Text)
18                  .GSM = Trim(txtGSM.Text)
19                  .CDMA = Trim(txtCDMA.Text)
20                  If ch.value = 1 Then Call .AddNew_SCust(noRangka, txtNama, txtAlamat, txtKodePos, txtKota, "", txtRumah, txtGSM, txtCDMA, txtNoPolisi)
21                  If xMe = True Then
22                     Call .UpdateItemCustomer(noRangka, 0)
                       '' --------------------------------------------------------------------------------------
                       '' 20180914 Nopolisi valid sudah diambil dari transaksi dengan Birojasa di project BBN
                       '' --------------------------------------------------------------------------------------
23                     'Call .UpdateItemCustomerNoPol(NoSPK, txtNoPolisi.text)
24                  Else
25                     Call .UpdateItemCustomerBengkel(noRangka)
26                  End If
27               End With
28               Call LoadControl
29            ElseIf n = 7 Then
30                Call OnLoad
31            End If
32        End If

33       On Error GoTo 0
34       Exit Sub

cmdSave_Click_Error:

35        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSave_Click of Form FormEditCust")
End Sub

