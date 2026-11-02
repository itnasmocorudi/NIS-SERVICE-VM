VERSION 5.00
Begin VB.Form frmtandaTangan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tanda Tangan"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4200
   Begin VB.Frame frmPrinter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "         Pilih Printer       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LQ2190"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LQ 2180"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LQ 300"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdLanjut 
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LX"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         BorderStyle     =   4  'Dash-Dot
         FillColor       =   &H00C00000&
         Height          =   1575
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdPrinter 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Printer"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chJabatan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Display Jabatan"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.ComboBox cboTtd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Otoritas yang menandatangani Faktur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3075
   End
End
Attribute VB_Name = "frmtandaTangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdLanjut_Click()
    frmPrinter.Visible = False
End Sub

Private Sub cmdPrinter_Click()
    frmPrinter.Visible = True
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset

    Set rsSearch = objO.FetchData("SELECT * FROM SETTANDA WITH(NOLOCK)")
    If Not rsSearch.EOF Then
        Call objO.UpdateData("UPDATE SETTANDA SET OTORITAS = '" & Trim(cboTtd) & "', DISPLAY = '" & IIf(chJabatan = Unchecked, 0, 1) & "', TPPRINT = '" & IIf(prtrn(0).value, 0, IIf(prtrn(1).value, 1, IIf(prtrn(2).value, 2, 3))) & "'")
    Else
        Call objO.InsertData("INSERT INTO SETTANDA VALUES ('" & Trim(cboTtd) & "','" & IIf(chJabatan = Unchecked, 0, 1) & "','" & IIf(prtrn(0).value, 0, IIf(prtrn(1).value, 1, IIf(prtrn(2).value, 2, 3))) & "')")
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim rsSearch As ADODB.Recordset
    With cboTtd
        .Clear
        .AddItem ("KACAB")
        .AddItem ("ADH")
        .AddItem ("KABENG")
    End With
    
    Set rsSearch = objO.FetchData("SELECT * FROM SETTANDA WITH(NOLOCK)")
    If Not rsSearch.EOF Then
        cboTtd = Trim(rsSearch("OTORITAS")) '"KABENG"
        chJabatan = IIf(rsSearch("DISPLAY") = 1, Checked, Unchecked)
        Select Case CDbl(rsSearch("TPPRINT"))
            Case 0
                prtrn(0).value = True
            Case 1
                prtrn(1).value = True
            Case 2
                prtrn(2).value = True
            Case 3
                prtrn(3).value = True
        End Select
    Else
        cboTtd = "KABENG"
        chJabatan = Checked
    End If
End Sub

