VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmProfilAAS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Profil Connection ke AAS"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6660
      Left            =   360
      TabIndex        =   16
      Top             =   0
      Width           =   11340
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   975
         Left            =   600
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2130
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   3360
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   2370
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2370
         Width           =   855
      End
      Begin MSMask.MaskEdBox Text1 
         Height          =   270
         Index           =   0
         Left            =   2130
         TabIndex        =   6
         Top             =   2040
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin VB.CheckBox ChkPartOliMan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Part && Oli dapat diisi manual"
         Height          =   210
         Left            =   3570
         TabIndex        =   11
         Top             =   585
         Visible         =   0   'False
         Width           =   3045
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2130
         TabIndex        =   0
         Top             =   555
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   23
         Left            =   2130
         TabIndex        =   2
         Top             =   5595
         Visible         =   0   'False
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   24
         Left            =   2130
         TabIndex        =   3
         Top             =   5895
         Visible         =   0   'False
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   5295
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   2130
         TabIndex        =   4
         Top             =   1155
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   2130
         TabIndex        =   5
         Top             =   1455
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   4950
         TabIndex        =   12
         Top             =   870
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   4950
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text1 
         Height          =   270
         Index           =   1
         Left            =   2130
         TabIndex        =   7
         Top             =   2790
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   26
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text1 
         Height          =   255
         Index           =   2
         Left            =   2130
         TabIndex        =   8
         Top             =   3060
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   600
         Top             =   1920
         Width           =   5655
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   3420
         TabIndex        =   34
         Top             =   2430
         Width           =   465
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   1530
         TabIndex        =   33
         Top             =   2430
         Width           =   405
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   1245
         TabIndex        =   32
         Top             =   3405
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "User "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   31
         Top             =   3090
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Server"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1470
         TabIndex        =   30
         Top             =   2828
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Cabang"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   29
         Top             =   2078
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   3870
         TabIndex        =   28
         Top             =   900
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(DD-MM-YYYY)"
         Height          =   255
         Left            =   6225
         TabIndex        =   27
         Top             =   900
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(DD-MM-YYYY)"
         Height          =   255
         Left            =   6225
         TabIndex        =   26
         Top             =   1230
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "s/d Tanggal"
         Height          =   255
         Left            =   3870
         TabIndex        =   25
         Top             =   1230
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gudang Part"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1035
         TabIndex        =   24
         Top             =   1493
         Width           =   900
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Faktur"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1065
         TabIndex        =   23
         Top             =   1193
         Width           =   870
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh : C:\AAS20\DATA\"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   6525
         TabIndex        =   22
         Top             =   5955
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Contoh : C:\AAS20\SYSTEM.MDA"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   6525
         TabIndex        =   21
         Top             =   5640
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Generasi Data"
         Height          =   195
         Index           =   1
         Left            =   915
         TabIndex        =   20
         Top             =   5355
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Direktori System AAS"
         Height          =   195
         Left            =   570
         TabIndex        =   19
         Top             =   5670
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Direktori Database AAS"
         Height          =   195
         Left            =   390
         TabIndex        =   18
         Top             =   5970
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiskal Bulan ini"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   885
         TabIndex        =   17
         Top             =   593
         Width           =   1050
      End
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9405
      TabIndex        =   15
      Top             =   6855
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   8160
      TabIndex        =   14
      Top             =   6795
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7335
      Left            =   0
      Picture         =   "frmProfilAAS.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmProfilAAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nFiskalLama As String

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC", "FN_BATAS", "FN_NILAI"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC", "FN_BATAS", "FN_NILAI"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub ChkPartOliMan_Click()
    If ChkPartOliMan.value = 1 Then
        Text(4).BackColor = &HFFFFFF
    Else
        Text(4).BackColor = &HC0E0FF
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As New ADODB.Recordset

    If Combo1(0).Text >= "01" And Combo1(0).Text <= "12" Then
       Else
       MsgBox "Pemasukkan Bulan harus antara '01' dan '12'"
       Combo1(0).SetFocus
    End If
            
    Call ClearsSql
    gsSql1(0) = "FS_GENERASI"
    gsSql2(0) = gGenerasi
    gsSql1(1) = "FS_FISKAL"
    gsSql2(1) = Text(0)
    gsSql1(2) = "FS_SYS_AAS"
    gsSql2(2) = Text(23)
    gsSql1(3) = "FS_DB_AAS"
    gsSql2(3) = Text(24)
    gsSql1(4) = "FS_KD_FAKTUR"
    gsSql2(4) = Text(2)
    gsSql1(5) = "FS_GUDANG"
    gsSql2(5) = Text(3)
    gsSql1(6) = "FN_MAN_PART"
    gsSql2(6) = ChkPartOliMan.value
    gsSql1(7) = "FD_MAN_AW"
    gsSql2(7) = TglYYYYMMDD(Text(4))
    gsSql1(8) = "FD_MAN_AK"
    gsSql2(8) = TglYYYYMMDD(Text(5))
    gsSql1(9) = "FS_NIS_KOMP"
    gsSql2(9) = Text1(1)
    
    If Not newEnkrip Then
        gsSql1(10) = "FS_NIS_USER"
        gsSql2(10) = Text1(2)
        gsSql1(11) = "FS_NIS_PASS"
        gsSql2(11) = Text2
        gsSql1(12) = "FS_NIS_DB"
        gsSql2(12) = "ACC" + Trim(Text1(0)) + "-" + Trim(Combo1(1).Text) + Trim(Combo1(0).Text)
        gsSql1(13) = "FS_NISSP_KOMP"
        gsSql2(13) = Text1(1)
        gsSql1(14) = "FS_NISSP_USER"
        gsSql2(14) = Text1(2)
        gsSql1(15) = "FS_NISSP_PASS"
        gsSql2(15) = Text2
        gsSql1(16) = "FS_NISSP_DB"
        gsSql2(16) = "ACC" + Trim(Text1(0)) + "-" + Trim(Combo1(1).Text) + Trim(Combo1(0).Text)
        gsSql1(17) = "FS_NISSP_DB2"
        gsSql2(17) = "ACC" + Trim(Text1(0)) + "-" + IIf(CDbl(Combo1(0).Text) = 1, Trim(CDbl(Combo1(1).Text) - 1), Trim(Combo1(1).Text)) + Format(IIf(CDbl(Combo1(0).Text) = 1, 12, Trim(CDbl(Combo1(0).Text) - 1)), "00")
    Else
        
        gsSql1(10) = "FS_NIS_DB"
        gsSql2(10) = "ACC" + Trim(Text1(0)) + "-" + Trim(Combo1(1).Text) + Trim(Combo1(0).Text)
        gsSql1(11) = "FS_NISSP_DB"
        gsSql2(11) = "ACC" + Trim(Text1(0)) + "-" + Trim(Combo1(1).Text) + Trim(Combo1(0).Text)
        gsSql1(12) = "FS_NISSP_DB2"
        gsSql2(12) = "ACC" + Trim(Text1(0)) + "-" + IIf(CDbl(Combo1(0).Text) = 1, Trim(CDbl(Combo1(1).Text) - 1), Trim(Combo1(1).Text)) + Format(IIf(CDbl(Combo1(0).Text) = 1, 12, Trim(CDbl(Combo1(0).Text) - 1)), "00")
    End If
    
    
    
   
    Set rsSearch = objO.FetchData("Select * from T_PROFIL WITH (NOLOCK)")
    If Not rsSearch.BOF Then
        sSQL = GensSql("T_PROFIL", True)
    Else
        sSQL = GensSql("T_PROFIL", False)
    End If
    Call objO.UpdateData(sSQL)
     
    
'        sSQL = "INSERT INTO LogProfilAcc (FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
'                "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, FD_MAN_AW, FD_MAN_AK, FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
'                "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
'                "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, TglPkp, FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER,USERENTRY, CLIENTNAME, TGLENTRY) SELECT FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
'                "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, CAST((CONVERT( VARCHAR(10), FD_MAN_AW,101)) AS VARCHAR(10)), CAST((CONVERT( VARCHAR(10), FD_MAN_AK,101)) AS VARCHAR(10)), FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
'                "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
'                "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, CAST((CONVERT( VARCHAR(10), TGLPKP,105)) AS VARCHAR(10)), FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER, '" & gSA & "','" & ComptName & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' FROM " & gDb & "..T_PROFIL"

        '- sintak tambah fn_no_book
        sSQL = "INSERT INTO LogProfilAcc (FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
                "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, FD_MAN_AW, FD_MAN_AK, FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
                "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
                "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, TglPkp, FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER,USERENTRY, CLIENTNAME, TGLENTRY,FN_NO_BOOK) SELECT FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
                "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, CAST((CONVERT( VARCHAR(10), FD_MAN_AW,101)) AS VARCHAR(10)), CAST((CONVERT( VARCHAR(10), FD_MAN_AK,101)) AS VARCHAR(10)), FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
                "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
                "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, CAST((CONVERT( VARCHAR(10), TGLPKP,105)) AS VARCHAR(10)), FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER, '" & gSA & "','" & ComptName & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "',FN_NO_BOOK FROM " & gDb & "..T_PROFIL WITH (NOLOCK)"
                
        Call objO.UpdateData_Profil(sSQL)
        If Not newEnkrip Then
            MsgBox gsSql2(16)
        Else
            MsgBox gsSql2(11)
        End If
    
    'gGenerasi = Trim(Text(1))
    gFiskal = Trim(Text(0))
    gGudang = Trim(Text(3))
    gManPart = ChkPartOliMan.value
    gTglManAw = TglDDMMYYYY(Text(4))
    gTglManAk = TglDDMMYYYY(Text(5))
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Call CekTombol(KeyCode)
End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
If Combo1(0).Text <= "01" And Combo1(0).Text > "12" Then
   MsgBox "Pemasukan Bulan  Salah "
   Combo1(0).SetFocus
   Cancel = True
Else
    If Trim(Combo1(0)) <> "" Then
        Text(0) = Trim(Combo1(0))
    End If
End If

End Sub

Private Sub Form_Load()
Dim txtBulan As String
Dim txtTahun As String
Dim rsSearch As ADODB.Recordset
    If newEnkrip Then
        Frame2.Visible = True
    End If
   Combo1(0).Clear
   Combo1(1).Clear
   i = 1
   For i = 1 To 12
         If i < 10 Then
             txtBulan = "0" + Trim(str(i))
             Else
             txtBulan = Trim(str(i))
          End If
          Combo1(0).AddItem txtBulan
    Next i
    
    i = 7
    For i = 7 To (year(isnow) + 10) - 2000 '20
         txtTahun = str(2000 + i)
         Combo1(1).AddItem txtTahun
    Next i
        
    Set rsSearch = objO.FetchData("Select * from T_PROFIL WITH (NOLOCK)")
    If Not rsSearch.BOF Then
        Text(1) = rsSearch("FS_GENERASI")
        Text(0) = rsSearch("FS_FISKAL")
        nFiskalLama = rsSearch("FS_FISKAL")
        Text(23) = rsSearch("FS_SYS_AAS")
        Text(24) = rsSearch("FS_DB_AAS")
        Text(2) = rsSearch("FS_KD_FAKTUR")
        Text(3) = rsSearch("FS_GUDANG")
        If rsSearch("FN_MAN_PART") Then
            ChkPartOliMan.value = 1
        Else
            ChkPartOliMan.value = 0
        End If
        Text(4) = TglDDMMYYYY(rsSearch("FD_MAN_AW"))
        Text(5) = TglDDMMYYYY(rsSearch("FD_MAN_AK"))
        Text1(0) = Mid(rsSearch("FS_NIS_DB"), 4, 2)
        Combo1(0).Text = IIf(rsSearch("FS_NIS_DB") <> "", Right(Trim(rsSearch("FS_NIS_DB")), 2), "")
        Combo1(1).Text = IIf(rsSearch("FS_NIS_DB") <> "", Mid(rsSearch("FS_NIS_DB"), 7, 4), "")
        Text1(1) = gServerDb
        Text1(2) = gUser
        Text2 = gPwd
        If Not IsNull(rsSearch("FS_NIS_PASS").value) Then
           Text2 = Trim(rsSearch("FS_NIS_PASS"))
        End If
    End If
     
    If ChkPartOliMan.value = 1 Then
        Text(4).BackColor = &HFFFFFF
    Else
        Text(4).BackColor = &HC0E0FF
    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Dim Tgl As Date
    
    Select Case Index
        Case 0, 1, 3
            Text(Index) = Format(Text(Index), "00")
            If Trim(Index) = 0 Then
                If Trim(Text(Index)) <> "" And Trim(Text(Index)) <> "00" Then
                    Dim nDbase As String
                    nDbase = "ACC" & Trim(Text1(0)) & "-" & Trim(Combo1(1)) & Trim(Text(Index))
                    'If Not Get_ConnectionAcces(nDbase, Trim(Text1(2)), Trim(Text2), Trim(Text1(1))) Then
                    If Not Get_ConnectionAcces(nDbase, Log.LogAccUser, Log.LogAccPass, Log.LogAccServ) Then
                        MsgBox "Data ACC" & Trim(Text1(0)) & "-" & Trim(Combo1(1)) & Trim(Text(Index)) & " belum ada !"
                        Text(0) = nFiskalLama
                    Else
                        Combo1(0) = Text(0)
                    End If
                End If
            End If
        Case 4, 5
            If CekTgl(Text(Index)) <> True Then
                Text(4) = gTanggal
            End If
            Tgl = Text(4)
            Tgl = GetTglAkhir(Month(Tgl), year(Tgl), 0)
            Text(5) = TglDDMMYYYY(Tgl)
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Call CekTombol(KeyCode)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If

End Sub

