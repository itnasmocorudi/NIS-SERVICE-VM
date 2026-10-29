VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmBayarBilling 
   Caption         =   "Pembayaran Invoice"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   12240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TblSimpan 
      Caption         =   "&Bayar"
      Height          =   390
      Left            =   7935
      TabIndex        =   5
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   6
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   5370
      TabIndex        =   13
      Top             =   120
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Height          =   2760
      Left            =   195
      TabIndex        =   10
      Top             =   435
      Width           =   11505
      Begin VB.Frame Frame4 
         Height          =   2505
         Left            =   6420
         TabIndex        =   19
         Top             =   135
         Width           =   4950
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   2
            Left            =   2085
            TabIndex        =   2
            Top             =   345
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   503
            _Version        =   393216
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
            Index           =   3
            Left            =   2085
            TabIndex        =   3
            Top             =   645
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "C"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   2085
            TabIndex        =   4
            Top             =   945
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   2085
            TabIndex        =   28
            Top             =   1245
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin VB.Label Label30 
            Caption         =   "(DD-MM-YYYY)"
            Height          =   240
            Left            =   3360
            TabIndex        =   33
            Top             =   405
            Width           =   1140
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Pembayaran"
            Height          =   255
            Left            =   420
            TabIndex        =   32
            Top             =   675
            Width           =   1530
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Pembayaran"
            Height          =   255
            Left            =   540
            TabIndex        =   31
            Top             =   390
            Width           =   1410
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pembayaran"
            Height          =   255
            Left            =   810
            TabIndex        =   30
            Top             =   990
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Selisih"
            Height          =   255
            Left            =   1005
            TabIndex        =   29
            Top             =   1290
            Width           =   945
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   1905
         TabIndex        =   7
         Top             =   225
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1905
         TabIndex        =   8
         Top             =   525
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   1905
         TabIndex        =   15
         Top             =   825
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   1905
         TabIndex        =   16
         Top             =   1125
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   1905
         TabIndex        =   20
         Top             =   1425
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   1905
         TabIndex        =   21
         Top             =   1725
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   1905
         TabIndex        =   22
         Top             =   2325
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   1905
         TabIndex        =   23
         Top             =   2025
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         Height          =   255
         Left            =   765
         TabIndex        =   27
         Top             =   2355
         Width           =   945
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pajak / PPN"
         Height          =   255
         Left            =   825
         TabIndex        =   26
         Top             =   1770
         Width           =   945
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sebelum Pajak"
         Height          =   255
         Left            =   630
         TabIndex        =   25
         Top             =   1470
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Meterai"
         Height          =   255
         Left            =   810
         TabIndex        =   24
         Top             =   2070
         Width           =   945
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Invoice"
         Height          =   255
         Left            =   825
         TabIndex        =   18
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
         Height          =   255
         Left            =   825
         TabIndex        =   17
         Top             =   1155
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   810
         TabIndex        =   12
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   810
         TabIndex        =   11
         Top             =   270
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   4350
      TabIndex        =   1
      Top             =   105
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
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
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   150
      Width           =   945
   End
   Begin VB.Label No 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Inv"
      Height          =   255
      Left            =   3255
      TabIndex        =   9
      Top             =   150
      Width           =   945
   End
End
Attribute VB_Name = "frmBayarBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
            Case "PPN", "DISC"
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
            Case "PPN", "DISC"
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

Private Sub hapusForm()
    For i = 1 To 13
        Select Case i
            Case 0, 2, 8
                Text(i) = ">>->>->>>>"
            Case Else
                Text(i) = ""
        End Select
    Next
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub SimpanData()
    Call hapusForm
    Text(0) = ""
End Sub

Private Sub TblSimpan_Click()
    Call SimpanData
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub PanggilDataOrder(NomorWO As String)

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "INVBAYAR"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Dim NoInv As String
    
    Select Case Index
        Case 1
            If Trim(Text(Index)) <> "" Then
                TblSimpan.Enabled = True
                NoInv = Text(1)
                sSQL = "Select * from T_INVOICE where NOINV ='" & NoInv & "' and INVDATE='" & TglYYYYMMDD(gTglInv) & "' and FS_NO_RETUR = ''"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "INVBAYAR"
                    frmCari.Show (vbModal)
                Else
                    Call PanggilDataInvoice
                End If
                oRs.Close
            End If
        Case 0, 2
            If CekTgl(Text(Index)) = True Then
                Text(Index) = TglDDMMYYYY(Text(Index))
            Else
                Text(Index) = TglDDMMYYYY(Date)
            End If
            If Index = 0 Then
                gTglInv = Text(0)
            End If
        Case 4
            Text(5) = Koma2Titik(Text(13)) - Koma2Titik(Text(4))
    End Select
End Sub

Private Sub PanggilDataInvoice()
    Dim SblmPajak As Double
    Dim Pajak As Double
    
    sSQL = "Select * from T_INVOICE where NOINV ='" & Text(1) & "' and INVDATE='" & TglYYYYMMDD(gTglInv) & "' and FS_NO_RETUR = ''"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        'Data Customer & Order
        Text(6) = oRs1("CUSTOMER")
        Text(7) = oRs1("NO_POLISI")
        Text(8) = TglDDMMYYYY(oRs1("INVDATE"))
        Text(9) = oRs1("NOINV")
        Text(12) = oRs1("MET")
        Text(13) = oRs1("TOTAL")
        SblmPajak = Round(oRs1("TOTAL") * 100 / (100 + (oRs("PPN") * 100)), 0)
        Pajak = oRs1("TOTAL") - SblmPajak
        
        Text(10) = SblmPajak
        Text(11) = Pajak
    End If
    oRs1.Close
End Sub
