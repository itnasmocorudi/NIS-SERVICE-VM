VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapBulanFlateRateMekanik 
   Caption         =   "Laporan Flate Rate Mekanik"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4155
      TabIndex        =   43
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2265
      TabIndex        =   42
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6045
      TabIndex        =   13
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7935
      TabIndex        =   14
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   15
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   150
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
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Laporan Flate Rate Mekanik"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   5250
      TabIndex        =   1
      Top             =   165
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
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   510
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
      Left            =   5250
      TabIndex        =   3
      Top             =   510
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
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   870
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
      Index           =   5
      Left            =   5250
      TabIndex        =   5
      Top             =   870
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
      Index           =   6
      Left            =   1560
      TabIndex        =   6
      Top             =   1230
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
      Index           =   7
      Left            =   5250
      TabIndex        =   7
      Top             =   1230
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
      Index           =   8
      Left            =   1560
      TabIndex        =   8
      Top             =   1590
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
      Index           =   9
      Left            =   5250
      TabIndex        =   9
      Top             =   1590
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
      Index           =   10
      Left            =   1560
      TabIndex        =   10
      Top             =   1950
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
      Index           =   11
      Left            =   5250
      TabIndex        =   11
      Top             =   1950
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
      Index           =   12
      Left            =   1560
      TabIndex        =   12
      Top             =   2340
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2220
      TabIndex        =   41
      Top             =   2370
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5490
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   615
      TabIndex        =   44
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapBulanFlateRateMekanik.frx":0000
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mekanik"
      Height          =   255
      Left            =   480
      TabIndex        =   40
      Top             =   2400
      Width           =   960
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   2010
      Width           =   960
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2850
      TabIndex        =   38
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   6570
      TabIndex        =   37
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   4170
      TabIndex        =   36
      Top             =   2010
      Width           =   960
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   1650
      Width           =   960
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2850
      TabIndex        =   34
      Top             =   1650
      Width           =   1425
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   6570
      TabIndex        =   33
      Top             =   1650
      Width           =   1425
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   4170
      TabIndex        =   32
      Top             =   1650
      Width           =   960
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   1290
      Width           =   960
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2850
      TabIndex        =   30
      Top             =   1290
      Width           =   1425
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   6570
      TabIndex        =   29
      Top             =   1290
      Width           =   1425
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   4170
      TabIndex        =   28
      Top             =   1290
      Width           =   960
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   930
      Width           =   960
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2850
      TabIndex        =   26
      Top             =   930
      Width           =   1425
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   6570
      TabIndex        =   25
      Top             =   930
      Width           =   1425
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   4170
      TabIndex        =   24
      Top             =   930
      Width           =   960
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   570
      Width           =   960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   22
      Top             =   540
      Width           =   1425
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   6570
      TabIndex        =   21
      Top             =   540
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   4170
      TabIndex        =   20
      Top             =   570
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   4155
      TabIndex        =   19
      Top             =   195
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   6570
      TabIndex        =   18
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   17
      Top             =   210
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   210
      Width           =   960
   End
End
Attribute VB_Name = "frmLapBulanFlateRateMekanik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub Form_Load()
    Call HapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptLapBulanFlateRateMekanik
    Dim Kondisi As String
    Dim Ket1, Ket2, Ket3, Ket4, Ket5 As String
    Dim Boleh As Boolean
    
    Boleh = True
    If Text(2) <> ">>->>->>>>" And Text(3) = ">>->>->>>>" Then
        Boleh = False
    End If
    If Text(4) <> ">>->>->>>>" And Text(5) = ">>->>->>>>" Then
        Boleh = False
    End If
    If Text(6) <> ">>->>->>>>" And Text(7) = ">>->>->>>>" Then
        Boleh = False
    End If
    If Text(8) <> ">>->>->>>>" And Text(9) = ">>->>->>>>" Then
        Boleh = False
    End If
    If Text(10) <> ">>->>->>>>" And Text(11) = ">>->>->>>>" Then
        Boleh = False
    End If
    
    If Not (CekTgl(Text(0)) And CekTgl(Text(1))) Or Boleh = False Then
        MsgBox "Masukkan Tgl Awal dan Akhir !"
    Else
        If Text(2) <> ">>->>->>>>" Or Text(3) <> ">>->>->>>>" Then
            Ket1 = Text(2) & " s/d " & Text(3)
        Else
            Ket1 = ""
        End If
        If Text(4) <> ">>->>->>>>" Or Text(5) <> ">>->>->>>>" Then
            Ket2 = Text(4) & " s/d " & Text(5)
        Else
            Ket2 = ""
        End If
        If Text(6) <> ">>->>->>>>" Or Text(7) <> ">>->>->>>>" Then
            Ket3 = Text(6) & " s/d " & Text(7)
        Else
            Ket3 = ""
        End If
        If Text(8) <> ">>->>->>>>" Or Text(9) <> ">>->>->>>>" Then
            Ket4 = Text(8) & " s/d " & Text(9)
        Else
            Ket4 = ""
        End If
        If Text(10) <> ">>->>->>>>" Or Text(11) <> ">>->>->>>>" Then
            Ket5 = Text(10) & " s/d " & Text(11)
        Else
            Ket5 = ""
        End If
        
        gTglTemp = Text(0) & " s/d " & Text(1) & " " & Ket1
        gTglTemp1 = Ket2 & " " & Ket3
        gTglTemp2 = Ket4 & " " & Ket5
        
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        Kondisi = " where ((invdate >= '" & TglYYYYMMDD(Text(0)) & "' and invdate <= '" & TglYYYYMMDD(Text(1)) & "') "
        If Text(2) <> ">>->>->>>>" Or Text(3) <> ">>->>->>>>" Then
             Kondisi = Kondisi + " or (invdate >= '" & TglYYYYMMDD(Text(2)) & "' and invdate <= '" & TglYYYYMMDD(Text(3)) & "')"
        End If
        If Text(4) <> ">>->>->>>>" Or Text(5) <> ">>->>->>>>" Then
             Kondisi = Kondisi + " or (invdate >= '" & TglYYYYMMDD(Text(4)) & "' and invdate <= '" & TglYYYYMMDD(Text(5)) & "')"
        End If
        If Text(6) <> ">>->>->>>>" Or Text(7) <> ">>->>->>>>" Then
             Kondisi = Kondisi + " or (invdate >= '" & TglYYYYMMDD(Text(6)) & "' and invdate <= '" & TglYYYYMMDD(Text(7)) & "')"
        End If
        If Text(8) <> ">>->>->>>>" Or Text(9) <> ">>->>->>>>" Then
             Kondisi = Kondisi + " or (invdate >= '" & TglYYYYMMDD(Text(8)) & "' and invdate <= '" & TglYYYYMMDD(Text(9)) & "')"
        End If
        If Text(10) <> ">>->>->>>>" Or Text(11) <> ">>->>->>>>" Then
             Kondisi = Kondisi + " or (invdate >= '" & TglYYYYMMDD(Text(10)) & "' and invdate <= '" & TglYYYYMMDD(Text(11)) & "')"
        End If
        
        Kondisi = Kondisi + ") and DELIV <> 'BTL' and noinv <> '' "

        vNamaTabel = "Flate_Mekanik" & DriveSerialNumber("c:")
        sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" _
             & " drop table [dbo].[" & vNamaTabel & "]"
        oConn.Execute sSQL

        sSQL = "CREATE TABLE [dbo].[" & vNamaTabel & "] (" _
             & " [NOINV] [varchar] (8) NULL ," _
             & " [NO_WO] [varchar] (8) NULL ," _
             & " [INVDATE] [datetime] NULL ," _
             & " [NO_POLISI] [varchar] (12) NULL ," _
             & " [MECH] [varchar] (4) NULL ," _
             & " [NAME] [varchar] (15) NULL ," _
             & " [FN_RATE] [decimal](18, 2) NULL ," _
             & " [FN_ACTRATE] [decimal](18, 2) NULL " _
             & " ) ON [PRIMARY]"
        oConn.Execute sSQL
        
        For i = 1 To 8
            sSql1 = "select noinv, no_wo, invdate, no_polisi, sum(rte" & Trim(Str(i)) & ") as RATE," _
                  & " sum(ACTRATE" & Trim(Str(i)) & ") as ACTRATE, MEKANIK" & Trim(Str(i)) & " as MECH from jobdta " _
                  & Kondisi & " and FN_RETURN=0 and MEKANIK" & Trim(Str(i)) & "='" & Trim(Text(12)) & "'" _
                  & " group by noinv, no_wo, invdate, no_polisi, MEKANIK" & Trim(Str(i)) & " order by " _
                  & " noinv, no_wo, invdate, no_polisi, MEKANIK" & Trim(Str(i))
            oRs1.Open sSql1, oConn, adOpenStatic
            If Not oRs1.BOF Then
                Do While Not oRs1.EOF
                    sSQL = "insert into [" & vNamaTabel & "]" _
                         & " (no_polisi, no_wo, invdate, NOINV, MECH, FN_RATE, FN_ACTRATE)" _
                         & " values('" & oRs1("NO_POLISI") & "','" & oRs1("NO_WO") & "','" & TglYYYYMMDD(TglDDMMYYYY(oRs1("INVDATE"))) & "','" & oRs1("NOINV") & "','" & oRs1("MECH") & "'," & SimpanKoma2Titik(oRs1("RATE")) & "," & SimpanKoma2Titik(oRs1("ACTRATE")) & ")"
                    oConn.Execute sSQL
                    oRs1.MoveNext
                Loop
            End If
            oRs1.Close
        Next
        
        sSql1 = "select MECH, NAME from [" & vNamaTabel & "]"
        oRs1.Open sSql1, oConn, adOpenStatic
        If Not oRs1.BOF Then
            Do While Not oRs1.EOF
                sSql2 = "select NAME from mekanik where MECH='" & oRs1("MECH") & "'"
                oRs2.Open sSql2, oConn, adOpenStatic
                If Not oRs2.BOF Then
                    sSQL = "update [" & vNamaTabel & "] set NAME='" & oRs2("NAME") & "' where MECH='" & oRs1("MECH") & "'"
                    oConn.Execute sSQL
                End If
                oRs2.Close
                oRs1.MoveNext
            Loop
        End If
        oRs1.Close
                
        sSQL = "select NO_WO, NO_POLISI, INVDATE, NOINV, MECH, [NAME], sum(FN_RATE) as RATE, sum(FN_ACTRATE) as ACTRTE from [" & vNamaTabel & "]" _
             & " group by INVDATE, NOINV, NO_WO, NO_POLISI, MECH, [NAME] order by INVDATE, NOINV, NO_WO, NO_POLISI, MECH, [NAME]"
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
        rpt.Printer.PaperSize = 39
        rpt.Printer.Orientation = ddOPortrait
       
        Select Case Layar
            Case 1
                rpt.Show
            Case 2
                ARViewer1.ReportSource = rpt
            Case Else
                rpt.PrintReport False
                Unload rpt
        End Select
    End If
End Sub

Private Sub TblCetak_Click()
    Call Cetak(0)
End Sub

Private Sub TblConvert_Click()
    Dim txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    txt.FileName = CommonDialog1.FileName
    txt.Export ARViewer1.Pages
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblLayar_Click()
    Call Cetak(1)
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

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 12
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
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
    Select Case Index
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
            Select Case Index
                Case 0, 1
                    If CekTgl(Text(Index)) <> True Then
                        Text(Index) = gTanggal
                    End If
                Case Else
                    If Text(Index) <> ">>->>->>>>" Then
                        If CekTgl(Text(Index)) <> True Then
                            Text(Index) = gTanggal
                        End If
                    End If
            End Select
            If TglYYYYMMDD(Text(0)) > TglYYYYMMDD(Text(1)) Then
                MsgBox "Tanggal awal lebih besar dari tanggal akhir !"
            End If
            Case 12
                If Trim(Text(Index)) <> "" Then
                    sSQL = "select * from mekanik where MECH='" & Text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.BOF Then
                        Cancel = True
                        gMainKeyCari = Text(Index)
                        gSearchOption = "MEKANIK"
                        frmCari.Show (vbModal)
                    End If
                    oRs.Close
                End If
    End Select
End Sub

