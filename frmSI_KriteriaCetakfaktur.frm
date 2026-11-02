VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmSI_KriteriaCetakfaktur 
   BorderStyle     =   0  'None
   Caption         =   "Laporan Neraca"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   9615
      Begin VB.Frame frame2 
         Height          =   2625
         Left            =   420
         TabIndex        =   13
         Top             =   3360
         Width           =   8625
         Begin VB.OptionButton Prt2170 
            Caption         =   "Epson LQ-2170"
            Height          =   255
            Left            =   4545
            TabIndex        =   23
            Top             =   2295
            Width           =   1665
         End
         Begin VB.OptionButton Prt1170 
            Caption         =   "Epson LQ-1170"
            Height          =   255
            Left            =   2655
            TabIndex        =   22
            Top             =   2295
            Width           =   1665
         End
         Begin VB.TextBox TxtJabatan 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2640
            TabIndex        =   8
            Top             =   1890
            Width           =   5145
         End
         Begin VB.CheckBox ChkDisc 
            Height          =   195
            Left            =   7500
            TabIndex        =   7
            Top             =   1590
            Value           =   1  'Checked
            Width           =   345
         End
         Begin VB.TextBox TxtNamaTTd 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2640
            TabIndex        =   6
            Top             =   1530
            Width           =   3345
         End
         Begin VB.CheckBox ChkNmPerusahaan 
            Height          =   195
            Left            =   2640
            TabIndex        =   2
            Top             =   600
            Value           =   1  'Checked
            Width           =   345
         End
         Begin VB.TextBox txtNamaKota 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2640
            TabIndex        =   5
            Top             =   1170
            Width           =   5175
         End
         Begin MSAdodcLib.Adodc AdoFaktur 
            Height          =   345
            Left            =   7320
            Top             =   630
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select kodegl,namagl from glmst"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CboTypeFaktur 
            Height          =   345
            Left            =   2640
            TabIndex        =   1
            Top             =   240
            Width           =   5205
            DataFieldList   =   "Column 1"
            _Version        =   196616
            DataMode        =   2
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnHeaders   =   0   'False
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   10054
            Columns(0).Caption=   "Ket"
            Columns(0).Name =   "Ket"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   3200
            Columns(1).Visible=   0   'False
            Columns(1).Caption=   "Kode"
            Columns(1).Name =   "Kode"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            _ExtentX        =   9181
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataFieldToDisplay=   "Column 0"
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CboFakturAwal 
            Height          =   345
            Left            =   2640
            TabIndex        =   3
            Top             =   780
            Width           =   2085
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   3678
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   -2147483643
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
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboFakturAkhir 
            Height          =   345
            Left            =   5430
            TabIndex        =   4
            Top             =   810
            Width           =   2085
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   3678
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   -2147483643
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Printer"
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
            Index           =   1
            Left            =   150
            TabIndex        =   24
            Top             =   2310
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type Faktur"
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
            Index           =   2
            Left            =   150
            TabIndex        =   21
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ditanda Tangani Oleh"
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
            Index           =   1
            Left            =   150
            TabIndex        =   20
            Top             =   1590
            Width           =   1875
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Kota"
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
            Index           =   1
            Left            =   150
            TabIndex        =   19
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cetak Nama Perusahaan"
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
            Index           =   0
            Left            =   150
            TabIndex        =   18
            Top             =   600
            Width           =   2115
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cetak Discount"
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
            Index           =   0
            Left            =   6090
            TabIndex        =   17
            Top             =   1590
            Width           =   1320
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nomor Faktur"
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
            Index           =   0
            Left            =   150
            TabIndex        =   16
            Top             =   930
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Bagian"
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
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   1950
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/d"
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
            Index           =   2
            Left            =   4800
            TabIndex        =   14
            Top             =   870
            Width           =   330
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2325
         Left            =   420
         TabIndex        =   12
         Top             =   1050
         Width           =   8595
         Begin MSComctlLib.ListView TxtFormatlaporan 
            Height          =   1545
            Left            =   300
            TabIndex        =   0
            Top             =   600
            Width           =   8085
            _ExtentX        =   14261
            _ExtentY        =   2725
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   8454143
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Kode"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Keterangan Kartu"
               Object.Width           =   11818
            EndProperty
         End
         Begin AIFCmp1.asxPowerBanner BannerKet 
            Height          =   375
            Left            =   30
            Top             =   120
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   661
            StartColor      =   6633234
            EndColor        =   16777215
            FormatString    =   "                                  KRITERIA CETAK FAKTUR DAN KWITANSI"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   16777215
         End
      End
      Begin ucXPButton.XPButton CmdQuit 
         Cancel          =   -1  'True
         Height          =   675
         Left            =   8250
         TabIndex        =   10
         Top             =   180
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1191
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483638
         BCOLO           =   -2147483638
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSI_KriteriaCetakfaktur.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton CmdPrint 
         Height          =   645
         Left            =   510
         TabIndex        =   9
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "&Open"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   -2147483638
         BCOLO           =   -2147483638
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmSI_KriteriaCetakfaktur.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmSI_KriteriaCetakfaktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Job ID          : 014 Include report --> Report_NeracaBi,Report_NeracaBiAndBl and Report_NeracaRasio
'Tgl. Mulai      : 12 April 2007
'Created Coding By Ade Wardana; Improvement By Ade Wardana
'Tgl. Modifikasi : -

'=======================================

Dim TblTemp     As New ADODB.Recordset

Private Sub CboFakturAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode, CboFakturAkhir
End Sub

Private Sub CboFakturAkhir_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub CboFakturAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode, CboFakturAwal
End Sub

Private Sub CboFakturAwal_KeyPress(KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub CboTypeFaktur_Click()
        CboFakturAwal.Text = ""
        CboFakturAkhir.Text = ""
If CboTypeFaktur.Columns(1).Text = "A" Then
   SetAdodc ("A")
ElseIf CboTypeFaktur.Columns(1).Text = "B" Then
   SetAdodc ("B")
Else
   SetAdodc ("")
End If
End Sub

Private Sub CboTypeFaktur_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode, CboTypeFaktur
End Sub

Private Sub ChkDisc_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub ChkNmPerusahaan_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub CmdPrint_Click()
 Dim Prt As String
 If Trim(CboFakturAkhir) = "" Then CboFakturAkhir = String(10, "z")
 Select Case TxtFormatlaporan.ListItems(TxtFormatlaporan.SelectedItem.Index)
    Case "K"
          FrmLap_SiCetakNotaDebet.Show
    Case "M"
          FrmLap_SiCetakfakturJual.Show
           If Prt1170.Value Then
               Prt = "1170"
           Else
               Prt = "2170"
           End If
           If user.TypePrinter <> Prt Then
              Set RsUser = DbConP.Execute("Update [user] Set TypePrinter='" & Prt & "'" & _
                                "WHERE user_id ='" & user.UserId & "' ")
                user.TypePrinter = Prt
           End If
 End Select

End Sub

Private Sub Form_Load()
Call Create_Banner
With FrmMain
     Set CmdQuit.PictureNormal = .ImageList1.ListImages("cmdquit").Picture
     CmdQuit.Tag = "keluar"
     CmdQuit.Caption = "&Exit"
     Set CmdPrint.PictureNormal = .ImageList1.ListImages("cmdprint").Picture
     CmdPrint.Tag = "Cetak Ke Bentuk Laporan"
     CmdPrint.Caption = "&Print"
End With
SetColorProfile Me
AdoFaktur.ConnectionString = AdoCon
    
If Left(user.TypePrinter, 1) = "1" Then
   Prt1170.Value = True
Else
   Prt2170.Value = True
End If
  With CboTypeFaktur
      .AddItem "Penjualan" & Chr(9) & "A"
      .AddItem "Retur Penjualan" & Chr(9) & "B"
      .Value = "A"
  End With
    
  Call SetAdodc("")
  Call FillListView
  
'If frmSI_MasukKeBengkel.Tag = "" Then
'  TxtFormatlaporan.ListItems.Item(2).Selected = True
'  CmdPrint_Click
'  Unload frmSI_KriteriaCetakfaktur
'End If
  
  txtNamaKota = Profile.KotaProfil
   
End Sub
Private Sub SetAdodc(Tipe)
        If Tipe = "A" Then
            Sql = "SELECT DISTINCT NoFaktur, TglFaktur FROM Sihdr where typetrn='A' or typetrn='B' ORDER BY NoFaktur"
        ElseIf Tipe = "B" Then
            Sql = "SELECT DISTINCT NoFaktur, TglFaktur FROM Sihdr where typetrn='K' or typetrn='L' ORDER BY NoFaktur"
        Else
            Sql = "SELECT DISTINCT NoFaktur, TglFaktur FROM Sihdr ORDER BY NoFaktur"
        End If
        
        AdoFaktur.RecordSource = Sql
        AdoFaktur.Refresh
        With CboFakturAwal
            .DataSourceList = AdoFaktur
            .DataFieldList = "NoFaktur"
            .Columns(0).Width = 3000
            .Columns(1).Width = 1300
            .Columns(0).Caption = "No. Faktur"
            .Columns(1).Caption = "Tgl. Faktur"
        End With

        With CboFakturAkhir
            .DataSourceList = Nothing
            .DataSourceList = AdoFaktur
            .DataFieldList = "NoFaktur"
            .Columns(0).Width = 3000
            .Columns(1).Width = 1300
            .Columns(0).Caption = "No. Faktur"
            .Columns(1).Caption = "Tgl. Faktur"
        End With
   
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Objek_MouseMove Frame1, X, Y
End Sub
Private Sub CmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Objek_MouseMove CmdPrint, X, Y
End Sub
Private Sub CmdQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Objek_MouseMove CmdQuit, X, Y
End Sub
Private Sub CmdQuit_Click()
    Unload Me
End Sub

Sub Create_Banner()
     Set Bannerkt = Controls.Add("AIFCmp1.asxPowerBanner", "BannerKt")
     With Bannerkt
        Set .Container = Me
        .Height = Me.Height
        .ZOrder 0
         Kaption = Trim(FrmMain.ListKiri.ListItems(FrmMain.ListKiri.SelectedItem.Index).SubItems(1)) & _
                " [" & Trim(Replace(FrmMain.TabStrip1.SelectedItem.Key, "&", "And")) & "]"
        .FormatString = Kaption
        .Top = 0
        .Left = -70
        
        .Visible = True
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
   HapusTab (Me.Name)
End Sub

Private Sub Form_Resize()
    Frame1.Top = 0
    Frame1.Left = (Me.Width - Frame1.Width) / 2
End Sub

Private Sub TxtFormatlaporan_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub TxtJabatan_KeyPress(KeyAscii As Integer)
   KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub txtNamaKota_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub TxtNamaTTd_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub TxtJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub TxtFormatlaporan_Click()
1       If TxtFormatlaporan.ListItems.Count < 1 Then Exit Sub
2       For a = 1 To TxtFormatlaporan.ListItems.Count
3         TxtFormatlaporan.ListItems(a).Checked = False
4       Next
5       TxtFormatlaporan.ListItems(TxtFormatlaporan.SelectedItem.Index).Checked = True
6       If TxtFormatlaporan.SelectedItem.Index = 2 Then
         Label9(1).Visible = True
         Prt1170.Visible = True
         Prt2170.Visible = True
        Else
         Label9(1).Visible = False
         Prt1170.Visible = False
         Prt2170.Visible = False
        End If
7       On Error GoTo 0
8       Exit Sub

Vb_Error:

9         MsgBox ErrMessage(Erl, err.Number, "Procedure : frmSI_KriteriaCetakfaktur.TxtFormatlaporan_Click"), vbExclamation, "Err Number : " & Erl
End Sub

Private Sub TxtFormatlaporan_DblClick()
    CmdPrint_Click
End Sub

Private Sub TxtFormatlaporan_ItemClick(ByVal Item As MSComctlLib.ListItem)
  TxtFormatlaporan_Click
End Sub


Private Sub FillListView()
    Dim VListItem      As MSComctlLib.ListItem
    
    TxtFormatlaporan.ListItems.Clear
'No;Keterangan Laporan;K;;M;Cetak Faktur Parts
    Set VListItem = TxtFormatlaporan.ListItems.Add(, , "K")
    VListItem.ListSubItems.Add , , "Cetak Nota Debet"
    'VListItem.ListSubItems.Add , , "Kecil"
             
    Set VListItem = TxtFormatlaporan.ListItems.Add(, , "M")
    VListItem.ListSubItems.Add , , "Cetak Faktur Parts"
    'VListItem.ListSubItems.Add , , "Kecil"
             
    TxtFormatlaporan.ListItems(1).Checked = True
End Sub

Private Sub TxtNama_KeyDown(KeyCode As Integer, Shift As Integer)
    Enter KeyCode
End Sub

Private Sub TxtNamattd_KeyPress(KeyAscii As Integer)
KeyAscii = UpCase(KeyAscii)
End Sub
