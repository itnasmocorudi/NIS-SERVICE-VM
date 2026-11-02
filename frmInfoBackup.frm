VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmInfoBackup 
   Caption         =   "Backup Tabel"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4455
      Top             =   3495
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   582
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
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4455
      Top             =   3135
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   582
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
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4455
      Top             =   2775
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   582
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
      RecordSource    =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5250
      Left            =   8385
      TabIndex        =   13
      Top             =   960
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   9260
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "TAHUN"
         Caption         =   "Tahun"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton TblCancel 
      Caption         =   "Ba&tal"
      Height          =   390
      Left            =   7950
      TabIndex        =   11
      Top             =   7050
      Width           =   1890
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5265
      Top             =   4920
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   900
      TabIndex        =   2
      Top             =   1485
      Width           =   3105
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   1125
      Width           =   3105
   End
   Begin VB.CommandButton TblBackup 
      Caption         =   "&Backup"
      Height          =   390
      Left            =   7950
      TabIndex        =   4
      Top             =   7050
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   5
      Top             =   7050
      Width           =   1890
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
      FormatString    =   "Backup Tabel"
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
      Index           =   2
      Left            =   1635
      TabIndex        =   0
      Top             =   420
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
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
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   0
      Left            =   2580
      TabIndex        =   3
      Top             =   3930
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   476
      _Version        =   393216
      BackColor       =   16777215
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
      Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
      PromptChar      =   ">"
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   5250
      Left            =   6765
      TabIndex        =   15
      Top             =   975
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   9260
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "TAHUN"
         Caption         =   "Tahun"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   5250
      Left            =   10005
      TabIndex        =   17
      Top             =   960
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   9260
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "TAHUN"
         Caption         =   "Tahun"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "3. Jika melakukan penghapusan jangan membatalkan Proses"
      Height          =   195
      Left            =   1065
      TabIndex        =   20
      Top             =   6090
      Width           =   4350
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Direktori file backup :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      TabIndex        =   19
      Top             =   855
      Width           =   3180
   End
   Begin VB.Label Label9 
      Caption         =   "Tabel Invoice"
      Height          =   240
      Left            =   10035
      TabIndex        =   18
      Top             =   690
      Width           =   1875
   End
   Begin VB.Label Label8 
      Caption         =   "Tabel Booking"
      Height          =   240
      Left            =   6765
      TabIndex        =   16
      Top             =   690
      Width           =   1530
   End
   Begin VB.Label Label7 
      Caption         =   "Tabel Order"
      Height          =   240
      Left            =   8400
      TabIndex        =   14
      Top             =   690
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Proses Backup....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   945
      TabIndex        =   12
      Top             =   4515
      Width           =   2190
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "2. Jika data setelah dibackup akan dihapus pilih Yes."
      Height          =   195
      Left            =   1065
      TabIndex        =   10
      Top             =   5820
      Width           =   3765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "1. Masukkan Bulan dan Tahun data yang akan dibackup."
      Height          =   195
      Left            =   1065
      TabIndex        =   9
      Top             =   5535
      Width           =   4110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Perhatian !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   945
      TabIndex        =   8
      Top             =   5205
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Direktori file BCP.EXE"
      Height          =   195
      Left            =   885
      TabIndex        =   7
      Top             =   3975
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Periode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   585
      TabIndex        =   6
      Top             =   465
      Width           =   975
   End
End
Attribute VB_Name = "frmInfoBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Max = 14
Dim Tahun As String
Dim vFile As String
Dim vSukses(Max) As String
Dim vBerhasil As Boolean
Dim i As Byte
Dim vWarna As Boolean
Dim Nilai As Byte

Private Sub hapusForm()
    Text(2) = ">>>>"
End Sub

Private Sub Drive1_Change()
    On Error GoTo Salah
    
    Dir1.Path = Drive1.Drive
    
    Exit Sub
    
Salah:
    MsgBox "Disket rusak atau belum dimasukkan !", vbOKOnly, "Drive Error"
    Drive1.Drive = Left(Dir1.Path, 2)
End Sub

Private Sub Form_Load()
    Dim strConn As String
    
    Call hapusForm
    Dir1.Path = CurDir
    Drive1.Drive = Left(CurDir, 2)
    Text(0) = "C:\Program Files\Bengkel\"
    Timer1.Enabled = False
    TblCancel.Visible = False
    Label6.Caption = ""
    Label6.ForeColor = vbRed
    
    strConn = "select distinct year(date) as TAHUN from jobdta order by year(date)"
    Set DataGrid1.DataSource = Adodc1
    Adodc1.ConnectionString = adoCon
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = strConn
    Adodc1.Refresh
    DataGrid1.ReBind
    
    strConn = "select distinct year(date) as TAHUN from book order by year(date)"
    Set DataGrid2.DataSource = Adodc2
    Adodc2.ConnectionString = adoCon
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = strConn
    Adodc2.Refresh
    DataGrid2.ReBind
    
    strConn = "select distinct year(invdate) as TAHUN from t_invoice order by year(invdate)"
    Set DataGrid3.DataSource = Adodc3
    Adodc3.ConnectionString = adoCon
    Adodc3.CommandType = adCmdText
    Adodc3.RecordSource = strConn
    Adodc3.Refresh
    DataGrid3.ReBind
End Sub

Private Sub TblBackup_Click()
    If Trim(Text(2)) = ">>>>" Then
        MsgBox "Masukkan tahun data yang akan dibackup !", vbOKOnly, "Perhatian"
    Else
        If vbYes = MsgBox("Yakin membackup data? ", vbYesNo, "Perhatian") Then
            Nilai = MsgBox("Data yang telah dibackup dikosongkan dari Tabel ?", vbYesNo, "Konfirmasi")
            If Nilai = vbYes Then
                TblCancel.Enabled = False
            Else
                TblCancel.Enabled = True
            End If
            Label6.Caption = "Proses Backup..... "
            TblBackup.Visible = False
            TblCancel.Visible = True
            TblKeluar.Enabled = False
        
            Tahun = Trim(Text(2))
            If Dir1.Path = "C:\" Then
                vFile = Dir1.Path & Tahun
            Else
                vFile = Dir1.Path & "\" & Tahun
            End If
            
            On Error Resume Next
            MkDir vFile
            
            Kill vFile & "\JOBDTA" & Tahun & ".DAT"
            Kill vFile & "\BOOK" & Tahun & ".DAT"
            Kill vFile & "\T_INVOICE" & Tahun & ".DAT"
            Kill vFile & "\T_MENU.DAT"
            Kill vFile & "\T_PETUGAS.DAT"
            Kill vFile & "\T_DEALER.DAT"
            Kill vFile & "\T_AKSES.DAT"
            Kill vFile & "\T_JOBLIST.DAT"
            Kill vFile & "\T_METERAI.DAT"
            Kill vFile & "\T_PROFIL.DAT"
            Kill vFile & "\HISTORY.DAT"
            Kill vFile & "\MEKANIK.DAT"
            Kill vFile & "\MODEL.DAT"
            Kill vFile & "\CUSTDAT.DAT"
        
            On Error GoTo msg
            FileCopy "C:\Program Files\Bengkel\" + "BCP.EXE", vFile + "\BCP.EXE"
        
            Open vFile + "\BckOutBengk.bat" For Output As #1
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..JOBDTA where year(DATE)= '" & Tahun & "'" & Chr(34) & " queryout " & vFile & "\JOBDTA" & Bulan & Tahun & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..BOOK where year(DATE)= '" & Tahun & "'" & Chr(34) & " queryout " & vFile & "\BOOK" & Bulan & Tahun & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_INVOICE where year(INVDATE)= '" & Tahun & "'" & Chr(34) & " queryout " & vFile & "\T_INVOICE" & Bulan & Tahun & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_MENU " & Chr(34) & " queryout " & vFile & "\T_MENU" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_PETUGAS " & Chr(34) & " queryout " & vFile & "\T_PETUGAS" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_DEALER " & Chr(34) & " queryout " & vFile & "\T_DEALER" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_AKSES " & Chr(34) & " queryout " & vFile & "\T_AKSES" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_JOBLIST " & Chr(34) & " queryout " & vFile & "\T_JOBLIST" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_METERAI " & Chr(34) & " queryout " & vFile & "\T_METERAI" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..T_PROFIL " & Chr(34) & " queryout " & vFile & "\T_PROFIL" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..HISTORY " & Chr(34) & " queryout " & vFile & "\HISTORY" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..MEKANIK " & Chr(34) & " queryout " & vFile & "\MEKANIK" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..MODEL " & Chr(34) & " queryout " & vFile & "\MODEL" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Print #1, "bcp " & Chr(34) & "select * from " & Trim(gDb) & "..CUSTDAT " & Chr(34) & " queryout " & vFile & "\CUSTDAT" & ".DAT -S" & Trim(gServerDb) & " -U" & gUser & " -P" & gPwd & " -c"
            Close #1
        
            ChDir (vFile)
            perintah = Shell(vFile + "\BckOutBengk.bat", vbHide)
            
            For i = 1 To Max
                vSukses(i) = ""
            Next i
            
            Timer1.Enabled = True
            
        End If
    End If
    Exit Sub

msg:
    MsgBox "Ketikan nama Direktori dengan file bcp.exe didalamnya !", vbOKOnly, "Kesalahan"
End Sub

Private Sub TblCancel_Click()
    On Error GoTo Err
    If vbYes = MsgBox("Yakin Membatalkan Proses Backup ?", vbYesNo, "Perhatian") Then
        TblBackup.Visible = True
        TblCancel.Visible = False
        Timer1.Enabled = False
        TblKeluar.Enabled = True
        Label6.ForeColor = vbRed
        Label6.Caption = "Proses Batal..... "
        
        Kill vFile & "\BckOutBengk.bat"
    End If
Err:
    MsgBox "File yang dibatalkan tidak ada !"
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub text_GotFocus(Index As Integer)
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

Private Sub text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
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

Private Sub Timer1_Timer()
        
        If vWarna = False Then
            Label6.ForeColor = vbBlack
            vWarna = True
        Else
            Label6.ForeColor = vbRed
            vWarna = False
        End If
        
        vSukses(1) = Dir(vFile & "\JOBDTA" & Tahun & ".DAT")
        vSukses(2) = Dir(vFile & "\BOOK" & Tahun & ".DAT")
        vSukses(3) = Dir(vFile & "\T_INVOICE" & Tahun & ".DAT")
        vSukses(4) = Dir(vFile & "\T_MENU.DAT")
        vSukses(5) = Dir(vFile & "\T_PETUGAS.DAT")
        vSukses(6) = Dir(vFile & "\T_DEALER.DAT")
        vSukses(7) = Dir(vFile & "\T_AKSES.DAT")
        vSukses(8) = Dir(vFile & "\T_JOBLIST.DAT")
        vSukses(9) = Dir(vFile & "\T_METERAI.DAT")
        vSukses(10) = Dir(vFile & "\T_PROFIL.DAT")
        vSukses(11) = Dir(vFile & "\HISTORY.DAT")
        vSukses(12) = Dir(vFile & "\MEKANIK.DAT")
        vSukses(13) = Dir(vFile & "\MODEL.DAT")
        vSukses(14) = Dir(vFile & "\CUSTDAT.DAT")

        For i = 1 To Max
            If vSukses(i) <> "" Then
                vBerhasil = True
            Else
                vBerhasil = False
                Exit For
            End If
        Next
        
        If vBerhasil = True Then
            If Nilai = vbYes Then
                sSQL = "delete " & Trim(gDb) & "..BOOK where year(DATE)= '" & Tahun & "'"
                oConn.Execute sSQL
                sSQL = "delete " & Trim(gDb) & "..JOBDTA where year(DATE)= '" & Tahun & "'"
                oConn.Execute sSQL
                sSQL = "delete " & Trim(gDb) & "..T_INVOICE where year(INVDATE)= '" & Tahun & "'"
                oConn.Execute sSQL
            End If
            TblBackup.Visible = True
            TblCancel.Visible = False
            TblKeluar.Enabled = True
            Timer1.Enabled = False
            Label6.ForeColor = vbRed
            Label6.Caption = "Proses Selesai..... "
        End If
End Sub
