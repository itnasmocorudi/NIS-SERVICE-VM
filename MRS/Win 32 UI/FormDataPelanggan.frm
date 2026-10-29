VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormDataPelanggan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormDataPelanggan.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCSV 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CSV"
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
      Left            =   8640
      Picture         =   "FormDataPelanggan.frx":EDA3
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   735
   End
   Begin VB.Frame FrSMS 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   600
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton Command2 
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
         Left            =   5160
         Picture         =   "FormDataPelanggan.frx":F17A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtPesan 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton cmdSMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SEND"
         Height          =   735
         Left            =   4440
         Picture         =   "FormDataPelanggan.frx":F90A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1980
         Width           =   735
      End
      Begin VB.CheckBox chCDMA 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CheckBox chGSM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   4200
         X2              =   120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   4200
         X2              =   120
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SMS Personal"
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
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   6135
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00404040&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   6135
      End
      Begin VB.Label Label52 
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
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label55 
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
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Isi Pesan"
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
         TabIndex        =   15
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label53 
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
         Left            =   840
         TabIndex        =   14
         Top             =   2400
         Width           =   75
      End
      Begin VB.Label Label56 
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
         Left            =   840
         TabIndex        =   13
         Top             =   2040
         Width           =   75
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00808080&
         Height          =   855
         Left            =   4320
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SMS"
      Height          =   735
      Left            =   1440
      Picture         =   "FormDataPelanggan.frx":FC8C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&History"
      Height          =   732
      Left            =   720
      Picture         =   "FormDataPelanggan.frx":1000E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   732
   End
   Begin VB.CommandButton CmdXls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X&LS"
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
      Left            =   9360
      Picture         =   "FormDataPelanggan.frx":1067B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   735
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
      Left            =   11160
      Picture         =   "FormDataPelanggan.frx":10AF7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton CmdRefreshLap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   10080
      Picture         =   "FormDataPelanggan.frx":11287
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   735
   End
   Begin TrueOleDBGrid70.TDBGrid GridPelanggan 
      Height          =   6615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11668
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tahun Mobil"
      Columns(0).DataField=   "Thn_Mobil"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Customer"
      Columns(1).DataField=   "Nama_Customer"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Propinsi"
      Columns(2).DataField=   "propinsi"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Kota"
      Columns(3).DataField=   "Kota"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Kecamatan"
      Columns(4).DataField=   "Kecamatan"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Kelurahan"
      Columns(5).DataField=   "Kelurahan"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Alamat Customer"
      Columns(6).DataField=   "Alamat_Customer"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Kota"
      Columns(7).DataField=   "Kota"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Kode Pos"
      Columns(8).DataField=   "KodePos"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Telepon"
      Columns(9).DataField=   "Telpon"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "GSM"
      Columns(10).DataField=   "GSM"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "CDMA"
      Columns(11).DataField=   "CDMA"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "No_Polisi"
      Columns(12).DataField=   "No_Polisi"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "No Engine"
      Columns(13).DataField=   "No_Engine"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "No Chassis"
      Columns(14).DataField=   "No_Chassis"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "MDL"
      Columns(15).DataField=   "MDL"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Model"
      Columns(16).DataField=   "Model"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "FMDL"
      Columns(17).DataField=   "FMDL"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "Warna"
      Columns(18).DataField=   "color"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "Religi"
      Columns(19).DataField=   "AGAMA"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "Kode Dealer"
      Columns(20).DataField=   "Kode_Dealer"
      Columns(20).NumberFormat=   "###,###,###,##0"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "Nama Dealer"
      Columns(21).DataField=   "Nama_Dealer"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "KM Akhir"
      Columns(22).DataField=   "Km_Akhir"
      Columns(22).NumberFormat=   "General Number"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Job Desc"
      Columns(23).DataField=   "Job_Desc"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Kode Langgan"
      Columns(24).DataField=   "Kode_Langgan"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "History"
      Columns(25).DataField=   "JMLHISTORY"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Service Terakhir"
      Columns(26).DataField=   "LASTSVC"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "KM TERAKHIR"
      Columns(27).DataField=   "KMAKHIR"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "Tgl Last Service"
      Columns(28).DataField=   "TGLLASTSVC"
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "SBE Last Service"
      Columns(29).DataField=   "LASTSBE"
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   30
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=30"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2858"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2752"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0).ButtonHeader=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5292"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5186"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1).ButtonHeader=1"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=7064"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=6959"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(6).ButtonHeader=1"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=2646"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2540"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(7).ButtonHeader=1"
      Splits(0)._ColumnProps(38)=   "Column(8).Width=2223"
      Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=2117"
      Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(42)=   "Column(8).ButtonHeader=1"
      Splits(0)._ColumnProps(43)=   "Column(9).Width=3519"
      Splits(0)._ColumnProps(44)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(9)._WidthInPix=3413"
      Splits(0)._ColumnProps(46)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(47)=   "Column(9).ButtonHeader=1"
      Splits(0)._ColumnProps(48)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(52)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(53)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(55)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(56)=   "Column(12).Width=2646"
      Splits(0)._ColumnProps(57)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(12)._WidthInPix=2540"
      Splits(0)._ColumnProps(59)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(60)=   "Column(12).ButtonHeader=1"
      Splits(0)._ColumnProps(61)=   "Column(13).Width=3519"
      Splits(0)._ColumnProps(62)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(13)._WidthInPix=3413"
      Splits(0)._ColumnProps(64)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(65)=   "Column(13).ButtonHeader=1"
      Splits(0)._ColumnProps(66)=   "Column(14).Width=4419"
      Splits(0)._ColumnProps(67)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(14)._WidthInPix=4313"
      Splits(0)._ColumnProps(69)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(70)=   "Column(14).ButtonHeader=1"
      Splits(0)._ColumnProps(71)=   "Column(15).Width=4419"
      Splits(0)._ColumnProps(72)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(15)._WidthInPix=4313"
      Splits(0)._ColumnProps(74)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(75)=   "Column(15).ButtonHeader=1"
      Splits(0)._ColumnProps(76)=   "Column(16).Width=1773"
      Splits(0)._ColumnProps(77)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(16)._WidthInPix=1667"
      Splits(0)._ColumnProps(79)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(80)=   "Column(16).ButtonHeader=1"
      Splits(0)._ColumnProps(81)=   "Column(17).Width=3889"
      Splits(0)._ColumnProps(82)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(17)._WidthInPix=3784"
      Splits(0)._ColumnProps(84)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(85)=   "Column(17).ButtonHeader=1"
      Splits(0)._ColumnProps(86)=   "Column(18).Width=3519"
      Splits(0)._ColumnProps(87)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(88)=   "Column(18)._WidthInPix=3413"
      Splits(0)._ColumnProps(89)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(90)=   "Column(18).ButtonHeader=1"
      Splits(0)._ColumnProps(91)=   "Column(19).Width=3281"
      Splits(0)._ColumnProps(92)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(93)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(94)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(95)=   "Column(20).Width=2646"
      Splits(0)._ColumnProps(96)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(97)=   "Column(20)._WidthInPix=2540"
      Splits(0)._ColumnProps(98)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(99)=   "Column(20).ButtonHeader=1"
      Splits(0)._ColumnProps(100)=   "Column(21).Width=3889"
      Splits(0)._ColumnProps(101)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(102)=   "Column(21)._WidthInPix=3784"
      Splits(0)._ColumnProps(103)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(104)=   "Column(22).Width=1931"
      Splits(0)._ColumnProps(105)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(106)=   "Column(22)._WidthInPix=1826"
      Splits(0)._ColumnProps(107)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(108)=   "Column(23).Width=3519"
      Splits(0)._ColumnProps(109)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(23)._WidthInPix=3413"
      Splits(0)._ColumnProps(111)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(112)=   "Column(24).Width=1931"
      Splits(0)._ColumnProps(113)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(114)=   "Column(24)._WidthInPix=1826"
      Splits(0)._ColumnProps(115)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(116)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(117)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(118)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(119)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(120)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(121)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(122)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(123)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(124)=   "Column(27).Width=3281"
      Splits(0)._ColumnProps(125)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(126)=   "Column(27)._WidthInPix=3175"
      Splits(0)._ColumnProps(127)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(128)=   "Column(28).Width=3281"
      Splits(0)._ColumnProps(129)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(130)=   "Column(28)._WidthInPix=3175"
      Splits(0)._ColumnProps(131)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(132)=   "Column(29).Width=3281"
      Splits(0)._ColumnProps(133)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(134)=   "Column(29)._WidthInPix=3175"
      Splits(0)._ColumnProps(135)=   "Column(29).Order=30"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   12577002
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=154,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=151,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=152,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=153,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=150,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=147,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=148,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=149,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=146,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=143,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=144,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=145,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=142,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=139,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=140,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=141,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=90,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=87,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=88,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=89,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=118,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=115,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=116,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=117,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=58,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=55,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=56,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=57,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=62,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=59,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=60,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=61,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=66,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=63,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=64,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=65,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=70,.parent=13,.alignment=3"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=67,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=68,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=69,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=74,.parent=13,.alignment=3"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=71,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=72,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=73,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=78,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=75,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=76,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=77,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=98,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=95,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=96,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=97,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=106,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=103,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=104,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=105,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=102,.parent=13,.alignment=3"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=99,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=100,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=101,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=114,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=111,.parent=14"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=112,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=113,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=94,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=91,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=92,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=93,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=86,.parent=13"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=83,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=84,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=85,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=122,.parent=13"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=119,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=120,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=121,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=110,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=107,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=108,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=109,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=126,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=123,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=124,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=125,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=130,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=127,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=128,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=129,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=134,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=131,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=132,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=133,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=138,.parent=13"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=135,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=136,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=137,.parent=17"
      _StyleDefs(156) =   "Named:id=33:Normal"
      _StyleDefs(157) =   ":id=33,.parent=0"
      _StyleDefs(158) =   "Named:id=34:Heading"
      _StyleDefs(159) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(160) =   ":id=34,.wraptext=-1"
      _StyleDefs(161) =   "Named:id=35:Footing"
      _StyleDefs(162) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(163) =   "Named:id=36:Selected"
      _StyleDefs(164) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(165) =   "Named:id=37:Caption"
      _StyleDefs(166) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(167) =   "Named:id=38:HighlightRow"
      _StyleDefs(168) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(169) =   "Named:id=39:EvenRow"
      _StyleDefs(170) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(171) =   "Named:id=40:OddRow"
      _StyleDefs(172) =   ":id=40,.parent=33"
      _StyleDefs(173) =   "Named:id=41:RecordSelector"
      _StyleDefs(174) =   ":id=41,.parent=34"
      _StyleDefs(175) =   "Named:id=42:FilterBar"
      _StyleDefs(176) =   ":id=42,.parent=33"
   End
   Begin MSComDlg.CommonDialog ComDialogXls 
      Left            =   7560
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label70 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2760
      X2              =   4440
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1920
      X2              =   4440
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   720
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIST DATA PELANGGAN"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label lblJml 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11715
   End
End
Attribute VB_Name = "FormDataPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit


Private OnLoad As Boolean

Private objPel As New ClsDataPelanggan
Private rsPel As ADODB.Recordset
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private myChassis As String

Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub cmdCSV_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsPel.RecordCount = 0 Then
3             Er.HeadCap = "Export To CSV !"
4             Er.InfoCap = "List Data Pelanggan Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".csv"
16        If Log.vinRetention = 1 Then
17           If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
18               GridPelanggan.Columns(6).Visible = True
19               GoTo lanjutexport
20           End If
21        End If
22        GridPelanggan.Columns(6).Visible = False
lanjutexport:
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
23        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
24        sSQL = sSQL & " values ('" & gSA & "','FormDataPelanggan',getdate(),'1900-01-01',getdate(),'Export CSV','" & rsPel.RecordCount & "','" & onVal & "','" & ComptName & "') "
25        Call objO.InsertData(sSQL)
26        GridPelanggan.ExportToDelimitedFile onVal, , ";", , , True 'GridPelanggan.ExportToFile onVal, True
27        GridPelanggan.Columns(6).Visible = True
28                Er.HeadCap = "Export to csv !"
29                Er.InfoCap = "Data sudah di Export ke File CSV." & Chr(13) & _
                               "File : " & onVal
                  
30                Beep
31                Er.show (vbModal)
32         Exit Sub
Err_Export:
33          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
34          MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdXls of Form FormDataPelanggan")
Keluar:
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdProcess_Click()
    Dim ft As New FormTransparan
    ft.xChassis = myChassis
    ft.xIsForm = "HIST"
    ft.show (vbModal)
End Sub

Private Sub CmdRefreshLap_Click()
   Call FillGridPelanggan
End Sub

Private Sub cmdSMS_Click()
          Dim objR        As New ClsRemainder
          Dim NewId As Double
          Dim rsSMS As ADODB.Recordset
1        On Error GoTo cmdSMS_Click_Error

2         Set rsSMS = objR.GetDataOutbox
          
3         If rsSMS.RecordCount = 0 Then
4             NewId = "1"
5         Else
6             rsSMS.MoveLast
7             NewId = Val(rsSMS.Fields(0).value) + 1
8         End If
          
9        If chGSM.value = 1 Then
10            With objR
11                .NoSMS = NewId
12                .pesan = IIf(txtPesan.Text = "", "-", txtPesan.Text)
13                .Tujuan = Replace(chGSM.Caption, "-", "")
14                .waktu = isnow
15                .AddNewItemSMS
16            End With
17            NewId = NewId + 1
18        End If
19       If chCDMA.value = 1 Then
20            With objR
21                .NoSMS = NewId
22                .pesan = IIf(txtPesan.Text = "", "-", txtPesan.Text)
23                .Tujuan = Replace(chCDMA.Caption, "-", "")
24                .waktu = isnow
25                .AddNewItemSMS
26            End With
27       End If
28       MsgBox "SMS dalam Proses Pengiriman !", vbInformation + vbOKOnly
29       On Error GoTo 0
30       Exit Sub
cmdSMS_Click_Error:
31        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSMS_Click of Form FormDataPelanggan")
End Sub

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsPel.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Pelanggan Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
16         If Log.vinRetention = 1 Then
17            If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
18                GridPelanggan.Columns(6).Visible = True
19                GoTo lanjutexport
20            End If
21         End If
22         GridPelanggan.Columns(6).Visible = False
lanjutexport:
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
23        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
24        sSQL = sSQL & " values ('" & gSA & "','FormDataPelanggan',getdate(),'1900-01-01',getdate(),'Export XLS','" & rsPel.RecordCount & "','" & onVal & "','" & ComptName & "') "
25        Call objO.InsertData(sSQL)
26         GridPelanggan.ExportToFile onVal, True
27         GridPelanggan.Columns(6).Visible = True
28                Er.HeadCap = "Export to excel !"
29                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
30                Beep
31                Er.show (vbModal)
32         Exit Sub
Err_Export:
33          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
34          MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdXls of Form FormDataPelanggan")
Keluar:
End Sub

Private Sub Command1_Click()
    FrSMS.Visible = True
End Sub

Private Sub Command2_Click()
    FrSMS.Visible = False
    txtPesan.Text = ""
End Sub

Private Sub Form_Load()
    '' 20200512DEV
    If Log.disableExportMra = 1 Then
        If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
            cmdCSV.Enabled = True
            CmdXls.Enabled = True
        Else
            cmdCSV.Enabled = False
            CmdXls.Enabled = False
        End If
    End If
   If OnLoad = False Then Exit Sub
   Call FillGridPelanggan
End Sub


Private Sub FillGridPelanggan()
1        On Error GoTo FillGridPelanggan_Error

2           Set rsPel = objPel.FetchDataPelanggan
3           If rsPel.RecordCount = 0 Then Exit Sub
4           Set Me.GridPelanggan.DataSource = rsPel
            
5           lblJml.Caption = rsPel.RecordCount
       
6        On Error GoTo 0
7        Exit Sub

FillGridPelanggan_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure FillGridPelanggan of Form FormDataPelanggan")
End Sub

Private Sub GridPelanggan_FilterChange()
On Error GoTo err
   Set cols = GridPelanggan.Columns
   Dim C As Integer
   
   C = GridPelanggan.col
   GridPelanggan.HoldFields
   rsPel.filter = getFilter()
   If rsPel.EOF = True Then
      For Each col In GridPelanggan.Columns
         col.FilterText = ""
      Next col
      rsPel.filter = ""
   End If
   GridPelanggan.col = C
   GridPelanggan.EditActive = True
   lblJml.Caption = rsPel.RecordCount
   Exit Sub
err:
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
Dim tmp As String
Dim n As Integer

   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         Select Case Trim(col.DataField)
            Case "JMLHISTORY", "LASTSVC", "KMAKHIR"
                tmp = tmp & col.DataField & " >= '" & CDbl(col.FilterText) & "'"
            Case Else
                tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
         End Select
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub GridPelanggan_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 23 Then rsPel.Sort = GridPelanggan.Columns(ColIndex).DataField & " " & Urut
4        GridPelanggan.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsPel.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormDataPelanggan")
End Sub

Private Sub GridPelanggan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsPel.EOF Then
        myChassis = NulSID(rsPel!NO_CHASSIS) ''GridPelanggan.Columns(10).value
        chGSM.Caption = NulSID(rsPel!GSM) ''GridPelanggan.Columns(6).value
        chCDMA.Caption = NulSID(rsPel!CDMA) ''GridPelanggan.Columns(7).value
    End If
End Sub
