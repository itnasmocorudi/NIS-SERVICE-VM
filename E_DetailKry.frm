VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form E_DetailKry 
   Caption         =   "Detail Karyawan"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   Icon            =   "E_DetailKry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabkry 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   4194304
      TabCaption(0)   =   "Identitas Karyawan"
      TabPicture(0)   =   "E_DetailKry.frx":08EE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(46)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picone"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSave"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdclose"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Mutasi Karyawan"
      TabPicture(1)   =   "E_DetailKry.frx":090A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "framut"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Keluarga Karyawan"
      TabPicture(2)   =   "E_DetailKry.frx":0926
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frakel"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pendidikan Karyawan"
      TabPicture(3)   =   "E_DetailKry.frx":0942
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FraPendidikan"
      Tab(3).ControlCount=   1
      Begin VB.Frame FraPendidikan 
         Height          =   6135
         Left            =   -74520
         TabIndex        =   135
         Top             =   840
         Width           =   12135
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   10560
            TabIndex        =   144
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   8880
            TabIndex        =   143
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   7560
            TabIndex        =   142
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6240
            TabIndex        =   141
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4920
            TabIndex        =   140
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3600
            TabIndex        =   139
            Top             =   360
            Width           =   1335
         End
         Begin VB.Frame FraisiPendidikan 
            Caption         =   "Pendidikan"
            Height          =   3495
            Left            =   360
            TabIndex        =   138
            Top             =   840
            Width           =   11535
            Begin VB.TextBox txtidPendidikan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   9840
               TabIndex        =   165
               Top             =   240
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CheckBox Chklulus 
               Height          =   255
               Left            =   2640
               TabIndex        =   161
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox txtpredikat 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7680
               TabIndex        =   160
               Top             =   1650
               Width           =   3735
            End
            Begin VB.TextBox txtalmamater 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7680
               TabIndex        =   159
               Top             =   1320
               Width           =   3735
            End
            Begin VB.ComboBox CboTahun 
               Height          =   315
               Index           =   1
               Left            =   7680
               Style           =   2  'Dropdown List
               TabIndex        =   158
               Top             =   600
               Width           =   2055
            End
            Begin VB.ComboBox CboTahun 
               Height          =   315
               Index           =   0
               Left            =   2640
               Style           =   2  'Dropdown List
               TabIndex        =   157
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtketPendidikan 
               Appearance      =   0  'Flat
               Height          =   1275
               Left            =   2640
               MultiLine       =   -1  'True
               TabIndex        =   156
               Top             =   2040
               Width           =   8775
            End
            Begin VB.TextBox txtnmpendidikan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2640
               TabIndex        =   145
               Top             =   240
               Width           =   7095
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbjenjang 
               Height          =   315
               Index           =   1
               Left            =   2640
               TabIndex        =   162
               Tag             =   "Kode"
               Top             =   960
               Width           =   3375
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   5953
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbpendidikanAkhir 
               Height          =   315
               Index           =   1
               Left            =   7680
               TabIndex        =   163
               Tag             =   "Kode"
               Top             =   960
               Width           =   3735
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   6588
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbJurusan 
               Height          =   315
               Index           =   1
               Left            =   2640
               TabIndex        =   164
               Tag             =   "Kode"
               Top             =   1320
               Width           =   3375
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   5953
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSAdodcLib.Adodc AdoJenjang 
               Height          =   345
               Index           =   1
               Left            =   360
               Top             =   960
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoPendidikanAkhir 
               Height          =   345
               Index           =   1
               Left            =   360
               Top             =   1320
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoJurusan 
               Height          =   345
               Index           =   1
               Left            =   360
               Top             =   1680
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   63
               Left            =   1200
               TabIndex        =   155
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Predikat Kelulusan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   62
               Left            =   5640
               TabIndex        =   154
               Top             =   1680
               Width           =   1935
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lulus :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   61
               Left            =   1200
               TabIndex        =   153
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Almamater :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   60
               Left            =   6240
               TabIndex        =   152
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Jurusan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   59
               Left            =   1200
               TabIndex        =   151
               Top             =   1320
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pendidikan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   58
               Left            =   6240
               TabIndex        =   150
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Jenjang  :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   57
               Left            =   1200
               TabIndex        =   149
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun Lulus :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   56
               Left            =   6240
               TabIndex        =   148
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun Masuk :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   55
               Left            =   1200
               TabIndex        =   147
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nama :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   54
               Left            =   1200
               TabIndex        =   146
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdPendidikan 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   136
            Top             =   360
            Width           =   1335
         End
         Begin TrueOleDBGrid70.TDBGrid GridPendidikan 
            Bindings        =   "E_DetailKry.frx":095E
            Height          =   1335
            Left            =   360
            TabIndex        =   137
            Top             =   4560
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   2355
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Thn Masuk"
            Columns(0).DataField=   "thnmasuk"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Thn Lulus"
            Columns(1).DataField=   "thnLulus"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Jenjang"
            Columns(2).DataField=   "jenjang"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Pendidikan"
            Columns(3).DataField=   "Pendidikan"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Jurusan"
            Columns(4).DataField=   "Jurusan"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Almamater"
            Columns(5).DataField=   "Almamater"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Lulus"
            Columns(6).DataField=   "lulus"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Predikat Kelulusan"
            Columns(7).DataField=   "Predikat"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Keterangan"
            Columns(8).DataField=   "Keterangan"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2302"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2143"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2037"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2328"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2223"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=3545"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3440"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=4683"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=4577"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=1640"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1535"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=3043"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2937"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=4101"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3995"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            AllowArrows     =   0   'False
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H69A3D6&"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=94,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=91,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=92,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=93,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
      Begin VB.PictureBox picone 
         BackColor       =   &H0080FFFF&
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5475
         ScaleWidth      =   12795
         TabIndex        =   43
         Top             =   1320
         Width           =   12855
         Begin VB.VScrollBar VScroll1 
            Height          =   5535
            Left            =   12480
            TabIndex        =   44
            Top             =   0
            Width           =   375
         End
         Begin VB.Frame fradd 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   0  'None
            Height          =   13080
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   12495
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbstatus 
               Height          =   315
               Left            =   8280
               TabIndex        =   174
               Tag             =   "Kode"
               Top             =   8715
               Width           =   3615
               DataFieldList   =   "Column 0"
               AllowInput      =   0   'False
               BevelType       =   0
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
               Row.Count       =   6
               Row(0)          =   "KARYAWAN TETAP"
               Row(1)          =   "KARYAWAN KONTRAK I"
               Row(2)          =   "KARYAWAN KONTRAK II"
               Row(3)          =   "KARYAWAN OUTSOURCING"
               Row(4)          =   "KARYAWAN MAGANG"
               Row(5)          =   "MARKETING"
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   6720
               Columns(0).Caption=   "Status"
               Columns(0).Name =   "Golongan"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               _ExtentX        =   6376
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSComctlLib.ListView lv 
               Height          =   30
               Index           =   0
               Left            =   2400
               TabIndex        =   134
               Top             =   9380
               Width           =   9255
               _ExtentX        =   16325
               _ExtentY        =   53
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.TextBox txtHomeBase 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   9120
               TabIndex        =   168
               Top             =   9840
               Width           =   2775
            End
            Begin VB.TextBox txtcatatan 
               Appearance      =   0  'Flat
               Height          =   670
               Left            =   2400
               MultiLine       =   -1  'True
               TabIndex        =   171
               Top             =   10200
               Width           =   4695
            End
            Begin TDBDate6Ctl.TDBDate DTTglKeluar 
               Height          =   330
               Left            =   5640
               TabIndex        =   170
               Top             =   9840
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               Calendar        =   "E_DetailKry.frx":0978
               Caption         =   "E_DetailKry.frx":0AA4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "E_DetailKry.frx":0B10
               Keys            =   "E_DetailKry.frx":0B2E
               Spin            =   "E_DetailKry.frx":0B8C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate DTTglmasuk 
               Height          =   330
               Left            =   2400
               TabIndex        =   169
               Top             =   9840
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               Calendar        =   "E_DetailKry.frx":0BB4
               Caption         =   "E_DetailKry.frx":0CE0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "E_DetailKry.frx":0D4C
               Keys            =   "E_DetailKry.frx":0D6A
               Spin            =   "E_DetailKry.frx":0DC8
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin VB.TextBox txtatasan 
               Enabled         =   0   'False
               Height          =   315
               Left            =   2400
               TabIndex        =   132
               Top             =   9070
               Width           =   9255
            End
            Begin VB.CommandButton cmddropatasan 
               BackColor       =   &H00404040&
               Caption         =   "+"
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
               Left            =   11640
               MaskColor       =   &H00808080&
               TabIndex        =   131
               Top             =   9060
               Width           =   255
            End
            Begin VB.Frame fra4 
               BackColor       =   &H0080FFFF&
               Height          =   3255
               Left            =   120
               TabIndex        =   94
               Top             =   7755
               Width           =   12255
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbGolongan 
                  Height          =   315
                  Left            =   8160
                  TabIndex        =   95
                  Tag             =   "Kode"
                  Top             =   600
                  Width           =   3615
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   6376
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbdealer 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   96
                  Tag             =   "Kode"
                  Top             =   240
                  Width           =   4575
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   8070
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbDivisi 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   97
                  Tag             =   "Kode"
                  Top             =   600
                  Width           =   4575
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   8070
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbpangkat 
                  Height          =   315
                  Left            =   8160
                  TabIndex        =   98
                  Tag             =   "Kode"
                  Top             =   240
                  Width           =   3615
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns.Count   =   3
                  Columns(0).Width=   3200
                  Columns(0).Caption=   "id_pangkat"
                  Columns(0).Name =   "id_pangkat"
                  Columns(0).DataField=   "Column 0"
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   256
                  Columns(1).Width=   3200
                  Columns(1).Caption=   "pangkat"
                  Columns(1).Name =   "pangkat"
                  Columns(1).DataField=   "Column 1"
                  Columns(1).DataType=   8
                  Columns(1).FieldLen=   256
                  Columns(2).Width=   3200
                  Columns(2).Caption=   "level"
                  Columns(2).Name =   "level"
                  Columns(2).DataField=   "Column 2"
                  Columns(2).DataType=   8
                  Columns(2).FieldLen=   256
                  _ExtentX        =   6376
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbatasan 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   100
                  Tag             =   "Kode"
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   375
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   661
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbDept 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   101
                  Tag             =   "Kode"
                  Top             =   960
                  Width           =   3615
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns.Count   =   3
                  Columns(0).Width=   3200
                  Columns(0).Caption=   "id_dept"
                  Columns(0).Name =   "id_dept"
                  Columns(0).DataField=   "Column 0"
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   256
                  Columns(1).Width=   3200
                  Columns(1).Caption=   "departemen"
                  Columns(1).Name =   "departemen"
                  Columns(1).DataField=   "Column 2"
                  Columns(1).DataType=   8
                  Columns(1).FieldLen=   256
                  Columns(2).Width=   3200
                  Columns(2).Caption=   "id_divisi"
                  Columns(2).Name =   "id_divisi"
                  Columns(2).DataField=   "Column 3"
                  Columns(2).DataType=   8
                  Columns(2).FieldLen=   256
                  _ExtentX        =   6376
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbPosisi 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   99
                  Tag             =   "Kode"
                  Top             =   1680
                  Width           =   4695
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   8281
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.TextBox txtgrade 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   9000
                  TabIndex        =   128
                  Top             =   1680
                  Width           =   2775
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Atasan :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   21
                  Left            =   360
                  TabIndex        =   133
                  Top             =   1320
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Catatan :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   33
                  Left            =   360
                  TabIndex        =   130
                  Top             =   2400
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Masuk :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   27
                  Left            =   360
                  TabIndex        =   126
                  Top             =   2040
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Dept/ Section (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   32
                  Left            =   480
                  TabIndex        =   108
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Golongan :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   20
                  Left            =   6240
                  TabIndex        =   107
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pangkat (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   22
                  Left            =   6240
                  TabIndex        =   106
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Divisi :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   23
                  Left            =   360
                  TabIndex        =   105
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Dealer (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   24
                  Left            =   360
                  TabIndex        =   104
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Posisi (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   26
                  Left            =   360
                  TabIndex        =   103
                  Top             =   1680
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   31
                  Left            =   6240
                  TabIndex        =   102
                  Top             =   960
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Homebase :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   25
                  Left            =   7560
                  TabIndex        =   127
                  Top             =   2040
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Keluar :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   28
                  Left            =   3600
                  TabIndex        =   125
                  Top             =   2040
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Grade :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   53
                  Left            =   7560
                  TabIndex        =   129
                  Top             =   1680
                  Width           =   1335
               End
            End
            Begin VB.Frame fra3 
               BackColor       =   &H0080FFFF&
               Height          =   1215
               Left            =   120
               TabIndex        =   85
               Top             =   6525
               Width           =   12255
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbjenjang 
                  Height          =   315
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   86
                  Tag             =   "Kode"
                  Top             =   360
                  Width           =   2295
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbJurusan 
                  Height          =   315
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   87
                  Tag             =   "Kode"
                  Top             =   360
                  Width           =   4815
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   8493
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbpendidikanAkhir 
                  Height          =   315
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   88
                  Tag             =   "Kode"
                  Top             =   720
                  Width           =   2295
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbGelar 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   89
                  Tag             =   "Kode"
                  Top             =   720
                  Width           =   2295
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jenjang Pendidikan (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   16
                  Left            =   0
                  TabIndex        =   93
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pendidikan Akhir (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   17
                  Left            =   120
                  TabIndex        =   92
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jurusan (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   18
                  Left            =   4920
                  TabIndex        =   91
                  Top             =   360
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Gelar :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   19
                  Left            =   4920
                  TabIndex        =   90
                  Top             =   720
                  Width           =   1815
               End
            End
            Begin VB.Frame fra2 
               BackColor       =   &H0080FFFF&
               Height          =   2535
               Left            =   120
               TabIndex        =   70
               Top             =   3960
               Width           =   12255
               Begin VB.TextBox txtKodePosKTP 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   112
                  Top             =   1320
                  Width           =   3015
               End
               Begin VB.TextBox txtEmail 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   110
                  Top             =   2040
                  Width           =   5055
               End
               Begin VB.TextBox txtAlmtTgal 
                  Appearance      =   0  'Flat
                  Height          =   670
                  Left            =   2280
                  MultiLine       =   -1  'True
                  TabIndex        =   77
                  Top             =   240
                  Width           =   5055
               End
               Begin VB.TextBox txtKotaTgal 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   76
                  Top             =   240
                  Width           =   3015
               End
               Begin VB.TextBox txtKdPos 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   75
                  Top             =   600
                  Width           =   3015
               End
               Begin VB.TextBox txtAlmtKTP 
                  Appearance      =   0  'Flat
                  Height          =   670
                  Left            =   2280
                  MultiLine       =   -1  'True
                  TabIndex        =   74
                  Top             =   960
                  Width           =   5055
               End
               Begin VB.TextBox txtKotaKtp 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   73
                  Top             =   960
                  Width           =   3015
               End
               Begin VB.TextBox txtTelpRmh 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   72
                  Top             =   1680
                  Width           =   2775
               End
               Begin VB.TextBox txtHp 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   71
                  Top             =   1680
                  Width           =   3015
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kode Pos KTP :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   48
                  Left            =   7200
                  TabIndex        =   111
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Email :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   47
                  Left            =   400
                  TabIndex        =   109
                  Top             =   2040
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Alamat Tinggal :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   9
                  Left            =   435
                  TabIndex        =   84
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kota Tinggal :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   10
                  Left            =   7440
                  TabIndex        =   83
                  Top             =   240
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kode Pos :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   11
                  Left            =   7560
                  TabIndex        =   82
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Alamat KTP :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   12
                  Left            =   780
                  TabIndex        =   81
                  Top             =   960
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kota KTP :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   13
                  Left            =   7560
                  TabIndex        =   80
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Telp Rumah :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   14
                  Left            =   405
                  TabIndex        =   79
                  Top             =   1680
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "No HP :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   15
                  Left            =   6960
                  TabIndex        =   78
                  Top             =   1680
                  Width           =   1815
               End
            End
            Begin MSAdodcLib.Adodc AdoAgama 
               Height          =   345
               Left            =   6480
               Top             =   480
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoStsKawin 
               Height          =   345
               Left            =   7080
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoJenjang 
               Height          =   345
               Index           =   0
               Left            =   7320
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoJurusan 
               Height          =   345
               Index           =   0
               Left            =   9000
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoPendidikanAkhir 
               Height          =   345
               Index           =   0
               Left            =   8520
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoGelar 
               Height          =   345
               Left            =   5280
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoCbgdetail 
               Height          =   345
               Left            =   7200
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoDivisi 
               Height          =   345
               Left            =   9000
               Top             =   480
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoGolongan 
               Height          =   345
               Left            =   7440
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc Adoatasan 
               Height          =   345
               Left            =   8160
               Top             =   480
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoPosisi 
               Height          =   345
               Left            =   5640
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSAdodcLib.Adodc AdoPangkat 
               Height          =   345
               Left            =   6840
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin MSComDlg.CommonDialog cdgopen 
               Left            =   8040
               Top             =   960
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSAdodcLib.Adodc AdoDept 
               Height          =   345
               Left            =   6240
               Top             =   600
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin VB.Frame Fra1 
               BackColor       =   &H0080FFFF&
               Height          =   3615
               Left            =   120
               TabIndex        =   49
               Top             =   360
               Width           =   12255
               Begin VB.TextBox txtNoKTP 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   124
                  Top             =   1320
                  Width           =   2775
               End
               Begin VB.TextBox TxtNoKK 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   123
                  Top             =   1680
                  Width           =   2775
               End
               Begin VB.TextBox TxtNPWP 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   122
                  Top             =   2040
                  Width           =   2775
               End
               Begin VB.TextBox txtNoBpjsSehat 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   118
                  Top             =   2400
                  Width           =   2775
               End
               Begin VB.TextBox txtJamsostek 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   114
                  Top             =   3120
                  Width           =   2775
               End
               Begin VB.TextBox txtNopolisAsuransi 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   113
                  Top             =   2760
                  Width           =   2775
               End
               Begin VB.CommandButton cmdbrowse 
                  Caption         =   "Browse"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   11040
                  TabIndex        =   56
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.TextBox txtNik 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   55
                  Top             =   240
                  Width           =   2775
               End
               Begin VB.TextBox txtNama 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   54
                  Top             =   600
                  Width           =   5535
               End
               Begin VB.OptionButton OptJK 
                  BackColor       =   &H0080FFFF&
                  Caption         =   "laki-laki"
                  Height          =   255
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   53
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.OptionButton OptJK 
                  BackColor       =   &H0080FFFF&
                  Caption         =   "Perempuan"
                  Height          =   255
                  Index           =   1
                  Left            =   3480
                  TabIndex        =   52
                  Top             =   1680
                  Width           =   1335
               End
               Begin VB.TextBox txtKewarganegaraan 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   51
                  Top             =   3120
                  Width           =   2295
               End
               Begin VB.TextBox txtTempatLahir 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   50
                  Top             =   960
                  Width           =   5535
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbAgama 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   57
                  Tag             =   "Kode"
                  Top             =   2040
                  Width           =   2295
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbStatusKawin 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   58
                  Tag             =   "Kode"
                  Top             =   2400
                  Width           =   2295
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   4096
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbGolDarah 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   59
                  Tag             =   "Kode"
                  Top             =   2760
                  Width           =   2295
                  DataFieldList   =   "Column 0"
                  AllowInput      =   0   'False
                  BevelType       =   0
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
                  Row.Count       =   4
                  Row(0)          =   "A"
                  Row(1)          =   "B"
                  Row(2)          =   "AB"
                  Row(3)          =   "O"
                  BevelColorHighlight=   -2147483634
                  BevelColorFace  =   -2147483627
                  CheckBox3D      =   0   'False
                  ForeColorEven   =   0
                  BackColorEven   =   8454143
                  BackColorOdd    =   65535
                  RowHeight       =   423
                  Columns(0).Width=   3200
                  Columns(0).Caption=   "Golongan"
                  Columns(0).Name =   "Golongan"
                  Columns(0).DataField=   "Column 0"
                  Columns(0).DataType=   8
                  Columns(0).FieldLen=   256
                  _ExtentX        =   4048
                  _ExtentY        =   556
                  _StockProps     =   93
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin TDBDate6Ctl.TDBDate DTTgllahir 
                  Height          =   330
                  Left            =   2280
                  TabIndex        =   60
                  Top             =   1320
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   582
                  Calendar        =   "E_DetailKry.frx":0DF0
                  Caption         =   "E_DetailKry.frx":0F1C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "E_DetailKry.frx":0F88
                  Keys            =   "E_DetailKry.frx":0FA6
                  Spin            =   "E_DetailKry.frx":1004
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  CursorPosition  =   0
                  DataProperty    =   0
                  DisplayFormat   =   "dd/mm/yyyy"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  FirstMonth      =   4
                  ForeColor       =   -2147483640
                  Format          =   "dd/mm/yyyy"
                  HighlightText   =   0
                  IMEMode         =   3
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxDate         =   2958465
                  MinDate         =   -657434
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  PromptChar      =   "_"
                  ReadOnly        =   0
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "__/__/____"
                  ValidateMode    =   0
                  ValueVT         =   2118189057
                  Value           =   2.12482692446619E-314
                  CenturyMode     =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "No KTP :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   52
                  Left            =   5040
                  TabIndex        =   121
                  Top             =   1320
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "No KK :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   51
                  Left            =   5040
                  TabIndex        =   120
                  Top             =   1680
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "NPWP :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   50
                  Left            =   5040
                  TabIndex        =   119
                  Top             =   2040
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "No BPJS Kesehatan :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   49
                  Left            =   5040
                  TabIndex        =   117
                  Top             =   2400
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "No BPJS Ketenagakerjaan :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   29
                  Left            =   4200
                  TabIndex        =   116
                  Top             =   3120
                  Width           =   2895
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "No Polis Asuransi :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   30
                  Left            =   5280
                  TabIndex        =   115
                  Top             =   2760
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "NIK (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   69
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   68
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tempat Lahir (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   67
                  Top             =   960
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Lahir (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   240
                  TabIndex        =   66
                  Top             =   1320
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jenis Kelamin (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   240
                  TabIndex        =   65
                  Top             =   1680
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Agama (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   240
                  TabIndex        =   64
                  Top             =   2040
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Status Perkawinan (*) :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   0
                  TabIndex        =   63
                  Top             =   2400
                  Width           =   2055
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Golongan Darah :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   7
                  Left            =   240
                  TabIndex        =   62
                  Top             =   2760
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kewarganegaraan :"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   8
                  Left            =   240
                  TabIndex        =   61
                  Top             =   3120
                  Width           =   1815
               End
               Begin VB.Image imgFoto 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   2295
                  Left            =   10200
                  Stretch         =   -1  'True
                  Top             =   240
                  Width           =   1935
               End
            End
            Begin VB.Label lblJudul 
               BackStyle       =   0  'Transparent
               Caption         =   "Tambah Data"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   50
               Width           =   2895
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00800000&
               BackStyle       =   1  'Opaque
               Height          =   340
               Left            =   0
               Top             =   0
               Width           =   12495
            End
         End
      End
      Begin VB.Frame framut 
         Height          =   5775
         Left            =   -74520
         TabIndex        =   39
         Top             =   1320
         Width           =   11895
         Begin VB.CommandButton cmdtutup 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdexcelmutasi 
            Caption         =   "export excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   167
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddMutasi 
            Caption         =   "Add Mutasi"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   166
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdrefresh 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
         Begin TrueOleDBGrid70.TDBGrid GridMutasi 
            Height          =   4455
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   7858
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "No SK"
            Columns(0).DataField=   "no sk"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tanggal SK"
            Columns(1).DataField=   "tanggal sk"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Status Mutasi"
            Columns(2).DataField=   "status mutasi"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Dealer Lama"
            Columns(3).DataField=   "dealer lama"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Divisi lama"
            Columns(4).DataField=   "divisiho_lama"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Departemen Lama"
            Columns(5).DataField=   "divisi lama"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Golongan Lama"
            Columns(6).DataField=   "golongan lama"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Pangkat Lama"
            Columns(7).DataField=   "pangkat lama"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Posisi Lama"
            Columns(8).DataField=   "posisi lama"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Tgl Awal Jabatan Lama"
            Columns(9).DataField=   "tgl awal jabatan lama"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Tgl Akhir Jabatan Lama"
            Columns(10).DataField=   "tgl akhir jabatan lama"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Dealer Baru"
            Columns(11).DataField=   "dealer baru"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Divisi baru"
            Columns(12).DataField=   "divisiho_baru"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Departemen baru"
            Columns(13).DataField=   "divisi baru"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Golongan Baru"
            Columns(14).DataField=   "golongan baru"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Pangkat Baru"
            Columns(15).DataField=   "pangkat baru"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "Posisi Baru"
            Columns(16).DataField=   "posisi baru"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "Tgl Awal Jabatan Baru"
            Columns(17).DataField=   "tanggal awal jabatan baru"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   18
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14933984
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=18"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
            Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
            Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
            Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
            Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
            Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
            Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
            Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
            Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
            Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
            Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
            Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
            Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(69)=   "Column(17).Width=3281"
            Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
            Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14933984
            RowDividerColor =   14933984
            RowSubDividerColor=   14933984
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
            _StyleDefs(26)  =   ":id=13,.fontname=Tahoma"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(29)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(30)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
            _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
            _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=102,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(13).Style:id=86,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(15).Style:id=94,.parent=13"
            _StyleDefs(101) =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
            _StyleDefs(104) =   "Splits(0).Columns(16).Style:id=98,.parent=13"
            _StyleDefs(105) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
            _StyleDefs(106) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
            _StyleDefs(107) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
            _StyleDefs(108) =   "Splits(0).Columns(17).Style:id=106,.parent=13"
            _StyleDefs(109) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
            _StyleDefs(110) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
            _StyleDefs(111) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
            _StyleDefs(112) =   "Named:id=33:Normal"
            _StyleDefs(113) =   ":id=33,.parent=0"
            _StyleDefs(114) =   "Named:id=34:Heading"
            _StyleDefs(115) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(116) =   ":id=34,.wraptext=-1"
            _StyleDefs(117) =   "Named:id=35:Footing"
            _StyleDefs(118) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(119) =   "Named:id=36:Selected"
            _StyleDefs(120) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(121) =   "Named:id=37:Caption"
            _StyleDefs(122) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(123) =   "Named:id=38:HighlightRow"
            _StyleDefs(124) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(125) =   "Named:id=39:EvenRow"
            _StyleDefs(126) =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
            _StyleDefs(127) =   "Named:id=40:OddRow"
            _StyleDefs(128) =   ":id=40,.parent=33"
            _StyleDefs(129) =   "Named:id=41:RecordSelector"
            _StyleDefs(130) =   ":id=41,.parent=34"
            _StyleDefs(131) =   "Named:id=42:FilterBar"
            _StyleDefs(132) =   ":id=42,.parent=33"
         End
         Begin MSComDlg.CommonDialog ComDialog 
            Left            =   8520
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.Frame frakel 
         Height          =   6015
         Left            =   -74040
         TabIndex        =   1
         Top             =   1020
         Width           =   10935
         Begin VB.CommandButton cmdclosekel 
            Caption         =   "&Close"
            Height          =   375
            Left            =   7440
            TabIndex        =   2
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   6000
            TabIndex        =   3
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdsavekel 
            Caption         =   "&Save"
            Height          =   375
            Left            =   4560
            TabIndex        =   4
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmddelkel 
            Caption         =   "&Delete"
            Height          =   375
            Left            =   3120
            TabIndex        =   5
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdeditkel 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CommandButton cmdrefkel 
            Caption         =   "Refresh"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdaddkel 
            Caption         =   "&Add"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Frame fraisikel 
            Caption         =   "Keluarga"
            Height          =   3255
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   10455
            Begin VB.ComboBox cmbhubungan 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   172
               Top             =   1280
               Width           =   2535
            End
            Begin VB.TextBox txtnamakel 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1920
               TabIndex        =   17
               Top             =   240
               Width           =   5175
            End
            Begin VB.OptionButton OptJK 
               Caption         =   "laki-laki"
               Height          =   255
               Index           =   2
               Left            =   1920
               TabIndex        =   16
               Top             =   960
               Width           =   1095
            End
            Begin VB.OptionButton OptJK 
               Caption         =   "Perempuan"
               Height          =   255
               Index           =   3
               Left            =   3360
               TabIndex        =   15
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtt4lahirkel 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1920
               TabIndex        =   14
               Top             =   600
               Width           =   3255
            End
            Begin VB.TextBox txtkerjaan 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   7080
               TabIndex        =   13
               Top             =   2040
               Width           =   3255
            End
            Begin VB.TextBox txtalmtKel 
               Appearance      =   0  'Flat
               Height          =   670
               Left            =   1920
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   2400
               Width           =   3255
            End
            Begin VB.TextBox txtketerangankel 
               Appearance      =   0  'Flat
               Height          =   670
               Left            =   7080
               MultiLine       =   -1  'True
               TabIndex        =   11
               Top             =   2400
               Width           =   3255
            End
            Begin VB.CheckBox chktanggung 
               Caption         =   "Tertanggung"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7800
               TabIndex        =   10
               Top             =   240
               Width           =   1455
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbagamakel 
               Height          =   315
               Left            =   7080
               TabIndex        =   18
               Tag             =   "Kode"
               Top             =   960
               Width           =   2295
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   4048
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin TDBDate6Ctl.TDBDate DttglLhrKel 
               Height          =   330
               Left            =   7080
               TabIndex        =   19
               Top             =   600
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               Calendar        =   "E_DetailKry.frx":102C
               Caption         =   "E_DetailKry.frx":1158
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "E_DetailKry.frx":11C4
               Keys            =   "E_DetailKry.frx":11E2
               Spin            =   "E_DetailKry.frx":1240
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd/mm/yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   -1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbgoldarahkel 
               Height          =   315
               Left            =   7080
               TabIndex        =   20
               Tag             =   "Kode"
               Top             =   1320
               Width           =   2295
               DataFieldList   =   "Column 0"
               AllowInput      =   0   'False
               BevelType       =   0
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
               Row.Count       =   4
               Row(0)          =   "A"
               Row(1)          =   "B"
               Row(2)          =   "AB"
               Row(3)          =   "O"
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).Caption=   "Golongan"
               Columns(0).Name =   "Golongan"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               _ExtentX        =   4048
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbHubkel 
               Height          =   315
               Left            =   1920
               TabIndex        =   21
               Tag             =   "Kode"
               Top             =   1680
               Width           =   2535
               DataFieldList   =   "Column 0"
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).Caption=   "Golongan"
               Columns(0).Name =   "Golongan"
               Columns(0).DataField=   "Column 0"
               Columns(0).DataType=   8
               Columns(0).FieldLen=   256
               _ExtentX        =   4471
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbstskawinkel 
               Height          =   315
               Left            =   7080
               TabIndex        =   22
               Tag             =   "Kode"
               Top             =   1680
               Width           =   2295
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   4048
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbpendidikanakhirkel 
               Height          =   315
               Left            =   1920
               TabIndex        =   23
               Tag             =   "Kode"
               Top             =   2040
               Width           =   2535
               AllowInput      =   0   'False
               BevelType       =   0
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
               BevelColorHighlight=   -2147483634
               BevelColorFace  =   -2147483627
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   8454143
               BackColorOdd    =   65535
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   4471
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSAdodcLib.Adodc AdoHubKel 
               Height          =   345
               Left            =   9240
               Top             =   120
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
               RecordSource    =   ""
               Caption         =   "AdoKdSup"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hubungan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   64
               Left            =   0
               TabIndex        =   173
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nama :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   480
               TabIndex        =   35
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Jenis Kelamin :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   360
               TabIndex        =   34
               Top             =   960
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tempat Lahir :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   0
               TabIndex        =   33
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Lahir :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   37
               Left            =   5160
               TabIndex        =   32
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Agama :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   5160
               TabIndex        =   31
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Golongan Darah :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   39
               Left            =   5160
               TabIndex        =   30
               Top             =   1320
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hubungan kel :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   40
               Left            =   0
               TabIndex        =   29
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Status Perkawinan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   41
               Left            =   5160
               TabIndex        =   28
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pendidikan Akhir :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   42
               Left            =   0
               TabIndex        =   27
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pekerjaan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   43
               Left            =   5160
               TabIndex        =   26
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat Tinggal :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   44
               Left            =   0
               TabIndex        =   25
               Top             =   2400
               Width           =   1815
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Keterangan :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   45
               Left            =   5160
               TabIndex        =   24
               Top             =   2400
               Width           =   1815
            End
         End
         Begin VB.TextBox txtnmkry 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Top             =   240
            Width           =   8775
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGridKeluarga 
            Bindings        =   "E_DetailKry.frx":1268
            Height          =   1335
            Left            =   240
            TabIndex        =   8
            Top             =   4560
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   2355
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   48
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "No Rec"
            Columns(0).DataField=   "norec"
            Columns(0).DataWidth=   10
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nama"
            Columns(1).DataField=   "Nama"
            Columns(1).DataWidth=   30
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tempat Lahir"
            Columns(2).DataField=   "Tempat Lahir"
            Columns(2).DataWidth=   15
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tanggal Lahir"
            Columns(3).DataField=   "Tanggal Lahir"
            Columns(3).DataWidth=   19
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Jenis Kelamin"
            Columns(4).DataField=   "Jenis Kelamin"
            Columns(4).DataWidth=   10
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Agama"
            Columns(5).DataField=   "Agama"
            Columns(5).DataWidth=   15
            Columns(5).DropDown=   "TDBDropAgama"
            Columns(5).DropDown.vt=   8
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Hubungan Keluarga"
            Columns(6).DataField=   "Hubkel"
            Columns(6).DataWidth=   15
            Columns(6).DropDown=   "TDBDropHubunganKlgr"
            Columns(6).DropDown.vt=   8
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Pendidikan Akhir"
            Columns(7).DataField=   "Pendidikan Akhir"
            Columns(7).DataWidth=   10
            Columns(7).DropDown=   "TDBDropPendidikan"
            Columns(7).DropDown.vt=   8
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Pekerjaan"
            Columns(8).DataField=   "Pekerjaan"
            Columns(8).DataWidth=   15
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Alamat"
            Columns(9).DataField=   "Alamat"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Tertanggung"
            Columns(10).DataField=   "Tertanggung"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Keterangan"
            Columns(11).DataField=   "Keterangan"
            Columns(11).DataWidth=   65535
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1429"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1323"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2752"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2090"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1984"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2619"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2514"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(3)._AlignLeft=0"
            Splits(0)._ColumnProps(18)=   "Column(4).Width=1931"
            Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=1826"
            Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(22)=   "Column(5).Width=2090"
            Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1984"
            Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(26)=   "Column(6).Width=2752"
            Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(30)=   "Column(7).Width=2355"
            Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2249"
            Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(34)=   "Column(8).Width=2090"
            Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=1984"
            Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(38)=   "Column(9).Width=3519"
            Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=3413"
            Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(42)=   "Column(10).Width=2752"
            Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(46)=   "Column(11).Width=2752"
            Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            AllowArrows     =   0   'False
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   14215660
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H69A3D6&"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=90,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=87,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=88,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=89,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=94,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=91,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=92,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=93,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=74,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
         End
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Ket : (*) Wajib diisikan"
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
         Height          =   255
         Index           =   46
         Left            =   2160
         TabIndex        =   48
         Top             =   840
         Width           =   2895
      End
   End
End
Attribute VB_Name = "E_DetailKry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public flag_ho As Boolean
Public Nik As String
Public Baru As String
Public judul As String
Dim id_divisi As String
Dim id_dept As String
Dim id_pangkat As String
Dim id_posisi As String
Dim nik_atasan As String
Dim id_dealer As String
Dim rsDtMutasi As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Dim col As TrueOleDBGrid70.Column
Dim Urut As String
Dim Kel_add As Boolean
Dim Pend_add As Boolean
Dim rsDtkel As New ADODB.Recordset
Dim rsDTPend As New ADODB.Recordset
Dim rsKry As New ADODB.Recordset

Private Function getFilter() As String
   On Error GoTo Vb_Error
   Dim tmp As String
   Dim n As Integer
   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         If col.ExternalEditor = "angka" Then
            tmp = tmp & col.DataField & " = " & col.FilterText & ""
         ElseIf col.ExternalEditor = "tgl" Then
            If Len(col.FilterText) <> 10 Then Exit Function
            tmp = tmp & col.DataField & " = '" & col.FilterText & "'"
         Else
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
         End If
      End If
   Next col
   getFilter = tmp
   Exit Function
Vb_Error:
   Resume Next
End Function
Private Sub clear_set()
    On Error Resume Next
    If rsDtkel.State = 1 Then rsDtkel.Close
    If rsDtMutasi.State = 1 Then rsDtMutasi.Close
    cmbAgama.DataSourceList = Nothing
    cmbStatusKawin.DataSourceList = Nothing
   cmbjenjang(0).DataSourceList = Nothing
   cmbJurusan(0).DataSourceList = Nothing
   cmbpendidikanAkhir(0).DataSourceList = Nothing
   cmbjenjang(1).DataSourceList = Nothing
   cmbJurusan(1).DataSourceList = Nothing
   cmbpendidikanAkhir(1).DataSourceList = Nothing
   cmbGelar.DataSourceList = Nothing
   cmbdealer.DataSourceList = Nothing
   cmbDivisi.DataSourceList = Nothing
   cmbDept.DataSourceList = Nothing
   cmbGolDarah.DataSourceList = Nothing
   cmbatasan.DataSourceList = Nothing
   cmbPosisi.DataSourceList = Nothing
   cmbpangkat.DataSourceList = Nothing
   cmbHubkel.DataSourceList = Nothing
End Sub

Private Sub CboTahun_DropDown(Index As Integer)
    
Dim X As Integer
    Dim Thn As Integer
    CboTahun(Index).Clear
    Thn = year(GetDate)
    For X = 1 To 130
        CboTahun(Index).AddItem Thn
        Thn = Thn - 1
    Next X
End Sub

Private Sub cmbAgama_DropDown()
    isi_agama
End Sub

Private Sub cmbagamakel_DropDown()
    isi_agama
End Sub


Private Sub cmbdealer_Click()
    If cmbdealer.Columns(2).value = True Then
        flag_ho = True
        lbl(23).Caption = "Divisi (*) :"
        cmbDivisi.Enabled = True
        cmbDept.text = ""
    Else
        lbl(23).Caption = "Divisi :": flag_ho = False
        cmbDivisi.Enabled = False
        cmbDivisi.text = ""
        cmbDept.text = ""
    End If
End Sub

Private Sub cmbdealer_DropDown()
    isi_cabang
End Sub

Private Sub cmbdivisi_DropDown()
    isi_divisi
End Sub
Private Sub cmbGelar_DropDown()
    isi_gelar
End Sub

Private Sub cmbGolongan_DropDown()
    isi_golongan
End Sub

Private Sub cmbHubkel_DropDown()
    isi_hubkel
End Sub

Private Sub cmbhubungan_DropDown()
    cmbhubungan.Clear
    cmbhubungan.AddItem "AYAH"
    cmbhubungan.AddItem "IBU"
    cmbhubungan.AddItem "SUAMI"
    cmbhubungan.AddItem "ISTRI"
    cmbhubungan.AddItem "ANAK"
    cmbhubungan.AddItem "KAKAK"
    cmbhubungan.AddItem "ADIK"
    cmbhubungan.AddItem "MERTUA"
    cmbhubungan.AddItem "KAKEK & NENEK"
    cmbhubungan.AddItem "WALI"
End Sub

Private Sub cmbjenjang_Click(Index As Integer)
    cmbpendidikanAkhir(Index) = ""
    cmbpendidikanAkhir(Index).value = ""
    cmbJurusan(Index) = ""
    cmbJurusan(Index).value = ""
End Sub

Private Sub cmbjenjang_DropDown(Index As Integer)
    isi_jenjang (Index)
End Sub

Private Sub cmbJurusan_DropDown(Index As Integer)
    isi_jurusan (Index)
End Sub
Private Sub cmbpangkat_DropDown()
    isi_pangkat
End Sub

Private Sub cmbpendidikanAkhir_Click(Index As Integer)
    cmbJurusan(Index) = ""
    cmbJurusan(Index).value = ""
End Sub

Private Sub cmbpendidikanAkhir_DropDown(Index As Integer)
    isi_pendAkhir (Index)
End Sub

Private Sub cmbpendidikanakhirkel_DropDown()
    isi_pendAkhirKeluarga
End Sub
Private Sub cmbStatusKawin_DropDown()
    isi_kawin
End Sub
Private Sub cmbstskawinkel_DropDown()
    isi_kawin
End Sub
Private Sub cmdaddkel_Click()
    If Trim(txtNik.text) = "" Then
        MsgBox "Nik musti ada jika belum ada create dahulu data karyawan!!!", vbInformation, "Info"
        Exit Sub
    End If
    bersih_kel
    Call tombol_kel(False, False, False, True, True, False, True)
    fraisikel.Enabled = True
    Kel_add = True
End Sub

Private Sub cmdAddMutasi_Click()
    If Tingkatuser = 1 And strLevelUser = "Cabang" Then
        With E_MutasiDealer
            .Nik = Trim(txtNik.text)
            .Show
        End With
    ElseIf Tingkatuser = 3 And strLevelUser = "Administrator" Then
        E_MutasiKry.flag_ho = flag_ho
        E_MutasiKry.new_record = True
        Call E_MutasiKry.isi_data(Trim(txtNik.text), Trim(cmbdealer.Columns(0).value), Trim(cmbdealer.text), IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)), IIf(flag_ho = True, Trim(cmbDivisi.text), ""), Trim(cmbDept.Columns(0).value), Trim(cmbDept.text), Trim(cmbGolongan.Columns(0).value), Trim(cmbGolongan.text), Trim(cmbpangkat.Columns(0).value), Trim(cmbpangkat.text), Trim(cmbPosisi.Columns(0).value), Trim(cmbPosisi.text))
        E_MutasiKry.Show vbModal
    Else
        MsgBox "Anda tidak berhak melakukan mutasi data karyawan"
    End If
End Sub

Private Sub cmdCancel_Click()
    bersih_kel
    isi_datakel
    Call tombol_kel(True, True, True, False, True, True, False)
End Sub

Private Sub cmdclosekel_Click()
    Unload Me
End Sub

Private Sub cmddelkel_Click()
On Error GoTo Salah
    Dim sqldel As String
    If TDBGridKeluarga.ApproxCount = 0 Then Exit Sub
    If MsgBox("Yakin akan menghapus data keluarga " & rsDtkel!Nama & " ?", vbYesNo, "Info") = vbNo Then Exit Sub
    sqldel = "delete from tkeluarga where norec = '" & Trim(rsDtkel!norec) & "'"
    If DbCon.State = 0 Then DbCon.Open StrCon
    DbCon.Execute sqldel
    MsgBox "Data berhasil dihapus"
    isi_datakel
    bersih_kel
    Exit Sub
Salah:
    MsgBox err.Description
End Sub

Private Sub cmddropatasan_Click()
    If lv(0).Height = 3500 Then
       lv(0).Height = 0
       cmddropatasan.Caption = "+"
    Else
       lv(0).Height = 3500
       cmddropatasan.Caption = "-"
       If txtatasan.text = "" Then
           isi_atasan
       End If
    End If
End Sub

Private Sub cmdeditkel_Click()
    fraisikel.Enabled = True
    Call tombol_kel(False, False, False, True, False, False, True)
    Kel_add = False
End Sub

Private Sub cmdexcelmutasi_Click()

  On Error GoTo Err_Export
  Dim nmFile As String
  ComDialog.InitDir = App.Path
  ComDialog.FileName = nmFile
  ComDialog.CancelError = True
  ComDialog.filter = "Excel Files (*.xls)|*.xls"
  'ComDialog.Filter = "Excel Files (*.csv)|*.csv"
  ComDialog.ShowSave
  nmFile = ComDialog.FileName
  GridMutasi.ExportToFile nmFile, True
  'gridmutasi.ExportToDelimitedFile NmFile, , ";"
  MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
  Exit Sub
Err_Export:
  If err.Number = 32755 Then GoTo Keluar
  MsgBox err.Description
Keluar:

End Sub

Private Sub cmdPendidikan_Click(Index As Integer)
      Dim SQL As String
1     On Error GoTo Salah
2         If Index = 0 Then 'refresh
3             txtnmpendidikan.text = txtNama.text
4             isi_pendKry
5         ElseIf Index = 1 Then 'add
6             If Trim(txtNik.text) = "" Then
7                 MsgBox "Nik karyawan kosong!!!"
8                 Exit Sub
9             End If
10            Call tombolPendidikan(False, False, False, False, True, True, False)
11            Pend_add = True
12            FraisiPendidikan.Enabled = True
13            bersih_isiPendidikan
14        ElseIf Index = 2 Then 'edit
15            Call tombolPendidikan(False, False, False, False, True, True, False)
16            Pend_add = False
17            FraisiPendidikan.Enabled = True
18        ElseIf Index = 3 Then 'delete
19            If MsgBox("Yakin akan menghapus data pendidikan " & Trim(cmbjenjang) & "?", vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
20            DbCon.Execute "delete from tpendidikan where nik = '" & Trim(txtNik.text) & "' and idpendkry = " & Trim(rsDTPend!idpendkry) & ""
21            Call log_user(strUserID, DealerClaim, "DelPendKry", cmbjenjang(1))
22            MsgBox "hapus data pendidikan karyawan berhasil", vbInformation, "Sukses"
23            isi_pendKry
24        ElseIf Index = 4 Then 'save
              
25            If CboTahun(0) = "" Then
26                MsgBox "Tahun masuk tidak boleh kosong!!!", vbInformation, "Warning"
27                Exit Sub
28            End If
29            If CboTahun(1) = "" Then
30                MsgBox "Tahun lulus tidak boleh kosong!!!", vbInformation, "Warning"
31                Exit Sub
32            End If
33            If cmbjenjang(1) = "" Then
34                MsgBox "Jenjang pendidikan tidak boleh kosong!!!", vbInformation, "Warning"
35                Exit Sub
36            End If
37            If cmbpendidikanAkhir(1) = "" Then
38                MsgBox "Pendidikan akhir tidak boleh kosong!!!", vbInformation, "Warning"
39                Exit Sub
40            End If
41            If cmbJurusan(1) = "" Then
42                MsgBox "Jurusan tidak boleh kosong!!!", vbInformation, "Warning"
43                Exit Sub
44            End If
45            If Chklulus.value = Unchecked Then
46                If MsgBox("Yakin tidak lulus??" & vbCrLf & "Yes untuk lanjut, No untuk batal", vbYesNo, "Konfirmasi") = vbNo Then
47                    Exit Sub
48                End If
49            End If
50            If Pend_add = True Then
51                SQL = "insert into tpendidikan (nik,[tahun masuk],[tahun lulus],jenjang,pendidikan,jurusan,almamater,lulus,[predikat kelulusan],keterangan,voptr,[date update]) values " & _
                        " ('" & Trim(txtNik.text) & "','" & CboTahun(0) & "','" & CboTahun(1) & "','" & Trim(cmbjenjang(1).Columns(0).value) & "','" & Trim(cmbpendidikanAkhir(1).Columns(0).value) & "','" & Trim(cmbJurusan(1).Columns(0).value) & "','" & Trim(txtalmamater.text) & "','" & IIf(Chklulus.value = 1, "Y", "T") & "','" & Trim(txtpredikat.text) & "','" & Trim(txtketPendidikan.text) & "','" & strUserID & "',getdate()) "
                  
52            Else
53                SQL = "update tpendidikan set [tahun masuk] = '" & CboTahun(0) & "',[tahun lulus] = '" & CboTahun(1) & "',jenjang = '" & Trim(cmbjenjang(1).Columns(0).value) & "',pendidikan = '" & Trim(cmbpendidikanAkhir(1).Columns(0).value) & "',jurusan = '" & Trim(cmbJurusan(1).Columns(0).value) & "',almamater = '" & Trim(txtalmamater.text) & "',lulus = '" & IIf(Chklulus.value = 1, "Y", "T") & "',[predikat kelulusan] = '" & Trim(txtpredikat.text) & "',keterangan = '" & Trim(txtketPendidikan.text) & "',voptr = '" & strUserID & "',[date update] = getdate() where idpendkry = " & Trim(txtidPendidikan.text) & ""
              
54            End If
55            DbCon.Execute SQL
              
56            MsgBox "Simpan data berhasil", vbInformation, "sukses"
57            Call tombolPendidikan(True, True, True, True, False, False, True)
58            FraisiPendidikan.Enabled = False
59            isi_pendKry
              
60        ElseIf Index = 5 Then 'cancel
61            Call tombolPendidikan(True, True, True, True, False, False, True)
62            FraisiPendidikan.Enabled = False
63            isi_pendKry
64        ElseIf Index = 6 Then 'close
65            Unload Me
66        End If
67        Exit Sub
Salah:
68        MsgBox Erl & vbCrLf & err.Description & vbCrLf & "cmdPendidikan_Click"
End Sub
Private Sub bersih_isiPendidikan()
    CboTahun(0).ListIndex = -1
    CboTahun(1).ListIndex = -1
    cmbjenjang(1) = ""
    cmbjenjang(1).value = ""
    cmbpendidikanAkhir(1) = ""
    cmbpendidikanAkhir(1).value = ""
    cmbJurusan(1) = ""
    cmbJurusan(1).value = ""
    txtalmamater = ""
    Chklulus.value = 0
    txtpredikat = ""
    txtketPendidikan = ""
    txtidPendidikan = ""
End Sub
Private Sub isi_pendKry()
    Dim SQL As String
    On Error GoTo Salah
    SQL = "select idPendKry,nik,[tahun masuk] thnmasuk,[tahun lulus] thnlulus,jenjang,pendidikan,jurusan,almamater,lulus,[predikat kelulusan] predikat,keterangan From tpendidikan where nik = '" & Trim(txtNik.text) & "'"
    Set rsDTPend = New ADODB.Recordset
    rsDTPend.Open SQL, DbCon, adOpenForwardOnly, adLockReadOnly
    If rsDTPend.RecordCount > 0 Then
        GridPendidikan.DataSource = rsDTPend
    Else
        GridPendidikan.DataSource = Nothing
    End If
    Exit Sub
Salah:
    MsgBox Erl & vbCrLf & err.Description & vbCrLf & "Isi_pendkry"
End Sub
Private Sub tombolPendidikan(refresh As Boolean, add As Boolean, Edit As Boolean, del As Boolean, save As Boolean, Cancel As Boolean, ttp As Boolean)
    cmdPendidikan(0).Enabled = refresh
    cmdPendidikan(1).Enabled = add
    cmdPendidikan(2).Enabled = Edit
    cmdPendidikan(3).Enabled = del
    cmdPendidikan(4).Enabled = save
    cmdPendidikan(5).Enabled = Cancel
    cmdPendidikan(6).Enabled = ttp
End Sub

Private Sub cmdrefkel_Click()
    isi_datakel
End Sub
Private Sub isi_datakel()
    Dim sqlkel As String
    'sqlkel = "select * from tkeluarga where nik = '" & Trim(txtNik.Text) & "'"
    
'    sqlkel = "select ISNULL(b.[Hubungan Keluarga],a.[Hubungan Keluarga] ) hubkel,b.norec norecklg,a.* " & _
'             " from tkeluarga a " & _
'             " left join HubunganKeluarga b on a.[Hubungan Keluarga] = cast(b.idhub as varchar(50)) where nik = '" & Trim(txtNik.Text) & "' "
    
    sqlkel = "select b.[Hubungan Keluarga] ,b.norec norecklg,b.idHub ,a.NIK,a.Nama,a.[Jenis Kelamin],a.Agama,b.[Hubungan Keluarga] hubkel,a.[Tempat Lahir],a.[Tanggal Lahir],a.[Golongan Darah],a.[Status Perkawinan],a.[Pendidikan Akhir],a.Pekerjaan,a.Alamat,isnull(a.Keterangan,'') ket,a.NoRec,a.Tertanggung,b.grup " & _
                " from tkeluarga a " & _
                " left join HubunganKeluarga b on a.[Hubungan Keluarga] = cast(b.idhub as varchar(50)) " & _
                " where nik = '" & Trim(txtNik.text) & "' and ISNULL(b.[Hubungan Keluarga],'') <> '' " & _
                " Union all" & _
                " select b.[Hubungan Keluarga] hubkel,b.norec norecklg,b.idHub ,a.NIK,a.Nama,a.[Jenis Kelamin],a.Agama,b.[Hubungan Keluarga],a.[Tempat Lahir],a.[Tanggal Lahir],a.[Golongan Darah],a.[Status Perkawinan],a.[Pendidikan Akhir],a.Pekerjaan,a.Alamat,isnull(a.Keterangan,'') ket,a.NoRec,a.Tertanggung,b.grup " & _
                " from tkeluarga a " & _
                " left join HubunganKeluarga b on a.[Hubungan Keluarga] = ltrim(rtrim(b.[Hubungan Keluarga])) " & _
                " where nik = '" & Trim(txtNik.text) & "' and ISNULL(b.[Hubungan Keluarga],'') <> ''"
    
    txtnmkry.text = txtNama.text
    If DbCon.State = 0 Then DbCon.Open StrCon
    Set rsDtkel = New ADODB.Recordset
    rsDtkel.Open sqlkel, DbCon
    TDBGridKeluarga.DataSource = Nothing
    If Not rsDtkel.EOF Then
        TDBGridKeluarga.DataSource = rsDtkel
        TDBGridKeluarga.refresh
    End If
End Sub

Private Sub cmdsavekel_Click()
On Error GoTo Salah
Dim norec As String
Dim SQL As String
    If Trim(txtnamakel.text) = "" Then
        MsgBox "Nama harus diisikan!!!", vbInformation, "Info"
        Exit Sub
    End If
    If DttglLhrKel.ValueIsNull = True Then
        MsgBox "Tgl lahir keluarga tidak boleh kosong!!!", vbInformation, "Info"
        Exit Sub
    End If
    If Kel_add = True Then
'        If UCase(cmbHubkel.Value) = "SENDIRI" Then
'            norec = Trim(txtNik.Text) & "-00"
'        ElseIf UCase(cmbHubkel.Value) = "ISTRI" Then
'            norec = Trim(txtNik.Text) & "-01"
'        ElseIf UCase(cmbHubkel.Value) = "ANAK PERTAMA" Then
'            norec = Trim(txtNik.Text) & "-11"
'        ElseIf UCase(cmbHubkel.Value) = "ANAK KEDUA" Then
'            norec = Trim(txtNik.Text) & "-12"
'        ElseIf UCase(cmbHubkel.Value) = "ANAK KETIGA" Then
'            norec = Trim(txtNik.Text) & "-13"
'        ElseIf UCase(cmbHubkel.Value) = "ANAK KEEMPAT" Then
'            norec = Trim(txtNik.Text) & "-14"
'        End If
        If cmbHubkel.Columns(1).value = "" Then
            MsgBox "Hubungan keluarga harus diinputkan"
            Exit Sub
        End If
        norec = Trim(txtNik.text) & "-" & cmbHubkel.Columns(1).value
        
        SQL = "insert into tkeluarga (nik,nama,[jenis kelamin],agama,[hubungan keluarga],[tempat lahir],[tanggal lahir],[golongan darah],[status perkawinan],[pendidikan akhir],pekerjaan,alamat,keterangan,voptr,[last updated],[date update],norec,tertanggung) values "
        SQL = SQL & "( '" & Trim(txtNik.text) & "','" & Trim(Replace(UCase(txtnamakel.text), "'", "''")) & "','" & IIf(OptJK(2).value = True, "L", "P") & "','" & Trim(cmbAgama.text) & "','" & Trim(cmbHubkel.value) & "','" & Trim(txtt4lahirkel.text) & "','" & Format(DttglLhrKel.value, "MM/DD/yyyy") & "','" & Trim(cmbgoldarahkel.value) & "','" & Trim(cmbstskawinkel.value) & "','" & IIf(cmbpendidikanakhirkel.text = "", "", (cmbpendidikanakhirkel.Columns(0).value)) & "','" & Trim(txtkerjaan.text) & "',' " & Replace(Trim(UCase(txtalmtKel.text)), vbNewLine, "", , , vbTextCompare) & " ',' " & Replace(Trim(UCase(txtketerangankel.text)), vbNewLine, "", , , vbTextCompare) & " ','" & strUserID & "',getdate(),getdate(),'" & norec & "','" & IIf(chktanggung.value = 1, "Y", "N") & "')"
        
    Else
        If Trim(Nul(rsDtkel!norec)) = "" Then
            norec = Trim(txtNik.text) & "-" & cmbHubkel.Columns(1).value
            SQL = "update tkeluarga set norec = '" & norec & "' ,nama = '" & Trim(Replace(UCase(txtnamakel.text), "'", "''")) & "',[jenis kelamin] = '" & IIf(OptJK(2).value = True, "L", "P") & "',agama = '" & Trim(cmbagamakel.Columns(0).value) & "',[hubungan keluarga] = '" & Trim(cmbHubkel.Columns(0).value) & "',[tempat lahir] = '" & Trim(txtt4lahirkel.text) & "', "
            SQL = SQL & " [tanggal lahir] = '" & Format(DttglLhrKel.value, "MM/DD/yyyy") & "',[golongan darah] ='" & Trim(cmbgoldarahkel.Columns(0).value) & "',[status perkawinan] = '" & Trim(cmbstskawinkel.Columns(0).value) & "',[pendidikan akhir] = '" & IIf(cmbpendidikanakhirkel.text = "", "", (cmbpendidikanakhirkel.Columns(0).value)) & "',pekerjaan = '" & Trim(txtkerjaan.text) & "',alamat = ' " & Replace(Trim(UCase(txtalmtKel.text)), vbNewLine, "", , , vbTextCompare) & " ',keterangan = ' " & Replace(Trim(UCase(txtketerangankel.text)), vbNewLine, "", , , vbTextCompare) & " ',voptr = '" & strUserID & "',[last updated]= getdate(),tertanggung = '" & IIf(chktanggung.value = 1, "Y", "N") & "' where nik = '" & Trim(rsDtkel!Nik) & "' and nama = '" & Trim(rsDtkel!Nama) & "'  "
        Else
            SQL = "update tkeluarga set nama = '" & Trim(Replace(UCase(txtnamakel.text), "'", "''")) & "',[jenis kelamin] = '" & IIf(OptJK(2).value = True, "L", "P") & "',agama = '" & Trim(cmbagamakel.Columns(0).value) & "',[hubungan keluarga] = '" & Trim(cmbHubkel.Columns(0).value) & "',[tempat lahir] = '" & Trim(txtt4lahirkel.text) & "',[tanggal lahir] = '" & Format(DttglLhrKel.value, "MM/DD/yyyy") & "',[golongan darah] ='" & Trim(cmbgoldarahkel.Columns(0).value) & "',[status perkawinan] = '" & Trim(cmbstskawinkel.Columns(0).value) & "', " & _
                  " [pendidikan akhir] = '" & IIf(cmbpendidikanakhirkel.text = "", "", (cmbpendidikanakhirkel.Columns(0).value)) & "',pekerjaan = '" & Trim(txtkerjaan.text) & "',alamat = ' " & Replace(Trim(UCase(txtalmtKel.text)), vbNewLine, "", , , vbTextCompare) & " ',keterangan = ' " & Replace(Trim(UCase(txtketerangankel.text)), vbNewLine, "", , , vbTextCompare) & " ',voptr = '" & strUserID & "',[last updated]= getdate(),tertanggung = '" & IIf(chktanggung.value = 1, "Y", "N") & "' where norec = '" & Trim(rsDtkel!norec) & "' "
                   
        End If
    End If
    If DbCon.State = 0 Then DbCon.Open StrCon
    DbCon.Execute SQL
    MsgBox "Simpan data berhasil", vbInformation, "Sukses"
    bersih_kel
    isi_datakel
    fraisikel.Enabled = False
    Call tombol_kel(True, True, True, False, False, True, False)
    Exit Sub
Salah:
    MsgBox err.Description
End Sub



Private Sub Form_Load()
   VScroll1.Min = 0
   VScroll1.Max = picone.ScaleHeight - fradd.Height
   VScroll1.LargeChange = picone.ScaleHeight
   VScroll1.SmallChange = picone.ScaleHeight / 5
   AdoAgama.ConnectionString = StrCon
   AdoStsKawin.ConnectionString = StrCon
   AdoJenjang(0).ConnectionString = StrCon
   AdoJenjang(1).ConnectionString = StrCon
   AdoJurusan(0).ConnectionString = StrCon
   AdoJurusan(1).ConnectionString = StrCon
   AdoPendidikanAkhir(0).ConnectionString = StrCon
   AdoPendidikanAkhir(1).ConnectionString = StrCon
   AdoGelar.ConnectionString = StrCon
   AdoCbgdetail.ConnectionString = StrCon
   AdoDivisi.ConnectionString = StrCon
   AdoDept.ConnectionString = StrCon
   AdoGolongan.ConnectionString = StrCon
   Adoatasan.ConnectionString = StrCon
   AdoPosisi.ConnectionString = StrCon
   AdoPangkat.ConnectionString = StrCon
   AdoHubKel.ConnectionString = StrCon
   tabkry.Tab = 0
'   Call isi_agama
'   Call isi_kawin
'   Call isi_jenjang
'   Call isi_pendAkhir
'   Call isi_gelar
'   Call isi_cabang
'   Call isi_divisi
'   Call isi_dept
'   Call isi_golongan
'   Call isi_pangkat
'   Call isi_hubkel
   Call tombol_kel(True, True, True, False, False, True, False)
   fraisikel.Enabled = False
   FraisiPendidikan.Enabled = False
   Call tombolPendidikan(True, True, True, True, False, False, True)
   bersih_kel
   txtNik.text = Trim(Nik)
   lblJudul.Caption = judul
    lv(0).ColumnHeaders.Clear
    lv(0).ColumnHeaders.add , , "Cek", 800
    lv(0).ColumnHeaders.add , , "Nik", 1500
    lv(0).ColumnHeaders.add , , "Nama", 5000
    lv(0).ColumnHeaders.add , , "Posisi", 4000
    lv(0).ColumnHeaders.add , , "Divisi", 3000
    
   If Tingkatuser = 1 Then
        fra4.Enabled = False
   Else
        fra4.Enabled = True
   End If
    
   If Baru = "Y" Then
        txtNik.Enabled = True
        bersih_isi
   ElseIf Baru = "N" Then
        txtNik.Enabled = False
        'flag_ho = E_Karyawan.lblflagho.Caption
        Call isi_data(txtNik.text)
   End If
   
End Sub

Private Sub GridMutasi_DblClick() 'like n edit
    If GridMutasi.ApproxCount = 0 Then Exit Sub
    If Tingkatuser = 3 And strLevelUser = "Administrator" Then
        E_MutasiKry.flag_ho = flag_ho
        E_MutasiKry.new_record = False
        Call E_MutasiKry.isi_dataedit(Trim(txtNik.text), rsDtMutasi![no sk])
        E_MutasiKry.Show vbModal
    ElseIf Tingkatuser = 1 And strLevelUser = "Cabang" Then
        E_MutasiKry.flag_ho = flag_ho
        E_MutasiKry.new_record = False
        Call E_MutasiKry.isi_dataedit(Trim(txtNik.text), rsDtMutasi![no sk])
        E_MutasiKry.Show vbModal
    Else
        MsgBox "Anda tidak berhak melakukan mutasi data karyawan"
    End If
End Sub

Private Sub GridMutasi_FilterChange()
On Error Resume Next
   Set cols = GridMutasi.Columns
   Dim C As Integer
   C = GridMutasi.col
   GridMutasi.HoldFields
   rsDtMutasi.filter = getFilter()
   If rsDtMutasi.EOF Then
      MsgBox "Data yang dimaksud tidak dapat ditemukan !", vbInformation
      For Each col In GridMutasi.Columns
         col.FilterText = ""
      Next col
      isi_data (Trim(txtNik.text))
      GridMutasi.SetFocus
      GridMutasi.col = C
      Exit Sub
   End If
   GridMutasi.col = C
   GridMutasi.EditActive = True
  
   Exit Sub

Vb_Error:
   If err.Number = 3001 Then MsgBox "Ada Kesalahan Inputan", vbInformation
   For Each col In GridMutasi.Columns
      col.FilterText = ""
   Next col
End Sub

Private Sub GridMutasi_HeadClick(ByVal ColIndex As Integer)
    Urut = IIf(Urut = "desc", "asc", "desc")
    If ColIndex <> 0 Then rsDtMutasi.Sort = "[" & GridMutasi.Columns(ColIndex).DataField & "] " & Urut
End Sub
'Private Sub cmbatasan_DropDown()
'    Dim rsFind As New ADODB.Recordset
'    Adoatasan.RecordSource = ""
'    Dim sql As String
'    If flag_ho = True Then
'        If Trim(cmbpangkat.Columns(2).Value) > 2 Then
'            sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat  where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).Value) & "' and a.divisiho = '" & Trim(cmbDivisi.Columns(0).Value) & "'  and a.[tanggal keluar] is null order by c.[level],a.nama "
'        Else
'            sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat  where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.[tanggal keluar] is null order by c.[level],a.nama "
'        End If
'    Else
'        If Trim(cmbpangkat.Columns(2).Value) > 3 Then
'            If UCase(Trim(cmbdealer.Columns(0).Value)) = "BRS" Then
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer in ('BRS','TGL')   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            ElseIf UCase(Trim(cmbdealer.Columns(0).Value)) = "BAM" Then
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer in ('KRA','BAM')   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            ElseIf UCase(Trim(cmbdealer.Columns(0).Value)) = "SOBA" Then
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer in ('KRA','SOBA')   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            Else
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).Value) & "'   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            End If
'        Else
'            sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.[tanggal keluar] is null order by c.[level],a.nama "
'        End If
'    End If
'
'    Set rsFind = DbCon.Execute(sql)
'    cmbatasan.Reset
'    If rsFind.BOF Then Exit Sub
'    Adoatasan.RecordSource = sql
'    Adoatasan.refresh
'    With cmbatasan
'            .DataSourceList = Adoatasan
'            .DataFieldList = "nik"
'            .DataFieldToDisplay = "nama"
'            .Columns(0).Width = 2000
'            .Columns(1).Width = 6000
'    End With
'    cmbatasan.refresh
'End Sub
Private Sub isi_atasan()
    Dim rsFind As New ADODB.Recordset
    Dim SQL As String
    Dim lv1 As ListItem
    On Error GoTo Salah
    If flag_ho = True Then
        If Trim(cmbpangkat.Columns(2).value) > 2 Then
            If UCase(Trim(cmbDept.Columns(0).value)) <> "KAIZEN" And UCase(Trim(cmbDept.Columns(0).value)) <> "INTERNAL AUDIT" Then
                SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat  where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).value) & "' and a.divisiho = '" & Trim(cmbDivisi.Columns(0).value) & "'  and a.[tanggal keluar] is null order by c.[level],a.nama "
            Else
                If Trim(cmbpangkat.Columns(2).value) = 3 Then
                    SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat  where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).value) & "'   and a.[tanggal keluar] is null order by c.[level],a.nama "
                Else
                    SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat  where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).value) & "' and a.divisi = '" & Trim(cmbDept.Columns(0).value) & "'  and a.[tanggal keluar] is null order by c.[level],a.nama "
                End If
            End If
        Else
            SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat  where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.[tanggal keluar] is null order by c.[level],a.nama "
        End If
    Else
        If Trim(cmbpangkat.Columns(2).value) > 3 Then
'            If UCase(Trim(cmbdealer.Columns(0).Value)) = "BRS" Then
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer in ('BRS','TGL')   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            ElseIf UCase(Trim(cmbdealer.Columns(0).Value)) = "BAM" Then
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer in ('KRA','BAM')   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            ElseIf UCase(Trim(cmbdealer.Columns(0).Value)) = "SOBA" Then
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer in ('KRA','SOBA')   and a.[tanggal keluar] is null order by c.[level],a.nama "
'            Else
'                sql = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.Text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).Value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).Value) & "'  and a.[tanggal keluar] is null order by c.[level],a.nama "
'            End If
             If Cek_Record("select dealer from dealer_atasan where dealer = '" & Trim(cmbdealer.Columns(0).value) & "'") = True Then
                SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.dealer in (select dealer_lain from dealer_atasan where dealer = '" & Trim(cmbdealer.Columns(0).value) & "')   and a.[tanggal keluar] is null order by c.[level],a.nama "
             Else
                SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.dealer = '" & Trim(cmbdealer.Columns(0).value) & "'  and a.[tanggal keluar] is null order by c.[level],a.nama "
             End If
        Else
            SQL = "SELECT nik,nama,posisi,divisi from datakaryawan a  left join e_pangkat c on a.pangkat = c.id_pangkat where a.nik <> '" & Trim(txtNik.text) & "' and c.[level] <= '" & Trim(cmbpangkat.Columns(2).value) & "' and a.[tanggal keluar] is null order by c.[level],a.nama "
        End If
    End If
    Set rsFind = DbCon.Execute(SQL)
   
    If rsFind.BOF Then Exit Sub
    rsFind.MoveFirst
    lv(0).ListItems.Clear
    Set lv1 = lv(0).ListItems.add(, , "")
        lv1.SubItems(1) = "x"
        lv1.SubItems(2) = "All"
    Do While Not rsFind.EOF
        Set lv1 = lv(0).ListItems.add(, , "")
        lv1.SubItems(1) = Trim(rsFind!Nik)
        lv1.SubItems(2) = Trim(rsFind!Nama)
        lv1.SubItems(3) = Trim(rsFind!posisi)
        lv1.SubItems(4) = Trim(rsFind!divisi)
        rsFind.MoveNext
    Loop
    Exit Sub
Salah:
    MsgBox err.Description, vbCritical, "Failure"
End Sub
Private Sub cmbDept_Click()
    cmbPosisi.text = ""
    cmbPosisi.value = ""
    If flag_ho = True Then
    If UCase(Trim(cmbDept.Columns(0).value)) = "KAIZEN" Or UCase(Trim(cmbDept.Columns(0).value)) = "INTERNAL AUDIT" Then
        lbl(23).Caption = "Divisi :"
        cmbDivisi.text = ""
        cmbDivisi.Enabled = False
'        cmbDept.DataSourceList = Nothing
'        cmbDept.Text = ""
    Else
        lbl(23).Caption = "Divisi (*) :"
'        cmbDept.DataSourceList = Nothing
'        cmbDept.Text = ""
        cmbDivisi.Enabled = True
        cmbDivisi.SetFocus
    End If
    End If
End Sub

Private Sub cmbDept_DropDown()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    AdoDept.RecordSource = ""
    If flag_ho = True Then
        If cmbDivisi.Enabled = True Then
            SQL = "SELECT id_dept,departemen,id_divisi from E_Dept where id_divisi = '" & Trim(cmbDivisi.value) & "' and flag_ho = 1"
        Else
            SQL = "SELECT id_dept,departemen,id_divisi from E_Dept where flag_ho = 1"
        End If
    Else
        SQL = "SELECT id_dept,departemen,id_divisi from E_Dept where flag_ho = 0"
    End If
    
    Set rsFind = DbCon.Execute(SQL)
    cmbDept.Reset
    If rsFind.BOF Then Exit Sub
    AdoDept.RecordSource = SQL
    AdoDept.refresh
    With cmbDept
            .DataSourceList = AdoDept
            .DataFieldList = "id_dept"
            .DataFieldToDisplay = "departemen"
            .Columns(0).Visible = False
            .Columns(1).Width = 6000
            .Columns(2).Visible = False
    End With
    rsFind.Close
    Set rsFind = Nothing
    cmbDept.refresh
End Sub

Private Sub isi_jurusan(Index As Integer)
    Dim rsFind As New ADODB.Recordset
    Dim SQL As String
    AdoJurusan(Index).RecordSource = ""
    SQL = "SELECT jurusan from tjurusan where pendidikan = '" & Trim(cmbpendidikanAkhir(Index).Columns(0).value) & "'"
    Set rsFind = DbCon.Execute(SQL)
    cmbJurusan(Index).Reset
    If rsFind.BOF Then Exit Sub
    AdoJurusan(Index).RecordSource = SQL
    AdoJurusan(Index).refresh
    With cmbJurusan(Index)
            .DataSourceList = AdoJurusan(Index)
            .DataFieldList = "jurusan"
            .DataFieldToDisplay = "jurusan"
            .Columns(0).Width = 4000
    End With
    rsFind.Close
    Set rsFind = Nothing
    cmbJurusan(Index).refresh
End Sub
Private Sub cmbpangkat_Click()
    cmbPosisi.text = ""
    cmbPosisi.value = ""
    If cmbpangkat.Columns(2).value <= 2 Then
        cmbDept.Columns(0).value = ""
        cmbDept.text = ""
        cmbDept.Enabled = False
    Else
        cmbDept.Enabled = True
    End If
End Sub
Private Sub cmbPosisi_DropDown()
    Dim rsFind As New ADODB.Recordset
    Dim SQL As String
    AdoPosisi.RecordSource = ""
    If flag_ho = False Then
        SQL = "SELECT id_posisi,posisi from e_posisi where flag_ho = 0 and id_dept = '" & Trim(cmbDept.Columns(0).value) & "' and id_pangkat = '" & Trim(cmbpangkat.Columns(0).value) & "'"
    Else
        If cmbpangkat.Columns(2).value <= 2 Then
            SQL = "SELECT id_posisi,posisi from e_posisi a where a.flag_ho = 1 and a.id_dept = '" & Trim(cmbDivisi.Columns(0).value) & "'  and id_pangkat = '" & Trim(cmbpangkat.Columns(0).value) & "'"
        Else
            SQL = "SELECT id_posisi,posisi from e_posisi a left join e_dept b on a.id_dept = b.id_dept and a.flag_ho = b.flag_ho where a.flag_ho = 1 and b.id_divisi = '" & Trim(cmbDivisi.Columns(0).value) & "' and a.id_dept = '" & Trim(cmbDept.Columns(0).value) & "' and id_pangkat = '" & Trim(cmbpangkat.Columns(0).value) & "'"
        End If
    End If
    Set rsFind = DbCon.Execute(SQL)
    cmbPosisi.Reset
    If rsFind.BOF Then Exit Sub
    AdoPosisi.RecordSource = SQL
    AdoPosisi.refresh
    With cmbPosisi
            .DataSourceList = AdoPosisi
            .DataFieldList = "id_posisi"
            .DataFieldToDisplay = "posisi"
            .Columns(0).Visible = False
            .Columns(1).Width = 4000
    End With
    rsFind.Close
    Set rsFind = Nothing
    cmbPosisi.refresh
End Sub

Private Sub cmdBrowse_Click()
    Dim PictBmp As String
    cdgopen.FileName = ""
    cdgopen.filter = "Pictures (*.jpg;*.jpeg;)|*.jpg;*.jpeg;"
    cdgopen.Flags = cdlOFNHideReadOnly
    cdgopen.ShowOpen
    PictBmp = cdgopen.FileName
    If PictBmp = "" Then Exit Sub
    If FileLen(PictBmp) > 210000 Then
        MsgBox "Ukuran foto terlalu besar, Ukuran foto maksimal 200 KB"
        Close #1
        Exit Sub
    End If
    If cdgopen.FileName = "" Then Exit Sub
    imgFoto.Picture = LoadPicture(cdgopen.FileName)
End Sub
Private Sub cmdClose_Click()
    clear_set
    'Unload E_DetailKry
    Unload Me
End Sub
Private Sub cmdRefresh_Click()
    Dim SQL As String
    SQL = "Select * from tmutasi where nik = '" & Trim(txtNik.text) & "'"
    If DbCon.State = 0 Then DbCon.Open StrCon
    Set rsDtMutasi = New ADODB.Recordset
    rsDtMutasi.Open SQL, DbCon
    GridMutasi.DataSource = Nothing
    If Not rsDtMutasi.EOF Then
        GridMutasi.DataSource = rsDtMutasi
        GridMutasi.refresh
    Else
        MsgBox "Data tidak ada!!!", vbInformation, "Info"
    End If
End Sub

Private Sub cmdSave_Click()
    Dim SQL As String
    Dim mStream As New ADODB.Stream
    Dim rsFoto As New ADODB.Recordset
    Dim atasan As String
    Dim catatan As String
    Dim i As Integer
    Dim PakaiDivisi As Boolean
    
    On Error GoTo Salah
    If Trim(txtNik.text) = "" Then
        MsgBox "Nik Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(txtNama.text) = "" Then
        MsgBox "Nama Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(txtTempatLahir.text) = "" Then
        MsgBox "Tempat Lahir Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If IsNull(DTTgllahir.text) = True Then
        MsgBox "Tgl Lahir Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If OptJK(0).value = False And OptJK(1).value = False Then
        MsgBox "Jenis kelamin Wajib dipilih!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbAgama.text) = "" Then
        MsgBox "Agama Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbStatusKawin.text) = "" Then
        MsgBox "Status Perkawinan Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbjenjang(0).text) = "" Then
        MsgBox "Jenjang Pendidikan Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbpendidikanAkhir(0).text) = "" Then
        MsgBox "Pendidikan Akhir Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbJurusan(0).text) = "" Then
        MsgBox "Jurusan Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbdealer.text) = "" Then
        MsgBox "Dealer Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If flag_ho = True Then
        If UCase(Trim(cmbDept.Columns(0).value)) <> "KAIZEN" And UCase(Trim(cmbDept.Columns(0).value)) <> "INTERNAL AUDIT" Then
            If Trim(cmbDivisi.text) = "" Then
                MsgBox "Divisi Wajib diisikan!!!", vbInformation, "Info": Exit Sub
            End If
        End If
        If cmbpangkat.Columns(2).value > 2 Then
            If Trim(cmbDept.text) = "" Then
                MsgBox "Departemen Wajib diisikan!!!", vbInformation, "Info": Exit Sub
            End If
        End If
    Else
        If Trim(cmbDept.text) = "" Then
            MsgBox "Departemen Wajib diisikan!!!", vbInformation, "Info": Exit Sub
        End If
    End If
    
    If Trim(cmbstatus.text) = "" Then
        MsgBox "Status karyawan wajib diisikan", vbInformation, "Info"
        Exit Sub
    End If
    
    If UCase(Trim(cmbstatus.text)) <> "KARYAWAN KONTRAK I" And UCase(Trim(cmbstatus.text)) <> "KARYAWAN KONTRAK II" And UCase(Trim(cmbstatus.text)) <> "KARYAWAN TETAP" And UCase(Trim(cmbstatus.text)) <> "KARYAWAN OUTSOURCING" And UCase(Trim(cmbstatus.text)) <> "KARYAWAN MAGANG" And UCase(Trim(cmbstatus.text)) <> "MARKETING" Then
        MsgBox "Status karyawan yang dimasukkan tidak sesuai format!!! (ex:KARYAWAN KONTRAK I,KARYAWAN KONTRAK II,KARYAWAN TETAP,KARYAWAN OUTSOURCING,KARYAWAN MAGANG,MARKETING)", vbInformation, "Info"
        Exit Sub
    End If
    
    If Trim(cmbpangkat.text) = "" Then
        MsgBox "Pangkat Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    If Trim(cmbPosisi.text) = "" Then
        MsgBox "Posisi Wajib diisikan!!!", vbInformation, "Info": Exit Sub
    End If
    For i = 1 To lv(0).ListItems.Count
        If lv(0).ListItems(i).Checked = True Then
            If lv(0).ListItems(i).SubItems(1) <> "x" Then
                atasan = atasan & IIf(atasan = "", "", ",") & lv(0).ListItems(i).SubItems(1)
            End If
        End If
    Next i
    
    If DTTglKeluar <> "" Then
        If MsgBox("Apakah karyawan yang bersangkutan Pensiun?", vbYesNo + vbQuestion, "Konfirmasi") = vbNo Then
balik:
            catatan = InputBox("Masukkan alasan Keluar!!", "Alasan")
            If catatan = "" Then
                If MsgBox("Alasan keluar harus diisikan!!! " & vbCrLf & " Klik Ok untuk isikan alasan, cancel untuk batal simpan!!!", vbOKCancel, "Warning") = vBok Then
                    GoTo balik
                Else
                    Exit Sub
                End If
            Else
                txtcatatan.text = catatan
            End If
        Else
            txtcatatan.text = "PENSIUN"
        End If
    End If
    
    PakaiDivisi = cmbDivisi.Enabled
    Me.MousePointer = vbHourglass
    If cek_nik(txtNik.text) = False Then 'add new
         SQL = "insert into datakaryawan (nik,nama,[tempat lahir],[tanggal lahir],[jenis kelamin],[agama],[status perkawinan],[golongan darah],kewarganegaraan,[alamat tinggal],[kota tinggal],[kode pos],[alamat ktp],[kota ktp],[telp rumah],hp,[jenjang pendidikan],[pendidikan akhir],jurusan, gelar,golongan,pangkat,divisi,dealer,homebase,posisi,[tanggal masuk],[tanggal keluar],[no jamsostek],[no polis asuransi],catatan,voptr,[last updated],norec,divisiho,status,nik_atasan,nama_atasan,noktp,nokk,npwp,nobpjskesehatan,kodeposktp,email) values  "
         SQL = SQL & vbCrLf & " ('" & Trim(txtNik.text) & "','" & Trim(Replace(UCase(txtNama.text), "'", "''")) & "','" & Trim(UCase(txtTempatLahir.text)) & "','" & Format(DTTgllahir.value, "MM/DD/yyyy") & "','" & IIf(OptJK(0).value = True, "L", "P") & "','" & Trim(cmbAgama.value) & "','" & Trim(cmbStatusKawin.value) & "','" & Trim(cmbGolDarah.value) & "','" & Trim(UCase(txtKewarganegaraan.text)) & "',' " & Replace(Trim(UCase(txtAlmtTgal.text)), vbNewLine, "", , , vbTextCompare) & " ','" & Trim(UCase(txtKotaTgal.text)) & "','" & Trim(txtKdPos.text) & "',' " & Replace(UCase(Trim(txtAlmtKTP.text)), vbNewLine, "", , , vbTextCompare) & " ','" & Trim(UCase(txtKotaKtp.text)) & "', "
         SQL = SQL & vbCrLf & " '" & Trim(txtTelpRmh.text) & "','" & Trim(txtHp.text) & "','" & Trim(cmbjenjang(0).value) & "','" & Trim(cmbpendidikanAkhir(0).value) & "','" & Trim(cmbJurusan(0).value) & "','" & Trim(cmbGelar.value) & "','" & Trim(cmbGolongan.value) & "','" & Trim(cmbpangkat.Columns(0).value) & "','" & Trim(cmbDept.Columns(0).value) & "','" & Trim(cmbdealer.Columns(0).value) & "','" & Trim(UCase(txtHomeBase.text)) & "','" & Trim(cmbPosisi.Columns(0).value) & "','" & Format(DTTglmasuk.value, "MM/DD/yyyy") & "'," & IIf(DTTglKeluar.ValueIsNull = True, "Null", "'" & Format(DTTglKeluar.value, "MM/DD/yyyy") & "'") & ", "
         SQL = SQL & vbCrLf & " '" & Trim(txtJamsostek.text) & "','" & Trim(txtNopolisAsuransi.text) & "',' " & Replace(Trim(UCase(txtcatatan.text)), vbNewLine, "", , , vbTextCompare) & " ','" & strUserID & "',getdate(),'" & Trim(txtNik.text & "-00") & "', '" & IIf(PakaiDivisi = True, cmbDivisi.Columns(0).value, cmbDept.Columns(2).value) & "','" & Trim(cmbstatus.value) & "','" & Trim(atasan) & "','" & Trim(Replace(txtatasan.text, "'", "''")) & "','" & Trim(UCase(txtNoKTP.text)) & "','" & Trim(UCase(TxtNoKK.text)) & "','" & Trim(UCase(TxtNPWP.text)) & "','" & Trim(UCase(txtNoBpjsSehat.text)) & "','" & Trim(UCase(txtKodePosKTP.text)) & "','" & Trim(txtEmail.text) & "') "
         If DbCon.State = 0 Then DbCon.Open StrCon
         DbCon.Execute SQL
         'simpan gambar
         If cdgopen.FileName <> "" Then
         
         Me.MousePointer = vbHourglass
         SQL = "select * from e_fotokry where nik = '" & Trim(txtNik.text) & "'"
         Set rsFoto = New ADODB.Recordset
         rsFoto.Open SQL, DbCon, adOpenKeyset, adLockOptimistic
          
         With mStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile cdgopen.FileName
            If Not rsFoto.EOF Then
                rsFoto("foto").value = .Read
                rsFoto.Update
            Else
                rsFoto.AddNew
                rsFoto("nik").value = Trim(txtNik.text)
                rsFoto("foto").value = .Read
                rsFoto("last_update").value = GetDate
                rsFoto("user_update").value = UserID
                rsFoto.Update
            End If
         End With
         Set mStream = Nothing
         Set rsFoto = Nothing
         Me.MousePointer = vbDefault
         End If
         Call log_user(strUserID, DealerClaim, "I", "Simpan Kry " & Trim(txtNik.text))
    Else ' edit
        If id_dealer <> Trim(cmbdealer.Columns(0).value) And id_dealer <> "" Then
            Select Case MsgBox("Ada perubahan data Dealer, penggantian data harus lewat mutasi karyawan!!!" & vbCrLf & "Tekan Yes untuk mengubah data via mutasi,No untuk melanjutkan, cancel untuk membatalkan", vbYesNoCancel, "Konfirmasi")
                Case vbYes
                    E_MutasiKry.flag_ho = flag_ho
                    Call E_MutasiKry.isi_data(Trim(txtNik.text), Trim(cmbdealer.Columns(0).value), Trim(cmbdealer.text), IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)), IIf(flag_ho = True, Trim(cmbDivisi.text), ""), Trim(cmbDept.Columns(0).value), Trim(cmbDept.text), Trim(cmbGolongan.Columns(0).value), Trim(cmbGolongan.text), Trim(cmbpangkat.Columns(0).value), Trim(cmbpangkat.text), Trim(cmbPosisi.Columns(0).value), Trim(cmbPosisi.text))
                    E_MutasiKry.Show vbModal
                    GoTo lanjut
                    Exit Sub
                Case vbNo
                    If MsgBox("Yakin data disimpan!!! perubahan data tidak terekam histori datanya!!!", vbOKCancel, "Konfirmasi") = vbCancel Then Exit Sub
                Case vbCancel
                    Exit Sub
            End Select
        End If
        If id_divisi <> IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)) And id_divisi <> "" Then
            Select Case MsgBox("Ada perubahan data divisi/departemen, penggantian data harus lewat mutasi karyawan!!!" & vbCrLf & "Tekan Yes untuk mengubah data via mutasi,No untuk melanjutkan, cancel untuk membatalkan", vbYesNoCancel, "Konfirmasi")
                Case vbYes
                    E_MutasiKry.flag_ho = flag_ho
                    Call E_MutasiKry.isi_data(Trim(txtNik.text), Trim(cmbdealer.Columns(0).value), Trim(cmbdealer.text), IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)), IIf(flag_ho = True, Trim(cmbDivisi.text), ""), Trim(cmbDept.Columns(0).value), Trim(cmbDept.text), Trim(cmbGolongan.Columns(0).value), Trim(cmbGolongan.text), Trim(cmbpangkat.Columns(0).value), Trim(cmbpangkat.text), Trim(cmbPosisi.Columns(0).value), Trim(cmbPosisi.text))
                    E_MutasiKry.Show vbModal
                    GoTo lanjut
                    Exit Sub
                Case vbNo
                    If MsgBox("Yakin data disimpan!!! perubahan data tidak terekam histori datanya!!!", vbOKCancel, "Konfirmasi") = vbCancel Then Exit Sub
                Case vbCancel
                    Exit Sub
            End Select
        End If
        If id_dept <> Trim(cmbDept.Columns(0).value) And id_dept <> "" Then
            Select Case MsgBox("Ada perubahan data departemen, penggantian data harus lewat mutasi karyawan!!!" & vbCrLf & "Tekan Yes untuk mengubah data via mutasi,No untuk melanjutkan, cancel untuk membatalkan", vbYesNoCancel, "Konfirmasi")
                Case vbYes
                    E_MutasiKry.flag_ho = flag_ho
                    Call E_MutasiKry.isi_data(Trim(txtNik.text), Trim(cmbdealer.Columns(0).value), Trim(cmbdealer.text), IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)), IIf(flag_ho = True, Trim(cmbDivisi.text), ""), Trim(cmbDept.Columns(0).value), Trim(cmbDept.text), Trim(cmbGolongan.Columns(0).value), Trim(cmbGolongan.text), Trim(cmbpangkat.Columns(0).value), Trim(cmbpangkat.text), Trim(cmbPosisi.Columns(0).value), Trim(cmbPosisi.text))
                    E_MutasiKry.Show vbModal
                    GoTo lanjut
                    Exit Sub
                Case vbNo
                    If MsgBox("Yakin data disimpan!!! perubahan data tidak terekam histori datanya!!!", vbOKCancel, "Konfirmasi") = vbCancel Then Exit Sub
                Case vbCancel
                    Exit Sub
            End Select
        End If
        If id_pangkat <> Trim(cmbpangkat.Columns(0).value) And id_pangkat <> "" Then
            Select Case MsgBox("Ada perubahan data pangkat, penggantian data harus lewat mutasi karyawan!!!" & vbCrLf & "Tekan Yes untuk mengubah data via mutasi,No untuk melanjutkan, cancel untuk membatalkan", vbYesNoCancel, "Konfirmasi")
                Case vbYes
                    E_MutasiKry.flag_ho = flag_ho
                    Call E_MutasiKry.isi_data(Trim(txtNik.text), Trim(cmbdealer.Columns(0).value), Trim(cmbdealer.text), IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)), IIf(flag_ho = True, Trim(cmbDivisi.text), ""), Trim(cmbDept.Columns(0).value), Trim(cmbDept.text), Trim(cmbGolongan.Columns(0).value), Trim(cmbGolongan.text), Trim(cmbpangkat.Columns(0).value), Trim(cmbpangkat.text), Trim(cmbPosisi.Columns(0).value), Trim(cmbPosisi.text))
                    E_MutasiKry.Show vbModal
                    GoTo lanjut
                    Exit Sub
                Case vbNo
                    If MsgBox("Yakin data disimpan!!! perubahan data tidak terekam histori datanya!!!", vbOKCancel, "Konfirmasi") = vbCancel Then Exit Sub
                Case vbCancel
                    Exit Sub
            End Select
        End If
        If id_posisi <> Trim(cmbPosisi.Columns(0).value) And id_posisi <> "" Then
            Select Case MsgBox("Ada perubahan data posisi, penggantian data harus lewat mutasi karyawan!!!" & vbCrLf & "Tekan Yes untuk mengubah data via mutasi,No untuk melanjutkan, cancel untuk membatalkan", vbYesNoCancel, "Konfirmasi")
                Case vbYes
                    E_MutasiKry.flag_ho = flag_ho
                    Call E_MutasiKry.isi_data(Trim(txtNik.text), Trim(cmbdealer.Columns(0).value), Trim(cmbdealer.text), IIf(flag_ho = True, Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(2).value)), IIf(flag_ho = True, Trim(cmbDivisi.text), ""), Trim(cmbDept.Columns(0).value), Trim(cmbDept.text), Trim(cmbGolongan.Columns(0).value), Trim(cmbGolongan.text), Trim(cmbpangkat.Columns(0).value), Trim(cmbpangkat.text), Trim(cmbPosisi.Columns(0).value), Trim(cmbPosisi.text))
                    E_MutasiKry.Show vbModal
                    GoTo lanjut
                    Exit Sub
                Case vbNo
                    If MsgBox("Yakin data disimpan!!! perubahan data tidak terekam histori datanya!!!", vbOKCancel, "Konfirmasi") = vbCancel Then Exit Sub
                Case vbCancel
                    Exit Sub
            End Select
        End If
        If nik_atasan <> Trim(atasan) And nik_atasan <> "" Then
            Select Case MsgBox("Ada perubahan data atasan, penggantian data harus lewat edit atasan!!!" & vbCrLf & "Tekan Yes untuk mengubah data edit atasan,No untuk melanjutkan, cancel untuk membatalkan", vbYesNoCancel, "Konfirmasi")
                Case vbYes
                    With E_GantiAtasan
                        .lv(0).ColumnHeaders.Clear
                        .lv(0).ColumnHeaders.add , , "Cek", 800
                        .lv(0).ColumnHeaders.add , , "Nik", 1500
                        .lv(0).ColumnHeaders.add , , "Nama", 5000
                        .lv(0).ColumnHeaders.add , , "Posisi", 4000
                        .lv(0).ColumnHeaders.add , , "Divisi", 3000
                        .lblpangkat.Caption = Trim(cmbpangkat.Columns(2).value)
                        .lbldealer.Caption = Trim(cmbdealer.Columns(0).value)
                        .lbldivisi.Caption = Trim(cmbDivisi.Columns(0).value)
                        .lbldept.Caption = Trim(cmbDept.Columns(0).value)
                    End With
                    Call E_GantiAtasan.isi_atasan(Trim(cmbpangkat.Columns(2).value), Trim(cmbdealer.Columns(0).value), Trim(cmbDivisi.Columns(0).value), Trim(cmbDept.Columns(0).value))
                    Call E_GantiAtasan.isi_data(Trim(txtNik.text), Trim(atasan), Trim(txtNama.text))
                    E_GantiAtasan.Show vbModal
                    GoTo lanjut
                    Exit Sub
                Case vbNo
                    If MsgBox("Yakin data disimpan!!! perubahan data tidak terekam histori datanya!!!", vbOKCancel, "Konfirmasi") = vbCancel Then Exit Sub
                Case vbCancel
                    Exit Sub
            End Select
        End If
        Me.MousePointer = vbHourglass
        SQL = " update datakaryawan set nama = '" & Trim(Replace(UCase(txtNama.text), "'", "''")) & "',[tempat lahir] = '" & Trim(UCase(txtTempatLahir.text)) & "',[tanggal lahir]= '" & Format(DTTgllahir.value, "MM/DD/yyyy") & "',[jenis kelamin]='" & IIf(OptJK(0).value = True, "L", "P") & "',[agama] = '" & Trim(cmbAgama.value) & "',[status perkawinan] = '" & Trim(cmbStatusKawin.value) & "' ,[golongan darah] = '" & Trim(cmbGolDarah.value) & "',kewarganegaraan = '" & Trim(UCase(txtKewarganegaraan.text)) & "',[alamat tinggal] = ' " & Replace(Trim(UCase(txtAlmtTgal.text)), vbNewLine, "", , , vbTextCompare) & " ',[kota tinggal] = '" & Trim(UCase(txtKotaTgal.text)) & "',[kode pos] = '" & Trim(txtKdPos.text) & "',[alamat ktp] = ' " & Replace(Trim(UCase(txtAlmtKTP.text)), vbNewLine, "", , , vbTextCompare) & " ' "
        SQL = SQL & " ,[kota ktp] = '" & Trim(UCase(txtKotaKtp.text)) & "',[telp rumah] = '" & Trim(txtTelpRmh.text) & "',hp = '" & Trim(txtHp.text) & "',[jenjang pendidikan] = '" & Trim(cmbjenjang(0).value) & "',[pendidikan akhir] = '" & Trim(cmbpendidikanAkhir(0).value) & "',jurusan = '" & Trim(cmbJurusan(0).value) & "', gelar = '" & Trim(cmbGelar.value) & "',golongan = '" & Trim(cmbGolongan.value) & "',pangkat = '" & Trim(cmbpangkat.Columns(0).value) & "',divisi = '" & cmbDept.Columns(0).value & "',dealer = '" & Trim(cmbdealer.Columns(0).value) & "',homebase = '" & Trim(UCase(txtHomeBase.text)) & "',posisi = '" & Trim(cmbPosisi.Columns(0).value) & "',[tanggal masuk] = '" & Format(DTTglmasuk.value, "MM/DD/yyyy") & "' "
        SQL = SQL & " ,[tanggal keluar] = " & IIf(DTTglKeluar.ValueIsNull = True, "Null", "'" & Format(DTTglKeluar.value, "MM/DD/yyyy") & "'") & ",[no jamsostek] = '" & Trim(txtJamsostek.text) & "',[no polis asuransi] = '" & Trim(txtNopolisAsuransi.text) & "',catatan = ' " & Replace(Trim(UCase(txtcatatan.text)), vbNewLine, "", , , vbTextCompare) & " ',voptr = '" & strUserID & "',[last updated] = getdate(),divisiho = '" & IIf(PakaiDivisi = True, cmbDivisi.Columns(0).value, cmbDept.Columns(2).value) & "',status = '" & Trim(Replace(cmbstatus.value, "'", "''")) & "' ,nik_atasan = '" & Trim(atasan) & "',nama_atasan = '" & Trim(Replace(txtatasan.text, "'", "''")) & "',noktp = '" & Trim(UCase(txtNoKTP.text)) & "',nokk = '" & Trim(UCase(TxtNoKK.text)) & "',npwp='" & Trim(UCase(TxtNPWP.text)) & "',nobpjskesehatan = '" & Trim(UCase(txtNoBpjsSehat.text)) & "',kodeposktp = '" & Trim(UCase(txtKodePosKTP.text)) & "',email = '" & Trim(txtEmail.text) & "' where nik = '" & Trim(txtNik.text) & "'"
        If DbCon.State = 0 Then DbCon.Open StrCon
        DbCon.Execute SQL
        
        'simpan gambar
        If cdgopen.FileName <> "" Then
         
         Me.MousePointer = vbHourglass
         SQL = "select * from e_fotokry where nik = '" & Trim(txtNik.text) & "'"
         Set rsFoto = New ADODB.Recordset
         rsFoto.Open SQL, DbCon, adOpenKeyset, adLockOptimistic
         With mStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile cdgopen.FileName
            If Not rsFoto.EOF Then
                rsFoto("foto").value = .Read
                rsFoto.Update
            Else
                rsFoto.AddNew
                rsFoto("nik").value = Trim(txtNik.text)
                rsFoto("foto").value = .Read
                rsFoto("last_update").value = GetDate
                rsFoto("user_update").value = UserID
                rsFoto.Update
            End If
         End With
         
         Set mStream = Nothing
         Set rsFoto = Nothing
        End If
        'update reminder
        SQL = "select id_reminder,nik,dealer,kategori from reminder_h where nik = '" & Trim(txtNik.text) & "' and kategori in ('PENSIUN','MAGANG','KONTRAK')"
        Set rsFoto = New ADODB.Recordset
        rsFoto.Open SQL, DbCon, adOpenForwardOnly, adLockReadOnly
        If rsFoto.RecordCount > 0 Then
            rsFoto.MoveFirst
            Do While Not rsFoto.EOF
                If Trim(rsFoto!kateGori) = "PENSIUN" Then
                    If DTTglKeluar <> "" And Trim(txtcatatan.text) = "PENSIUN" Then
                        DbCon.Execute "update reminder_h set tgleksekusi = getdate(),userupdate = '" & strUserID & "', tglupdate = getdate() where id_reminder = " & rsFoto!Id_reminder & ""
                    End If
                ElseIf Trim(rsFoto!kateGori) = "MAGANG" Then
                    If InStr(1, Trim(UCase(cmbstatus.value)), Trim(UCase(rsKry!Status))) = 0 Then
                        DbCon.Execute "update reminder_h set tgleksekusi = getdate(),userupdate = '" & strUserID & "', tglupdate = getdate() where id_reminder = " & rsFoto!Id_reminder & ""
                    End If
                ElseIf Trim(rsFoto!kateGori) = "KONTRAK" Then
                    If InStr(1, Trim(UCase(cmbstatus.value)), Trim(UCase(rsKry!Status))) = 0 Then
                        DbCon.Execute "update reminder_h set tgleksekusi = getdate(),userupdate = '" & strUserID & "', tglupdate = getdate() where id_reminder = " & rsFoto!Id_reminder & ""
                    End If
                End If
                rsFoto.MoveNext
            Loop
        End If
        rsFoto.Close
        
        Call log_user(strUserID, DealerClaim, "U", "Ubah Kry " & Trim(txtNik.text))
    End If
    MsgBox "Simpan data berhasil!!!", vbInformation, "Info"
lanjut:
    Me.MousePointer = vbDefault
    clear_set
    Unload Me
    Exit Sub
Salah:
    MsgBox err.Description, vbInformation, "Failure"
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    Dim persen_lebar As Double
    Dim persen_tinggi As Double
    On Error Resume Next
'    If Me.Width > 13395 Or Me.Height > 8640 Then
'        persen_lebar = ((Me.Width - 13395) / 13395)
'        persen_tinggi = ((Me.Height - 9090) / 9090)
'        tabkry.Height = (8295 * persen_tinggi) + 8295
'        picone.Height = (6735 * persen_tinggi) + 6735
'        VScroll1.Height = (6855 * persen_tinggi) + 6855
'        framut.Height = (5775 * persen_tinggi) + 5775
'        frakel.Height = (6855 * persen_tinggi) + 6855
'        GridMutasi.Height = (4455 * persen_tinggi) + 4455
'        TDBGridKeluarga.Height = (2175 * persen_tinggi) + 2175
'    End If
    tabkry.Height = Me.Height - 500
    picone.Height = tabkry.Height - 1575
    VScroll1.Height = picone.Height
    framut.Height = picone.Height
    GridMutasi.Height = framut.Height - 1200
    frakel.Height = picone.Height
    TDBGridKeluarga.Height = frakel.Height - fraisikel.Height - 1500
    FraPendidikan.Height = picone.Height
    GridPendidikan.Height = FraPendidikan.Height - FraisiPendidikan.Height - 1500
    
    tabkry.Move (Me.Width - tabkry.Width) / 2, tabkry.Top
    cmdSave.Move picone.Left, cmdSave.Top
    cmdclose.Move picone.Left + cmdclose.Width, cmdclose.Top
End Sub
Private Function cek_nik(Nik As String) As Boolean
    Dim SQL As String
    Dim rsnik As New ADODB.Recordset
    SQL = "select count(*) jml from datakaryawan where nik = '" & Trim(Nik) & "'"
    If DbCon.State = 0 Then DbCon.Open StrCon
    Set rsnik = New ADODB.Recordset
    rsnik.Open SQL, DbCon
    cek_nik = False
    If Not rsnik.EOF Then
        If rsnik!Jml > 0 Then
            cek_nik = True
        End If
    End If
    rsnik.Close
    Set rsnik = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    clear_set
    If Dir(App.Path & "\Temp.bmp") <> "" Then
        Kill (App.Path & "\Temp.bmp")
    End If
End Sub

Private Sub GridPendidikan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If GridPendidikan.ApproxCount = 0 Then Exit Sub
    
    If rsDTPend.RecordCount > 0 Then
        CboTahun(0).Clear
        CboTahun(1).Clear
        CboTahun(0).AddItem Nul(rsDTPend!thnmasuk)
        CboTahun(0).ListIndex = 0
        CboTahun(1).AddItem Nul(rsDTPend!thnlulus)
        CboTahun(1).ListIndex = 0
        cmbjenjang(1) = Nul(rsDTPend!jenjang)
        cmbjenjang(1).value = Nul(rsDTPend!jenjang)
        cmbpendidikanAkhir(1) = Nul(rsDTPend!pendidikan)
        cmbpendidikanAkhir(1).value = Nul(rsDTPend!pendidikan)
        cmbJurusan(1) = Nul(rsDTPend!jurusan)
        cmbJurusan(1).value = Nul(rsDTPend!jurusan)
        txtalmamater = Nul(rsDTPend!almamater)
        Chklulus.value = IIf(Nul(rsDTPend!lulus) = "Y", 1, 0)
        txtpredikat = Nul(rsDTPend!predikat)
        txtketPendidikan = Nul(rsDTPend!keterangan)
        txtidPendidikan = Nul(rsDTPend!idpendkry)
    End If
End Sub

Private Sub lv_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    Dim d As Integer
    
    txtatasan.text = ""

    If Item.SubItems(2) = "All" Then
        For i = 1 To lv(Index).ListItems.Count
            If Item.Checked = True Then
                If UCase(lv(Index).ListItems(i).SubItems(2)) <> "ALL" Then
                    lv(Index).ListItems(i).Checked = True
                    txtatasan.text = txtatasan.text & IIf(Trim(txtatasan.text) = "", "", ",") & lv(Index).ListItems(i).SubItems(2)
                End If
            Else
                lv(Index).ListItems(i).Checked = False
            End If
        Next i
    Else
        For i = 1 To lv(Index).ListItems.Count
            If UCase(lv(Index).ListItems(i).SubItems(2)) = "ALL" Then
                lv(Index).ListItems(i).Checked = False
            Else
                If lv(Index).ListItems(i).Checked = True Then
                     txtatasan.text = txtatasan.text & IIf(Trim(txtatasan.text) = "", "", ",") & lv(Index).ListItems(i).SubItems(2)
                End If
            End If
        Next i
    End If

'    If Item.Checked = True Then
'        If Item.SubItems(2) = "All" Then
'            txtatasan.Text = ""
'            For i = 1 To lv(Index).ListItems.Count
'                If lv(Index).ListItems(i).Checked = True Then
'                    lv(Index).ListItems(i).Checked = False
'                    txtatasan.Text = ""
'                Else
'                    lv(Index).ListItems(i).Checked = True
'                    txtatasan.Text = txtatasan.Text & IIf(Trim(txtatasan.Text) = "", "", ",") & lv(Index).ListItems(i).SubItems(2)
'                End If
'            Next i
'        Else
'            txtatasan.Text = txtatasan.Text & IIf(Trim(txtatasan.Text) = "", "", ",") & Item.SubItems(2)
'        End If
'    Else
'           If txtatasan.Text <> "" Then
'                txtatasan.Text = Replace(txtatasan.Text, Item.SubItems(2), "")
'                txtatasan.Text = Replace(txtatasan.Text, ",,", ",")
'                If Left(txtatasan.Text, 1) = "," Then
'                    txtatasan.Text = Right(Trim(txtatasan.Text), Val(Len(Trim(txtatasan.Text)) - 1))
'                End If
'                If Right(txtatasan.Text, 1) = "," Then
'                    txtatasan.Text = Left(Trim(txtatasan.Text), Val(Len(Trim(txtatasan.Text)) - 1))
'                End If
'           End If
'    End If
End Sub

Private Sub lv_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lv(0).Height = 3500 Then
            lv(0).Height = 0
        End If
    End If
End Sub

Private Sub lv_LostFocus(Index As Integer)
    If lv(0).Height = 3500 Then
       lv(0).Height = 0
    End If
End Sub

Private Sub TDBGridKeluarga_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If TDBGridKeluarga.ApproxCount = 0 Then Exit Sub
    txtnamakel.text = Trim(rsDtkel!Nama)
    txtt4lahirkel.text = Trim(rsDtkel![tempat lahir])
    DttglLhrKel.value = Trim(rsDtkel![tanggal lahir])
    cmbagamakel.text = Trim(rsDtkel!Agama)
    cmbagamakel.Columns(0).value = Trim(rsDtkel!Agama)
    If Trim(rsDtkel![jenis kelamin]) = "L" Then
        OptJK(2).value = True
    ElseIf Trim(rsDtkel![jenis kelamin]) = "P" Then
        OptJK(3).value = True
    Else
        OptJK(2).value = False
        OptJK(3).value = False
    End If
    
    If Nul(rsDtkel!grup) = "" Then
        cmbhubungan.ListIndex = -1
    Else
        cmbhubungan.Clear
        cmbhubungan.AddItem Nul(rsDtkel!grup)
        cmbhubungan.ListIndex = 0
    End If
    
    cmbgoldarahkel.Columns(0).value = Trim(rsDtkel![golongan darah])
    cmbgoldarahkel.text = Trim(rsDtkel![golongan darah])
    cmbHubkel.Columns(0).value = IIf(IsNull(rsDtkel!idhub) = True, Trim(rsDtkel![hubkel]), rsDtkel!idhub)
    cmbHubkel.Columns(1).value = Nul(rsDtkel!norecklg)
    cmbHubkel.text = Trim(rsDtkel![hubkel])
    cmbstskawinkel.Columns(0).value = Trim(rsDtkel![status perkawinan])
    cmbstskawinkel.text = Trim(rsDtkel![status perkawinan])
    cmbpendidikanakhirkel.Columns(0).value = Trim(rsDtkel![pendidikan akhir])
    cmbpendidikanakhirkel.text = Trim(rsDtkel![pendidikan akhir])
    txtkerjaan.text = Trim(rsDtkel![Pekerjaan])
    txtalmtKel.text = Trim(rsDtkel!Alamat)
    txtketerangankel.text = Trim(rsDtkel!keterangan)
    
    If Trim(rsDtkel![tertanggung]) = "Y" Then
        chktanggung.value = 1
    Else
        chktanggung.value = 0
    End If
End Sub

Private Sub txtNik_LostFocus()
    If cek_nik(txtNik.text) = True Then
        MsgBox "Nik Sudah ada,pakai nik yang lain!!!", vbInformation, "Info": Exit Sub
    End If
End Sub
Private Sub VScroll1_Change()
    fradd.Top = VScroll1.value
End Sub
Public Sub bersih_isi()
    Dim tglserver As Date
    tglserver = DateSvr()
    txtNik.text = ""
    txtNama.text = ""
    txtTempatLahir.text = ""
    DTTgllahir.value = tglserver
    OptJK(0).value = False
    OptJK(1).value = False
    cmbAgama.text = ""
    cmbAgama.value = ""
    cmbStatusKawin.text = ""
    cmbStatusKawin.value = ""
    cmbGolDarah.text = ""
    cmbGolDarah.value = ""
    txtKewarganegaraan.text = ""
    txtAlmtTgal.text = ""
    txtKotaTgal.text = ""
    txtKdPos.text = ""
    txtAlmtKTP.text = ""
    txtKotaKtp.text = ""
    txtTelpRmh.text = ""
    txtHp.text = ""
    cmbjenjang(0).text = ""
    cmbjenjang(0).value = ""
    cmbJurusan(0).text = ""
    cmbJurusan(0).value = ""
    cmbpendidikanAkhir(0).text = ""
    cmbpendidikanAkhir(0).value = ""
    cmbGelar.text = ""
    cmbGelar.value = ""
    cmbdealer.text = ""
    cmbdealer.value = ""
    cmbDivisi.text = ""
    cmbDivisi.value = ""
    cmbGolongan.text = ""
    cmbGolongan.value = ""
    txtHomeBase.text = ""
    cmbpangkat.text = ""
    cmbpangkat.value = ""
    cmbstatus.text = ""
    cmbstatus.value = ""
    cmbPosisi.text = ""
    cmbPosisi.value = ""
    DTTglmasuk.value = tglserver
    DTTglKeluar.value = Null
    txtJamsostek.text = ""
    txtNopolisAsuransi.text = ""
    txtcatatan.text = ""
    cmbDivisi.Enabled = IIf(Baru = "Y", True, flag_ho)
    txtNoKTP.text = ""
    TxtNoKK.text = ""
    txtNoBpjsSehat.text = ""
    txtKodePosKTP.text = ""
    TxtNPWP.text = ""
    txtEmail.text = ""
End Sub
Public Sub isi_data(Nik As String)
On Error GoTo Salah
    Dim SQL As String
    
    Dim streamx As New ADODB.Stream
    Dim rsAts As ADODB.Recordset
    Dim i As Integer
    SQL = "select a.nama,a.[tempat lahir],a.[tanggal lahir],a.[jenis kelamin],a.agama,a.[status perkawinan],isnull(a.[golongan darah],'')[golongan darah],isnull(a.kewarganegaraan,'')kewarganegaraan, " & _
            " isnull(a.[alamat tinggal],'')[alamat tinggal],isnull(a.[kota tinggal],'')[kota tinggal],isnull(a.[kode pos],'')[kode pos],isnull(a.[alamat ktp],'')[alamat ktp],isnull(a.[kota ktp],'')[kota ktp], " & _
            " isnull(a.[telp rumah],'')[telp rumah],isnull(a.[hp],'')[hp],isnull(a.[jenjang pendidikan],'')[jenjang pendidikan],isnull(a.[pendidikan akhir],'')[pendidikan akhir],isnull(a.[jurusan],'')[jurusan], " & _
            " isnull(a.[gelar],'')[gelar],isnull(a.[golongan],'')[golongan],isnull(a.[pangkat],'')[idpangkat],isnull(g.[pangkat],'')[namapangkat],isnull(a.[divisiho],'')[iddivisiho],isnull(c.[divisi],'')[divisiho],isnull(a.[divisi],'')[iddivisi],isnull(d.[departemen],'')[divisi],isnull(a.[dealer],'')[dealer],isnull(b.[namadealer],'') namadealer, " & _
            " isnull(a.[homebase],'')[homebase],isnull(a.[posisi],'')[idposisi],isnull(f.[posisi],'')[namaposisi],a.[tanggal masuk],a.[tanggal keluar],isnull(a.[no jamsostek],'')[no jamsostek], " & _
            " isnull(a.[no polis asuransi],'')[no polis asuransi],isnull(a.[status],'')[status],isnull(a.catatan,'') catatan,isnull(b.flag_ho,0) flag_ho,i.foto,isnull(i.nik,'') as idfoto,g.[level],isnull(k.nik,isnull(a.nik_atasan,'')) nik_atasan,isnull(a.noktp,'')noktp,isnull(a.nokk,'')nokk,isnull(a.npwp,'')npwp,isnull(a.nobpjskesehatan,'')nobpjskesehatan,isnull(a.kodeposktp,'')kodeposktp,isnull(a.email,'')email,isnull(j.gradeakhir,'') gradeakhir "
    SQL = SQL & "From datakaryawan a " & _
            " left join tdealer b on a.dealer = b.dealer " & _
            " left join E_Divisi c on a.divisiho = c.id_divisi " & _
            " left join E_Dept d on a.divisi = d.id_dept  and d.id_divisi = a.divisiho and d.flag_ho = b.flag_ho" & _
            " left join e_pangkat g on a.pangkat = g.id_pangkat " & _
            " left join e_posisi f on a.posisi = f.id_posisi  and f.id_pangkat = a.pangkat " & _
            " left join e_fotokry i on a.nik = i.nik left join " & dbGrading & "..tblmaster j on a.nik = j.nik and j.tanggalkeluar is null " & _
            " left join " & dbGrading & "..TBLMasterMan  k on j.Spv = k.idNama " & _
            " where a.nik = '" & Trim(Nik) & "' "
    
    If DbCon.State = 0 Then DbCon.Open StrCon
    Set rsKry = New ADODB.Recordset
    rsKry.Open SQL, DbCon
    If Not rsKry.EOF Then
        txtNama.text = Trim(rsKry!Nama): txtTempatLahir.text = Trim(rsKry![tempat lahir]): DTTgllahir.value = Trim(rsKry![tanggal lahir])
        If UCase(Trim(rsKry![jenis kelamin])) = "L" Then
            OptJK(0).value = True
        ElseIf UCase(Trim(rsKry![jenis kelamin])) = "P" Then
            OptJK(1).value = True
        Else
            OptJK(0).value = False: OptJK(1).value = False
        End If
        flag_ho = rsKry!flag_ho
        cmbAgama.text = Trim(rsKry!Agama): cmbAgama.value = Trim(rsKry!Agama)
        cmbStatusKawin.text = Trim(rsKry![status perkawinan]): cmbStatusKawin.value = Trim(rsKry![status perkawinan])
        cmbGolDarah.text = Trim(rsKry![golongan darah]): cmbGolDarah.value = Trim(rsKry![golongan darah]): txtKewarganegaraan.text = Trim(rsKry![kewarganegaraan]): txtAlmtTgal.text = Trim(rsKry![alamat tinggal]): txtKotaTgal.text = Trim(rsKry![kota tinggal]): txtKdPos.text = Trim(rsKry![kode pos]): txtAlmtKTP.text = Trim(rsKry![alamat ktp]): txtKotaKtp.text = Trim(rsKry![kota ktp])
        txtTelpRmh.text = Trim(rsKry![telp rumah]): txtHp.text = Trim(rsKry![HP]): cmbjenjang(0).text = Trim(rsKry![jenjang pendidikan]): cmbjenjang(0).value = Trim(rsKry![jenjang pendidikan]): cmbJurusan(0).text = Trim(rsKry![jurusan]): cmbJurusan(0).value = Trim(rsKry![jurusan]): cmbpendidikanAkhir(0).text = Trim(rsKry![pendidikan akhir]): cmbpendidikanAkhir(0).value = Trim(rsKry![pendidikan akhir]): cmbGelar.text = Trim(rsKry![gelar]): cmbGelar.value = Trim(rsKry![gelar])
        cmbdealer.Columns(0).value = Trim(rsKry![dealer])
        cmbdealer.text = Trim(rsKry!namaDealer): txtgrade.text = Trim(rsKry!gradeakhir)
        txtNoKTP.text = Trim(rsKry!noKTP): TxtNoKK.text = Trim(rsKry!nokk): TxtNPWP.text = Trim(rsKry!npwp): txtNoBpjsSehat.text = Trim(rsKry!nobpjskesehatan): txtKodePosKTP.text = Trim(rsKry!kodeposktp): txtEmail.text = Trim(rsKry!email)
        If flag_ho = True Then
            If Trim(rsKry![iddivisi]) <> "KAIZEN" And Trim(rsKry![iddivisi]) <> "INTERNAL AUDIT" Then
                cmbDivisi.text = Trim(rsKry![divisiho])
                cmbDivisi.Columns(0).value = Trim(rsKry![iddivisiho])
                lbl(23).Caption = "Divisi (*) :"
                lbl(32).Caption = "Departemen (*) :"
                lbl(23).Visible = True
                cmbDivisi.Visible = True
                lbl(20).Left = 6240
                cmbGolongan.Left = 8160
            Else
                lbl(23).Caption = "Divisi :"
                cmbDivisi.Enabled = False
                cmbDivisi.Columns(0).value = Trim(rsKry![iddivisiho])
                lbl(32).Caption = "Section (*) :"
                lbl(23).Visible = False
                cmbDivisi.Visible = False
                lbl(20).Left = 360
                cmbGolongan.Left = 2280
            End If
        Else
            lbl(23).Caption = "Divisi :"
            cmbDivisi.Enabled = False
            lbl(32).Caption = "Section (*) :"
            lbl(23).Visible = False
            cmbDivisi.Visible = False
            cmbGolongan.Left = 2280
            lbl(20).Left = 360
        End If
        id_dealer = Trim(rsKry![dealer])
        id_divisi = Trim(rsKry!iddivisiho)
        id_dept = Trim(rsKry!iddivisi)
        id_pangkat = Trim(rsKry![idpangkat])
        id_posisi = Trim(rsKry![idposisi])
        nik_atasan = Trim(rsKry!nik_atasan)
        cmbDept.text = Trim(rsKry![divisi])
        cmbDept.Columns(0).value = Trim(rsKry![iddivisi])
        cmbDept.Columns(2).value = Trim(rsKry!iddivisiho)
        cmbGolongan.text = Trim(rsKry![golongan]): cmbGolongan.value = Trim(rsKry![golongan])
        txtHomeBase.text = Trim(rsKry![homebase])
        cmbpangkat.text = Trim(rsKry![namapangkat])
        cmbpangkat.Columns(0).value = Trim(rsKry![idpangkat])
        cmbpangkat.Columns(2).value = Trim(rsKry![level])
        cmbstatus.text = Trim(rsKry![Status]): cmbstatus.value = Trim(rsKry![Status])
'        cmbatasan.Text = Trim(rsKry!namaatasan)
'        cmbatasan.Columns(0).Value = Trim(rsKry!nik_atasan)
        Call isi_atasan
        SQL = "select nik,nama from datakaryawan where nik in('" & Replace(Trim(rsKry!nik_atasan), ",", "','") & "')"
        Set rsAts = New ADODB.Recordset
        rsAts.Open SQL, DbCon
        Do While Not rsAts.EOF
            For i = 1 To lv(0).ListItems.Count
                If lv(0).ListItems(i).SubItems(1) = Trim(rsAts!Nik) Then
                    lv(0).ListItems(i).Checked = True
                End If
            Next i
            txtatasan.text = txtatasan.text & IIf(txtatasan.text = "", "", ",") & Trim(rsAts!Nama)
            rsAts.MoveNext
        Loop
        rsAts.Close
        
        cmbPosisi.text = Trim(rsKry![namaposisi])
        cmbPosisi.Columns(0).value = Trim(rsKry![idposisi])
        DTTglmasuk.value = Trim(rsKry![tanggal masuk]): DTTglKeluar.value = Trim(rsKry![tanggal keluar]): txtJamsostek.text = Trim(rsKry![no jamsostek]): txtNopolisAsuransi.text = Trim(rsKry![no polis asuransi]): txtcatatan.text = Trim(rsKry!catatan)
        'foto
        If Trim(rsKry!idfoto) <> "" And IsNull(rsKry!foto) = False Then
            Set streamx = New ADODB.Stream
            With streamx
                .Type = adTypeBinary
                .Open
                .Write rsKry.Fields("foto").value
                .SaveToFile App.Path & "\Temp.bmp", adSaveCreateOverWrite
                imgFoto.Picture = LoadPicture(App.Path & "\Temp.bmp")
            End With
        End If
    Else
        MsgBox "Data tidak ditemukan!!!", vbInformation, "Info"
    End If
'    rsKry.Close
'    Set rsKry = Nothing
    Exit Sub
Salah:
    MsgBox err.Description, vbInformation, "Failure"
End Sub
Private Sub isi_agama()
Dim SQL As String
Dim rsFind As New ADODB.Recordset
    SQL = "SELECT agama id,agama from tagama"
    Set rsFind = DbCon.Execute(SQL)
    cmbAgama.Reset
    cmbagamakel.Reset
    If rsFind.BOF Then Exit Sub
    AdoAgama.RecordSource = SQL
    With cmbAgama
            .DataSourceList = AdoAgama
            .DataFieldList = "agama"
            .DataFieldToDisplay = "agama"
            .Columns(0).Width = 1000
    End With
    cmbAgama.refresh
    With cmbagamakel
            .DataSourceList = AdoAgama
            .DataFieldList = "id"
            .DataFieldToDisplay = "agama"
            .Columns(0).Width = 1000
    End With
    cmbagamakel.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_kawin()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT [satus perkawinan] as status from statusperkawinan"
    Set rsFind = DbCon.Execute(SQL)
    cmbStatusKawin.Reset
    cmbstskawinkel.Reset
    If rsFind.BOF Then Exit Sub
    AdoStsKawin.RecordSource = SQL
    With cmbStatusKawin
            .DataSourceList = AdoStsKawin
            .DataFieldList = "status"
            .DataFieldToDisplay = "status"
            .Columns(0).Width = 5000
    End With
    cmbStatusKawin.refresh
    With cmbstskawinkel
            .DataSourceList = AdoStsKawin
            .DataFieldList = "status"
            .DataFieldToDisplay = "status"
            .Columns(0).Width = 5000
    End With
    cmbstskawinkel.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub

Private Sub isi_jenjang(Index As Integer)
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT distinct [jenjang pendidikan] jenjang from t_jenjang"
    Set rsFind = DbCon.Execute(SQL)
    cmbjenjang(Index).Reset
    If rsFind.BOF Then Exit Sub
    AdoJenjang(Index).RecordSource = SQL
    AdoJenjang(Index).refresh
    With cmbjenjang(Index)
            .DataSourceList = AdoJenjang(Index)
            .DataFieldList = "jenjang"
            .DataFieldToDisplay = "jenjang"
            .Columns(0).Width = 1000
    End With
    cmbjenjang(Index).refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_pendAkhir(Index As Integer)
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    If cmbjenjang(Index) = "" Then Exit Sub
    SQL = "SELECT pendidikan from t_pendakhir where jenjang = '" & cmbjenjang(Index).Columns(0).value & "' "
    Set rsFind = DbCon.Execute(SQL)
    cmbpendidikanAkhir(Index).Reset
    If rsFind.BOF Then Exit Sub
    AdoPendidikanAkhir(Index).RecordSource = SQL
    AdoPendidikanAkhir(Index).refresh
    With cmbpendidikanAkhir(Index)
            .DataSourceList = AdoPendidikanAkhir(Index)
            .DataFieldList = "pendidikan"
            .DataFieldToDisplay = "pendidikan"
            .Columns(0).Width = 1000
    End With
    cmbpendidikanAkhir(Index).refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_pendAkhirKeluarga()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT pendidikan from t_pendakhir  "
    Set rsFind = DbCon.Execute(SQL)
    cmbpendidikanakhirkel.Reset
    If rsFind.BOF Then Exit Sub
    AdoPendidikanAkhir(0).RecordSource = SQL
    With cmbpendidikanakhirkel
            .DataSourceList = AdoPendidikanAkhir(0)
            .DataFieldList = "pendidikan"
            .DataFieldToDisplay = "pendidikan"
            .Columns(0).Width = 1000
    End With
    cmbpendidikanakhirkel.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_gelar()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT gelar from tgelar"
    Set rsFind = DbCon.Execute(SQL)
    cmbGelar.Reset
    If rsFind.BOF Then Exit Sub
    AdoGelar.RecordSource = SQL
    With cmbGelar
            .DataSourceList = AdoGelar
            .DataFieldList = "gelar"
            .DataFieldToDisplay = "gelar"
            .Columns(0).Width = 1000
    End With
    cmbGelar.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_cabang()
Dim rsFind As New ADODB.Recordset
Dim aDealer As String
Dim SQL As String
    If strLevelUser = "Cabang" Then
        aDealer = " where dealer = '" & Trim(DealerClaim) & "'"
    Else
        aDealer = ""
    End If
    AdoCbgdetail.RecordSource = ""
    SQL = "SELECT dealer,namadealer,flag_ho from tdealer " & aDealer & ""
    Set rsFind = DbCon.Execute(SQL)
    cmbdealer.Reset
    If rsFind.BOF Then Exit Sub
    AdoCbgdetail.RecordSource = SQL
    With cmbdealer
            .DataSourceList = AdoCbgdetail
            .DataFieldList = "kodedealer"
            .DataFieldToDisplay = "namadealer"
            .Columns(0).Width = 700
            .Columns(1).Width = 4000
            .Columns(2).Visible = False
    End With
    cmbdealer.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub

Private Sub isi_divisi()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT id_divisi,divisi from E_Divisi"
    Set rsFind = DbCon.Execute(SQL)
    cmbDivisi.Reset
    If rsFind.BOF Then Exit Sub
    AdoDivisi.RecordSource = SQL
    With cmbDivisi
            .DataSourceList = AdoDivisi
            .DataFieldList = "id_divisi"
            .DataFieldToDisplay = "divisi"
            .Columns(0).Visible = False
            .Columns(1).Width = 5000
    End With
    cmbDivisi.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_golongan()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT golongan from golongan"
    Set rsFind = DbCon.Execute(SQL)
    cmbGolongan.Reset
    If rsFind.BOF Then Exit Sub
    AdoGolongan.RecordSource = SQL
    With cmbGolongan
            .DataSourceList = AdoGolongan
            .DataFieldList = "golongan"
            .DataFieldToDisplay = "golongan"
            .Columns(0).Width = 1000
    End With
    cmbGolongan.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_pangkat()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "SELECT id_pangkat,pangkat,[level] from e_pangkat "
    Set rsFind = DbCon.Execute(SQL)
    cmbpangkat.Reset
    If rsFind.BOF Then Exit Sub
    AdoPangkat.RecordSource = SQL
    With cmbpangkat
            .DataSourceList = AdoPangkat
            .DataFieldList = "id_pangkat"
            .DataFieldToDisplay = "pangkat"
            .Columns(0).Visible = False
            .Columns(1).Width = 2000
    End With
    cmbpangkat.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_hubkel()
Dim rsFind As New ADODB.Recordset
Dim SQL As String
    SQL = "select idhub id,norec,[hubungan keluarga] hubungan,case isnull(statuskel,0) when 1 then 'Kandung' when 2 then 'Tiri' when 3 then '-' else '' end as statuskel From hubungankeluarga where grup = '" & UCase(Trim(cmbhubungan.text)) & "' "
    Set rsFind = DbCon.Execute(SQL)
    cmbHubkel.Reset
    
    If rsFind.BOF Then Exit Sub
    AdoHubKel.RecordSource = SQL
    AdoHubKel.refresh
    With cmbHubkel
            .DataSourceList = AdoHubKel
            .DataFieldList = "id"
            .DataFieldToDisplay = "hubungan"
            .Columns(0).Visible = False
            .Columns(1).Width = 700
            .Columns(2).Width = 3000
            .Columns(3).Width = 2000
    End With
    cmbHubkel.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub
Private Sub isi_dept()
Dim SQL As String
Dim rsFind As New ADODB.Recordset
AdoDept.RecordSource = ""
    If flag_ho = True Then
        SQL = "SELECT id_dept,departemen,id_divisi from E_Dept where id_divisi = '" & Trim(cmbDivisi.Columns(0).value) & "' and flag_ho = 1"
    Else
        SQL = "SELECT id_dept,departemen,id_divisi from E_Dept where flag_ho = 0"
    End If
    Set rsFind = DbCon.Execute(SQL)
    cmbDept.Reset
    If rsFind.BOF Then Exit Sub
    AdoDept.RecordSource = SQL
    AdoDept.refresh
    With cmbDept
            .DataSourceList = AdoDept
            .DataFieldList = "id_dept"
            .DataFieldToDisplay = "departemen"
            .Columns(0).Visible = False
            .Columns(1).Width = 6000
            .Columns(2).Visible = False
    End With
    cmbDept.refresh
    rsFind.Close
    Set rsFind = Nothing
End Sub

Private Sub tombol_kel(add As Boolean, Edit As Boolean, del As Boolean, save As Boolean, tutup As Boolean, refresh As Boolean, Cancel As Boolean)
    cmdaddkel.Enabled = add
    cmdeditkel.Enabled = Edit
    cmddelkel.Enabled = del
    cmdsavekel.Enabled = save
    cmdclosekel.Enabled = tutup
    cmdrefkel.Enabled = refresh
    cmdcancel.Enabled = Cancel
End Sub
Private Sub bersih_kel()
    txtnamakel.text = ""
    txtt4lahirkel.text = ""
    DttglLhrKel.value = GetDate
    cmbagamakel.Columns(0).value = ""
    cmbagamakel.text = ""
    OptJK(2).value = False
    OptJK(3).value = False
    cmbgoldarahkel.Columns(0).value = ""
    cmbgoldarahkel.text = ""
    cmbHubkel.Columns(0).value = ""
    cmbHubkel.text = ""
    cmbstskawinkel.Columns(0).value = ""
    cmbstskawinkel.text = ""
    cmbpendidikanakhirkel.Columns(0).value = ""
    cmbpendidikanakhirkel.text = ""
    txtkerjaan.text = ""
    txtalmtKel.text = ""
    txtketerangankel.text = ""
End Sub



