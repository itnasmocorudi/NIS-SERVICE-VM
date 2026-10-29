VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOMasterEmailTL 
   Caption         =   ".: SETTING MASTER EMAIL TL DEALER :."
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16320
   Begin VB.Frame frameSettingEmail 
      Caption         =   "Setting Email"
      Height          =   7575
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   12255
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   9735
         Begin VB.TextBox txtJudulEmail 
            Height          =   375
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   47
            Top             =   650
            Width           =   4575
         End
         Begin ucXPButton.XPButton cmdNamaDealerJudul 
            Height          =   495
            Left            =   2160
            TabIndex        =   49
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BTYPE           =   6
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmHOMasterEmailTL.frx":0000
            PICN            =   "frmHOMasterEmailTL.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdNopolJudul 
            Height          =   495
            Left            =   3480
            TabIndex        =   50
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BTYPE           =   6
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmHOMasterEmailTL.frx":34AE
            PICN            =   "frmHOMasterEmailTL.frx":34CA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdModelJudul 
            Height          =   495
            Left            =   4800
            TabIndex        =   51
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            BTYPE           =   6
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmHOMasterEmailTL.frx":6856
            PICN            =   "frmHOMasterEmailTL.frx":6872
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Judul Email : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   660
            Width           =   2145
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   9735
         Begin VB.TextBox txtEmailATL2 
            Height          =   375
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   42
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txtEmailATL1 
            Height          =   375
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   41
            Top             =   120
            Width           =   4575
         End
         Begin VB.TextBox txtPengirim 
            Height          =   375
            Left            =   2160
            MaxLength       =   100
            TabIndex        =   40
            Top             =   840
            Width           =   4575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email ATL 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   45
            Top             =   540
            Width           =   2145
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email ATL 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pengirim : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   900
            Width           =   2145
         End
      End
      Begin VB.TextBox txtBodyEmail 
         Height          =   2655
         Left            =   2160
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   3240
         Width           =   7695
      End
      Begin ucXPButton.XPButton XPButton1 
         Height          =   495
         Left            =   8520
         TabIndex        =   17
         Top             =   6000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":9ABB
         PICN            =   "frmHOMasterEmailTL.frx":9AD7
         PICH            =   "frmHOMasterEmailTL.frx":100D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpanEmail 
         Height          =   495
         Left            =   7320
         TabIndex        =   18
         Top             =   6000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":1697F
         PICN            =   "frmHOMasterEmailTL.frx":1699B
         PICH            =   "frmHOMasterEmailTL.frx":1D19E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdNamaDealer 
         Height          =   495
         Left            =   2160
         TabIndex        =   19
         Top             =   2710
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":23D44
         PICN            =   "frmHOMasterEmailTL.frx":23D60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSumberRSA 
         Height          =   495
         Left            =   3480
         TabIndex        =   20
         Top             =   2710
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":271F2
         PICN            =   "frmHOMasterEmailTL.frx":2720E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdFO 
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   2710
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":2A67E
         PICN            =   "frmHOMasterEmailTL.frx":2A69A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdNoPolisi 
         Height          =   495
         Left            =   6120
         TabIndex        =   22
         Top             =   2710
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":2D869
         PICN            =   "frmHOMasterEmailTL.frx":2D885
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdModel 
         Height          =   495
         Left            =   7440
         TabIndex        =   23
         Top             =   2710
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":30C11
         PICN            =   "frmHOMasterEmailTL.frx":30C2D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKeterangan 
         Height          =   495
         Left            =   8640
         TabIndex        =   24
         Top             =   2710
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":33E76
         PICN            =   "frmHOMasterEmailTL.frx":33E92
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   ": F O"
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
         Left            =   3120
         TabIndex        =   38
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "#FO"
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
         Left            =   2400
         TabIndex        =   37
         Top             =   6480
         Width           =   615
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   ": Keterangan"
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
         Left            =   3120
         TabIndex        =   36
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   ": Model"
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
         Left            =   3120
         TabIndex        =   35
         Top             =   6960
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   ": Nomor Polisi"
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
         Left            =   3120
         TabIndex        =   34
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   ": Sumber RSA"
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
         Left            =   3120
         TabIndex        =   33
         Top             =   6240
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   ": Nama Dealer"
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
         Left            =   3120
         TabIndex        =   32
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "#KT"
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
         Left            =   2400
         TabIndex        =   31
         Top             =   7200
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "#MD"
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
         Left            =   2400
         TabIndex        =   30
         Top             =   6960
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "#NP"
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
         Left            =   2400
         TabIndex        =   29
         Top             =   6720
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "#SR"
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
         Left            =   2400
         TabIndex        =   28
         Top             =   6240
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "#NM"
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
         Left            =   2400
         TabIndex        =   27
         Top             =   6000
         Width           =   615
      End
      Begin VB.Shape Shape8 
         Height          =   1455
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   6000
         Width           =   4215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Body Email : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   3300
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   12255
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kategori Problem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   10695
         Begin VB.TextBox txtEmail2 
            Height          =   375
            Left            =   1560
            TabIndex        =   11
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox txtEmail1 
            Height          =   375
            Left            =   1560
            TabIndex        =   10
            Top             =   600
            Width           =   4575
         End
         Begin MSAdodcLib.Adodc AdcKode 
            Height          =   330
            Left            =   7800
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbDealer 
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   240
            Width           =   3195
            AllowInput      =   0   'False
            AllowNull       =   0   'False
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            RowHeight       =   423
            Columns(0).Width=   3200
            _ExtentX        =   5644
            _ExtentY        =   661
            _StockProps     =   93
            Text            =   "- Pilih Nama Dealer -"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -600
            TabIndex        =   14
            Top             =   300
            Width           =   2145
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email TL 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -600
            TabIndex        =   12
            Top             =   1020
            Width           =   2145
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email TL 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -600
            TabIndex        =   9
            Top             =   660
            Width           =   2145
         End
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   3
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":37243
         PICN            =   "frmHOMasterEmailTL.frx":3725F
         PICH            =   "frmHOMasterEmailTL.frx":3DB92
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid GridKomponen 
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "idemail"
         Columns(0).DataField=   "idemail"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Dealer"
         Columns(1).DataField=   "FS_NM_DEALER"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Email TL1"
         Columns(2).DataField=   "emailTL1"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Email TL2"
         Columns(3).DataField=   "emailTL2"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=8017"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7938"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
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
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=-1,.fontsize=825"
         _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H93C6FD&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
         _StyleDefs(73)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(74)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(75)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(76)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   4920
         TabIndex        =   5
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":44849
         PICN            =   "frmHOMasterEmailTL.frx":44865
         PICH            =   "frmHOMasterEmailTL.frx":4B008
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   6120
         TabIndex        =   6
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":51B31
         PICN            =   "frmHOMasterEmailTL.frx":51B4D
         PICH            =   "frmHOMasterEmailTL.frx":580AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   495
         Left            =   9720
         TabIndex        =   1
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":5E8D2
         PICN            =   "frmHOMasterEmailTL.frx":5E8EE
         PICH            =   "frmHOMasterEmailTL.frx":64EE7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   8520
         TabIndex        =   0
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":6B796
         PICN            =   "frmHOMasterEmailTL.frx":6B7B2
         PICH            =   "frmHOMasterEmailTL.frx":71FB5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   7350
         TabIndex        =   7
         Top             =   7080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":78B5B
         PICN            =   "frmHOMasterEmailTL.frx":78B77
         PICH            =   "frmHOMasterEmailTL.frx":7F6EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSettingEmail 
         Height          =   495
         Left            =   2640
         TabIndex        =   15
         Top             =   7080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOMasterEmailTL.frx":86484
         PICN            =   "frmHOMasterEmailTL.frx":864A0
         PICH            =   "frmHOMasterEmailTL.frx":8DC8B
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
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmHOMasterEmailTL.frx":95B11
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmHOMasterEmailTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOMasterEmailTL
' DateTime  : 15/11/2019 15:52
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

 
 
Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset

Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean


Private Sub ber_sih()
    txtEmail1.Text = ""
    txtEmail2.Text = ""
    cbDealer.Text = ""
    isTambah = True
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
 
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    cbDealer.Enabled = Flg
    txtEmail1.Enabled = Flg
    txtEmail2.Enabled = Flg
    
    GridKomponen.Enabled = Not Flg
    cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Not Flg
End Sub
 
Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    GridKomponen.Enabled = True
    Call refreshGrid
End Sub

Private Sub cmdEdit_Click()
    isTambah = False
    Call buka_Tutup(True)
End Sub

Private Sub cmdFO_Click()
    txtBodyEmail.Text = txtBodyEmail.Text & "#FO"
    txtBodyEmail.SetFocus
    txtBodyEmail.SelStart = Len(txtBodyEmail.Text)
End Sub

Private Sub cmdHapus_Click()
              
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         If MsgBox("Apakah yakin menghapus data dealer " & NulSID(rsGrid!fs_nm_dealer) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
4             Call objHT.OpenTransactionHO
5             isBegin = True
6             erTranAll = True
7             sSQL = "UPDATE m_emailTL SET userDel = '" & gSA & "',tglDel = GETDATE() WHERE idemail = '" & NulSID(rsGrid!idemail) & "'"
8             Call objHT.InsertData_Ho(sSQL)
9             If erTranAll Then
10                Call objHT.CloseTransactionHO
11                MsgBox "Data berhasil dihapus", vbInformation
12                Call refreshGrid
13                Call buka_Tutup(False)
14            Else
15                Call objHT.CloseTransactionHO_err
16                MsgBox "Gagal hapus data", vbInformation
17                Exit Sub
18            End If
19        End If

20       On Error GoTo 0
21       Exit Sub

cmdHapus_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
22        varErlError = Erl
23        varErrNumber = err.Number
24        If isBegin Then
25            Call objHT.CloseTransactionHO_err
26        End If
27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmHOMasterEmailTL")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdKeterangan_Click()
    txtBodyEmail.Text = txtBodyEmail.Text & "#KT"
    txtBodyEmail.SetFocus
    txtBodyEmail.SelStart = Len(txtBodyEmail.Text)
End Sub

Private Sub cmdModel_Click()
    txtBodyEmail.Text = txtBodyEmail.Text & "#MD"
    txtBodyEmail.SetFocus
    txtBodyEmail.SelStart = Len(txtBodyEmail.Text)
End Sub

Private Sub cmdModelJudul_Click()

    txtJudulEmail.Text = txtJudulEmail.Text & "#MD"
    txtJudulEmail.SetFocus
    txtJudulEmail.SelStart = Len(txtJudulEmail.Text)
End Sub

Private Sub cmdNamaDealer_Click()
    txtBodyEmail.Text = txtBodyEmail.Text & "#NM"
    txtBodyEmail.SetFocus
    txtBodyEmail.SelStart = Len(txtBodyEmail.Text)
End Sub

Private Sub cmdNamaDealerJudul_Click()
    txtJudulEmail.Text = txtJudulEmail.Text & "#NM"
    txtJudulEmail.SetFocus
    txtJudulEmail.SelStart = Len(txtJudulEmail.Text)
End Sub

Private Sub cmdNoPolisi_Click()
    txtBodyEmail.Text = txtBodyEmail.Text & "#NP"
    txtBodyEmail.SetFocus
    txtBodyEmail.SelStart = Len(txtBodyEmail.Text)
End Sub

Private Sub cmdNopolJudul_Click()
    
    txtJudulEmail.Text = txtJudulEmail.Text & "#NP"
    txtJudulEmail.SetFocus
    txtJudulEmail.SelStart = Len(txtJudulEmail.Text)
End Sub

Private Sub cmdSettingEmail_Click()
    Dim rsSetting As ADODB.Recordset
    Set rsSetting = objO.FetchDataHO("SELECT alamatATL1,alamatATL2,judulEmail,bodyEmail,pengirim FROM settingEmailRSA WITH (NOLOCK)")
    If Not rsSetting.EOF Then
        txtEmailATL1.Text = Trim(rsSetting!alamatATL1)
        txtEmailATL2.Text = Trim(rsSetting!alamatATL2)
        txtJudulEmail.Text = Trim(rsSetting!judulEmail)
        txtBodyEmail.Text = Trim(rsSetting!bodyEmail)
        txtPengirim.Text = Trim(rsSetting!Pengirim)
    End If
    frameSettingEmail.Visible = True
End Sub

Private Sub cmdSimpan_Click()
    
   On Error GoTo cmdSimpan_Click_Error
    Dim isBegin     As Boolean
    Dim boleh       As Boolean
    Dim pesan       As String
    Dim allDealer   As Boolean
    Dim rsCek       As ADODB.Recordset
    allDealer = False
    boleh = True
    isBegin = False
    
    If NulSID(cbDealer.value) = "" Then
        boleh = False
        pesan = " + Dealer harus dipilih."
    End If
    
    If NulSID(txtEmail1.Text) = "" And NulSID(txtEmail2.Text) = "" Then
        boleh = False
        pesan = " + Email harus diisi."
    End If
    
    If isTambah And boleh Then
        Set rsCek = objO.FetchDataHO("SELECT * FROM m_emailTL WITH (NOLOCK) WHERE FS_KD_DEALER = '" & NulSID(cbDealer.value) & "' and userDel = ''")
        If Not rsCek.EOF Then
            boleh = False
            pesan = " + Dealer sudah ada data, ganti dealer lain."
        End If
    ElseIf isTambah = False And boleh Then
        Set rsCek = objO.FetchDataHO("SELECT * FROM m_emailTL WITH (NOLOCK) WHERE FS_KD_DEALER = '" & NulSID(cbDealer.value) & "' and userDel = '' AND idemail <> '" & NulSID(rsGrid!idemail) & "'")
        If Not rsCek.EOF Then
            boleh = False
            pesan = " + Dealer sudah ada data, ganti dealer lain."
        End If
    End If
    If boleh Then
        Screen.MousePointer = vbHourglass
        Call objHT.OpenTransactionHO
        isBegin = True
        If isTambah Then
            sSQL = "INSERT INTO m_emailTL (FS_KD_DEALER,emailTL1,emailTL2,userEntry,tglEntry) VALUES ('" & NulSID(cbDealer.value) & "','" & NulSID(txtEmail1.Text) & "','" & NulSID(txtEmail2.Text) & "','" & gSA & "',GETDATE())"
        Else
            sSQL = "UPDATE m_emailTL SET FS_KD_DEALER = '" & NulSID(cbDealer.value) & "',emailTL1 = '" & NulSID(txtEmail1.Text) & "',emailTL2 = '" & NulSID(txtEmail2.Text) & "' WHERE idemail = '" & NulSID(rsGrid!idemail) & "'"
        End If
        Call objHT.InsertData_Ho(sSQL)
        If erTranAll = False Then
            Call objHT.CloseTransactionHO
            Screen.MousePointer = vbDefault
            MsgBox "Berhasil menyimpan data.", vbInformation
            Call ber_sih
            Call buka_Tutup(False)
            Call refreshGrid
        Else
errorRollback:
            Call objHT.CloseTransactionHO_err
            Screen.MousePointer = vbDefault
            MsgBox "Gagal menyimpan data.", vbInformation
            Exit Sub
        End If
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
    End If

   On Error GoTo 0
   Exit Sub

cmdSimpan_Click_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    Screen.MousePointer = vbDefault
    If isBegin = True Then
        Call objHT.CloseTransactionHO_err
    End If
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOMasterEmailTL")
End Sub

Private Sub cmdSimpanEmail_Click()
          Dim boleh   As Boolean
          Dim pesan   As String
          Dim begin   As Boolean
          Dim rsEmail As ADODB.Recordset
          
1        On Error GoTo cmdSimpanEmail_Click_Error
2         begin = False
3         boleh = True
4         If NulSID(txtEmailATL1.Text) = "" Or NulSID(txtEmailATL2.Text) = "" Then
5             boleh = False
6             pesan = pesan & " Alamat email ATL harus diisi."
7         End If
8         If NulSID(txtJudulEmail.Text) = "" Then
9             boleh = False
10            pesan = pesan & " Judul Email harus diisi."
11        End If
12        If NulSID(txtBodyEmail.Text) = "" Then
13            boleh = False
14            pesan = pesan & " Bodi Email harus diisi."
15        End If
16        If boleh Then
17            Call objHT.OpenTransactionHO
18            begin = True
                  
19            Set rsEmail = objHT.FetchDataHO("SELECT * FROM settingEmailRSA WITH (NOLOCK) ")
20            If Not rsEmail.EOF Then
21                sSQL = "UPDATE settingEmailRSA SET pengirim = '" & Trim(txtPengirim.Text) & "',alamatATL1 = '" & NulSID(txtEmailATL1.Text) & "',alamatATL2 = '" & NulSID(txtEmailATL2.Text) & "',judulEmail = '" & Trim(txtJudulEmail.Text) & "',bodyEmail = '" & Trim(txtBodyEmail.Text) & "',userEdit = '" & gSA & "',tglEdit = getdate()"
22            Else
23                sSQL = "INSERT INTO settingEmailRSA (alamatATL1,alamatATL2,judulEmail,bodyEmail,userEntry,tglEntry,pengirim) VALUES ('" & NulSID(txtEmailATL1.Text) & "','" & NulSID(txtEmailATL2.Text) & "','" & Trim(txtJudulEmail.Text) & "','" & Trim(txtBodyEmail.Text) & "','" & gSA & "',GETDATE(),'" & Trim(txtPengirim.Text) & "')"
24            End If
              
25            Call objHT.InsertData_Ho(sSQL)
26            Call objHT.CloseTransactionHO
27            MsgBox "Data berhasil disimpan", vbInformation
28            frameSettingEmail.Visible = False
29            Exit Sub
30        Else
31            MsgBox "Silahkan isi data berikut : " & vbCrLf & pesan, vbInformation
32            Exit Sub
33        End If

34       On Error GoTo 0
35       Exit Sub

cmdSimpanEmail_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
36        varErlError = Erl
37        varErrNumber = err.Number
38        If begin Then
39            Call objHT.CloseTransactionHO_err
40        End If
41        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpanEmail_Click of Form frmHOMasterEmailTL")
End Sub

Private Sub cmdSumberRSA_Click()
    txtBodyEmail.Text = txtBodyEmail.Text & "#SR"
    txtBodyEmail.SetFocus
    txtBodyEmail.SelStart = Len(txtBodyEmail.Text)
End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    Call buka_Tutup(True)
    GridKomponen.Enabled = False
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False)
    Call FetchCombo
    Call refreshGrid
    frameSettingEmail.Visible = False
End Sub
  
Private Sub FetchCombo()
    AdcKode.ConnectionString = SQLConHO
    AdcKode.RecordSource = "select FS_KD_DEALER, FS_NM_DEALER,VPNCABANG,DBASECBG,USERCBG,PASSCBG,NEWENKRIP from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
    cbDealer.DataSourceList = AdcKode
    cbDealer.DataFieldList = "FS_KD_DEALER"
    cbDealer.DataFieldToDisplay = "FS_NM_DEALER"
    cbDealer.Columns(0).Width = 1200
    cbDealer.Columns(1).Width = 4000
    cbDealer.Columns(0).Caption = "Kode Dealer"
    cbDealer.Columns(1).Caption = "Nama Dealer"
    cbDealer.Columns(2).Visible = False
    cbDealer.Columns(3).Visible = False
    cbDealer.Columns(4).Visible = False
    cbDealer.Columns(5).Visible = False
    cbDealer.Columns(6).Visible = False
End Sub
  

Private Sub gridkomponen_FilterChange()
    On Error GoTo errHandler
    Set cols = GridKomponen.Columns
    Dim C As Integer
    
    C = GridKomponen.col
    GridKomponen.HoldFields
    rsGrid.filter = getFilter()
    GridKomponen.col = C
    GridKomponen.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridkomponen_HeadClick(ByVal ColIndex As Integer)
          
1        On Error GoTo gridkomponen_HeadClick_Error
2         Urut = IIf(Urut = "DESC", "ASC", "DESC")
3         rsGrid.Sort = GridKomponen.Columns(ColIndex).DataField & " " & Urut
4         GridKomponen.SetFocus
5         Sendkeys "{Right " & Val(ColIndex) & "}"
          

6        On Error GoTo 0
7        Exit Sub

gridkomponen_HeadClick_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
8         varErlError = Erl
9         varErrNumber = err.Number
10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure gridkomponen_HeadClick of Form frmHOMasterEmailTL")
End Sub
 
 

 
Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 15
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub refreshGrid()
    sSQL = "select a.idemail,a.FS_KD_DEALER,b.fs_nm_dealer,a.emailTL1,a.emailTL2 from m_emailTL a WITH (NOLOCK) LEFT JOIN T_dealer b WITH (NOLOCK) ON a.fs_kd_dealer = b.fs_kd_dealer WHERE a.USERDEL = ''"
    Set rsGrid = objO.FetchDataHO(sSQL)
    Set GridKomponen.DataSource = rsGrid
End Sub

Private Sub hargaBeli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub hargaJual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
End Sub
 

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Private Sub GridKomponen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            txtEmail1.Text = NulSID(rsGrid!emailTL1)
            txtEmail2.Text = NulSID(rsGrid!emailTL2)
            cbDealer.value = NulSID(rsGrid!FS_KD_DEALER)
        End If
    End If
End Sub

Private Sub txtemail1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtemail1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub
 

Private Sub txtEmail2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub txtEmail2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub XPButton1_Click()
    frameSettingEmail.Visible = False
End Sub
