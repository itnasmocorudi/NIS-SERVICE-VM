VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRetensiAsuransiCall 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   Icon            =   "frmRetensiAsuransiCall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmRetensiAsuransiCall.frx":29C12
   ScaleHeight     =   10500
   ScaleWidth      =   12060
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   10000
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   94
         Top             =   1440
         Width           =   2610
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   7380
         MaxLength       =   20
         TabIndex        =   90
         Top             =   1425
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   7380
         MaxLength       =   20
         TabIndex        =   75
         Top             =   1095
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   8
         Left            =   1545
         MaxLength       =   2
         TabIndex        =   74
         Top             =   1455
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   9
         Left            =   2055
         MaxLength       =   4
         TabIndex        =   73
         Top             =   1455
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   10
         Left            =   2805
         MaxLength       =   3
         TabIndex        =   72
         Top             =   1455
         Visible         =   0   'False
         Width           =   635
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   71
         Top             =   465
         Width           =   3885
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1545
         MaxLength       =   50
         TabIndex        =   70
         Top             =   1125
         Width           =   3885
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1545
         MaxLength       =   25
         TabIndex        =   69
         Top             =   795
         Width           =   2610
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3180
         MaxLength       =   20
         TabIndex        =   68
         Top             =   1845
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   7380
         MaxLength       =   210
         TabIndex        =   67
         Top             =   465
         Width           =   4005
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7380
         MaxLength       =   300
         TabIndex        =   66
         Top             =   795
         Width           =   4005
      End
      Begin VB.PictureBox PicMove 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   105
         ScaleHeight     =   7095
         ScaleWidth      =   11235
         TabIndex        =   1
         Top             =   2010
         Width           =   11235
         Begin VB.Frame FrThird 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   105
            TabIndex        =   55
            Top             =   4500
            Width           =   11250
            Begin VB.ComboBox cboPos 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   7440
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   1125
               Width           =   3495
            End
            Begin VB.ComboBox cboPos 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   3795
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   1125
               Width           =   3495
            End
            Begin VB.ComboBox cboPos 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               ItemData        =   "frmRetensiAsuransiCall.frx":2E8FD
               Left            =   165
               List            =   "frmRetensiAsuransiCall.frx":2E8FF
               Style           =   2  'Dropdown List
               TabIndex        =   63
               Top             =   1125
               Width           =   3495
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Terorisme && sabotase"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   5835
               TabIndex        =   59
               Top             =   405
               Width           =   3570
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Gempa bumi, tsunami"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   5820
               TabIndex        =   58
               Top             =   105
               Width           =   3570
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Huru hara && kerusuhan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   57
               Top             =   405
               Width           =   3570
            End
            Begin VB.CheckBox Ch 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Banjir termasuk angin topan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   165
               TabIndex        =   56
               Top             =   105
               Width           =   3570
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kecelakaan diri penumpang"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   7440
               TabIndex        =   62
               Top             =   885
               Width           =   2715
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kecelakaan diri driver"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   3795
               TabIndex        =   61
               Top             =   900
               Width           =   2175
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pihak ke-3"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   165
               TabIndex        =   60
               Top             =   885
               Width           =   1050
            End
         End
         Begin VB.Frame FrJenis 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1545
            Left            =   105
            TabIndex        =   29
            Top             =   2985
            Width           =   11250
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   7
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   1080
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   6
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   765
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   5
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   450
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   150
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   1080
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   765
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   450
               Width           =   2325
            End
            Begin VB.ComboBox cbojns 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1410
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   150
               Width           =   2325
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   7000
               TabIndex        =   54
               Top             =   1125
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 8"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   5955
               TabIndex        =   53
               Top             =   1125
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   7000
               TabIndex        =   51
               Top             =   810
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 7"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   5950
               TabIndex        =   50
               Top             =   810
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   7000
               TabIndex        =   48
               Top             =   495
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 6"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   5950
               TabIndex        =   47
               Top             =   495
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   7000
               TabIndex        =   45
               Top             =   195
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 5"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   5950
               TabIndex        =   44
               Top             =   195
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   1260
               TabIndex        =   42
               Top             =   1125
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 4"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   150
               TabIndex        =   41
               Top             =   1125
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   1260
               TabIndex        =   39
               Top             =   810
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 3"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   150
               TabIndex        =   38
               Top             =   810
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   1260
               TabIndex        =   36
               Top             =   495
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 2"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   150
               TabIndex        =   35
               Top             =   495
               Width           =   780
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   1260
               TabIndex        =   33
               Top             =   195
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun 1"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   150
               TabIndex        =   32
               Top             =   195
               Width           =   780
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               Caption         =   "Jenis Asuransi "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   75
               TabIndex        =   30
               Top             =   -45
               Width           =   1470
            End
            Begin VB.Shape Shape6 
               BorderColor     =   &H000080FF&
               Height          =   1455
               Left            =   30
               Top             =   45
               Width           =   11085
            End
         End
         Begin VB.Frame FrTenor 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   90
            TabIndex        =   20
            Top             =   1995
            Width           =   11250
            Begin VB.ComboBox cboTenor 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7140
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   165
               Width           =   1140
            End
            Begin VB.ComboBox cboLea 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1425
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   165
               Width           =   2325
            End
            Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DcAsuransi 
               Height          =   315
               Left            =   1425
               TabIndex        =   93
               Tag             =   "Kode Supplier TAM"
               Top             =   495
               Width           =   3090
               BevelType       =   0
               _Version        =   196616
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox3D      =   0   'False
               ForeColorEven   =   0
               BackColorEven   =   12648447
               BackColorOdd    =   16777215
               RowHeight       =   423
               Columns(0).Width=   3200
               Columns(0).DataType=   8
               Columns(0).FieldLen=   4096
               _ExtentX        =   5450
               _ExtentY        =   556
               _StockProps     =   93
               BackColor       =   -2147483643
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
            Begin VB.Shape Shape1 
               BorderColor     =   &H000080FF&
               Height          =   900
               Left            =   30
               Top             =   30
               Width           =   11100
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   7000
               TabIndex        =   28
               Top             =   225
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   1275
               TabIndex        =   27
               Top             =   540
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   1275
               TabIndex        =   26
               Top             =   225
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tenor"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   5950
               TabIndex        =   23
               Top             =   210
               Width           =   570
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Asuransi"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   165
               TabIndex        =   22
               Top             =   540
               Width           =   855
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Leasing"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   165
               TabIndex        =   21
               Top             =   225
               Width           =   750
            End
         End
         Begin VB.Frame FrFollow 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   1830
            Left            =   105
            TabIndex        =   6
            Top             =   135
            Width           =   11250
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   1440
               MaxLength       =   200
               TabIndex        =   19
               Top             =   945
               Width           =   9540
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   390
               Left            =   1440
               TabIndex        =   14
               Top             =   540
               Width           =   3450
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tidak Deal"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   3
                  Left            =   1455
                  TabIndex        =   16
                  Top             =   0
                  Width           =   2250
               End
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Deal"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   2
                  Left            =   0
                  TabIndex        =   15
                  Top             =   0
                  Width           =   1410
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   390
               Left            =   1440
               TabIndex        =   11
               Top             =   135
               Width           =   3450
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Tidak Terhubung"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   1
                  Left            =   1455
                  TabIndex        =   13
                  Top             =   0
                  Width           =   2250
               End
               Begin VB.OptionButton Opt 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Terhubung"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   0
                  Left            =   0
                  TabIndex        =   12
                  Top             =   0
                  Width           =   1410
               End
            End
            Begin MSAdodcLib.Adodc AdoAs 
               Height          =   330
               Left            =   6075
               Top             =   120
               Visible         =   0   'False
               Width           =   1200
               _ExtentX        =   2117
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
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   1290
               TabIndex        =   18
               Top             =   1005
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   1290
               TabIndex        =   17
               Top             =   638
               Width           =   60
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   1290
               TabIndex        =   10
               Top             =   233
               Width           =   60
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alasan"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   165
               TabIndex        =   9
               Top             =   1005
               Width           =   660
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hasil"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   165
               TabIndex        =   8
               Top             =   638
               Width           =   480
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Follow Up"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   165
               TabIndex        =   7
               Top             =   233
               Width           =   945
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H000080FF&
               Height          =   1710
               Left            =   30
               Top             =   75
               Width           =   11085
            End
         End
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   8955
         TabIndex        =   2
         Top             =   9360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&s"
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
         MICON           =   "frmRetensiAsuransiCall.frx":2E901
         PICN            =   "frmRetensiAsuransiCall.frx":2E91D
         PICH            =   "frmRetensiAsuransiCall.frx":2EE30
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   10200
         TabIndex        =   3
         Top             =   9360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&x"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRetensiAsuransiCall.frx":2F3D2
         PICN            =   "frmRetensiAsuransiCall.frx":2F3EE
         PICH            =   "frmRetensiAsuransiCall.frx":2F8C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   98
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   97
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1425
         TabIndex        =   96
         Top             =   1500
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   360
         TabIndex        =   95
         Top             =   1500
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7245
         TabIndex        =   92
         Top             =   1485
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   6165
         TabIndex        =   91
         Top             =   1485
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   7245
         TabIndex        =   89
         Top             =   1155
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6165
         TabIndex        =   88
         Top             =   1155
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1425
         TabIndex        =   87
         Top             =   1518
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   86
         Top             =   1875
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1425
         TabIndex        =   85
         Top             =   528
         Width           =   60
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NoRangka"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   84
         Top             =   528
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   83
         Top             =   1215
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1425
         TabIndex        =   82
         Top             =   1215
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1965
         TabIndex        =   81
         Top             =   1905
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3045
         TabIndex        =   80
         Top             =   1905
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6165
         TabIndex        =   79
         Top             =   855
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   7245
         TabIndex        =   78
         Top             =   855
         Width           =   60
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   7245
         TabIndex        =   77
         Top             =   525
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6165
         TabIndex        =   76
         Top             =   525
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   6000
         TabIndex        =   5
         Top             =   9480
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".::Poses Call Customer ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4567
         TabIndex        =   4
         Top             =   60
         Width           =   2400
      End
      Begin VB.Line Line24 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11520
         Y1              =   9960
         Y2              =   9960
      End
      Begin VB.Line Line23 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11520
         Y1              =   9240
         Y2              =   9240
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11295
      End
   End
End
Attribute VB_Name = "frmRetensiAsuransiCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
Private rsFind      As New ADODB.Recordset
Private rsOli       As New ADODB.Recordset
Private rsGridJob   As New ADODB.Recordset
Private rsGridPart  As New ADODB.Recordset
Private rsD         As New ADODB.Recordset
Private rsProg      As New ADODB.Recordset
Private rsDsc       As New ADODB.Recordset
Private rso         As ADODB.Recordset
Private rsP         As ADODB.Recordset
Private isShow      As String
Private isAkt       As String
Private ntxt(12)    As String
Private nWOHO       As String
Private isModel     As String
Private nBook       As String
Private nProg       As String
Private ToTPart     As Double
Private TotOli      As Double
Private nPart       As Double
Private nJob        As Double
Private DiscProg    As Double
Private isBtl       As Boolean
Private isClose     As Boolean
Private ValidProg   As Boolean
Private n           As Integer
Private nKey        As Integer
Private rsRepl      As New ADODB.Recordset
Private sSQL        As String
Private Tri         As New ADODB.Recordset
Private KodeReQ     As String
Private isFalse     As String
Private idAsuransi  As Integer
Private pengirimEmail   As String
Private toEmail         As String
Private ccEmail         As String
Private rsEmail         As ADODB.Recordset

Public Property Let myShow(value As String)
    isShow = value
End Property
Public Property Let myAkt(value As String)
    isAkt = value
End Property
Public Property Let MyData(value As ADODB.Recordset)
    Set rsD = value
End Property


Public Property Let MyIdAsuransi(value As Integer)
    idAsuransi = value
End Property


Private Sub cmdSave_Click()
1       On Error GoTo Erro
2       If isValid = True Then
            ''20201216 mengambil email pic asuransi
            ''----------------------------------------
3           sSQL = "select Email1,* from profilspk..profil"
4           pengirimEmail = objO.FetchData_String(sSQL)
            
5           sSQL = "select * from oss..MasterEmail"
6           Set rsEmail = objO.FetchDataHO(sSQL)
7           ccEmail = ""
8           If Not rsEmail.EOF Then
9             rsEmail.MoveFirst
10            Do While Not rsEmail.EOF
11                If NulSID(rsEmail!Email) <> "" Then
12                    ccEmail = ccEmail & NulSID(rsEmail!Email) & ";"
13                End If
14                rsEmail.MoveNext
15            Loop
16          End If
            
17          sSQL = "select b.Email,b.Email2"
18          sSQL = sSQL & " from OSS..MstAsuransi A"
19          sSQL = sSQL & " left join OSS..[User] b on b.areaAsuransi = cast(a.KodeAsuransi as varchar(50))"
20          sSQL = sSQL & " where a.KodeAsuransi = " & CInt(DcAsuransi.Columns(2).value)
21          Set rsEmail = objO.FetchDataHO(sSQL)
22          toEmail = ""
23          If Not rsEmail.EOF Then
24            rsEmail.MoveFirst
25            Do While Not rsEmail.EOF
26                If NulSID(rsEmail!Email) <> "" Then
27                    toEmail = toEmail & NulSID(rsEmail!Email) & ";"
28                End If
29                If NulSID(rsEmail!Email2) <> "" Then
30                    toEmail = toEmail & NulSID(rsEmail!Email2) & ";"
31                End If
32                rsEmail.MoveNext
33            Loop
34          End If
            
35          Select Case UCase(isAkt)
                Case "NEW CUSTOMER"
36                  KodeReQ = "": Set rso = Nothing: Set rso = New ADODB.Recordset: rsD.MoveFirst
37                  sSQL = "SELECT TOP 1 KodeReq FROM FUAsuransi With (NoLock) WHERE SUBSTRING(KodeReq,1,11) = '" & Left(Trim(rsD("DealerPenjual")), 4) & "-NW" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & "' ORDER BY KodeReq DESC" '2020-NW2010
38                  Set rso = objO.FetchData(sSQL)
39                  If rso.RecordCount > 0 Then
40                      KodeReQ = Left(Trim(rsD("DealerPenjual")), 4) & "-NW" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & Format(Right(rso("kodereq"), 4) + 1, "0000")
41                  Else
42                      KodeReQ = Left(Trim(rsD("DealerPenjual")), 4) & "-NW" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & "0001"
43                  End If
44                  If Opt(2).value = True Then
45                      sSQL = "INSERT INTO FUAsuransi (Aktifitas,Nama,Alamat,Telp,HP1,HP2,NoRangka,Model, " & _
                               "NoPolisi,NamaLeasing,KodeAsuransi,MasaberlakuAwal,MasaberlakuAkhir,Tenor, " & _
                               "JenisAsuransi1,JenisAsuransi2,JenisAsuransi3,JenisAsuransi4,JenisAsuransi5,JenisAsuransi6,JenisAsuransi7, " & _
                               "JenisAsuransi8,DealerPenjual,Terhubung,Deal,Alasan,TglFU,userFU,kodeReq, " & _
                               "Banjir,Gempa,SRCC,Teroris,TanggungP3,CelakaDiri,CelakaDiriPass) VALUES " & _
                               "('" & NulSID("new") & "','" & s2d(Trim(txt(0).Text)) & "','" & s2d(Trim(txt(1).Text)) & "','" & s2d(Trim(txt(2).Text)) & "','" & s2d(Trim(txt(3).Text)) & "','" & s2d(Trim(txt(4).Text)) & "','" & Trim(txt(5).Text) & "','" & s2d(NulSID(rsD!modelkategory)) & "', " & _
                               "'" & NulSID(rsD!NoPolisi) & "','" & Trim(rsD("kode_leasing")) & "'," & CInt(DcAsuransi.Columns(2).value) & ",'" & Format(rsD("masaberlakuawal"), "YYYY-MM-DD") & "','" & Format(rsD("masaberlakuakhir"), "YYYY-MM-DD") & "'," & CInt(cboTenor.Text) & ", " & _
                               "'" & s2d(Trim(isAs(cbojns(0)))) & "','" & s2d(Trim(isAs(cbojns(1)))) & "','" & s2d(Trim(isAs(cbojns(2)))) & "','" & s2d(Trim(isAs(cbojns(3)))) & "','" & isAs(s2d(Trim(cbojns(4)))) & "','" & isAs(s2d(Trim(cbojns(5)))) & "','" & s2d(Trim(isAs(cbojns(6)))) & "', " & _
                               "'" & s2d(Trim(isAs(cbojns(7)))) & "','" & Trim(Left(rsD("DealerPenjual"), 4)) & "','" & IIf(Opt(0).value = True, 1, 0) & "','" & IIf(Opt(2).value = True, 1, 0) & "','" & Trim(s2d(txt(11).Text)) & "',Getdate(),'" & Trim(gSA) & "', '" & Trim(KodeReQ) & "', " & _
                               "'" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ")"
46                  Else
47                      sSQL = "INSERT INTO FUAsuransi (Aktifitas,Nama,Alamat,Telp,HP1,HP2,NoRangka,Model," & _
                               "NoPolisi,MasaberlakuAwal,MasaberlakuAkhir," & _
                               "DealerPenjual,Terhubung,Deal,Alasan,TglFU,userFU,kodeReq) VALUES " & _
                               "('" & NulSID("new") & "','" & s2d(Trim(txt(0).Text)) & "','" & s2d(Trim(txt(1).Text)) & "','" & s2d(Trim(txt(2).Text)) & "','" & Trim(txt(3).Text) & "','" & s2d(Trim(txt(4).Text)) & "','" & Trim(txt(5).Text) & "','" & s2d(NulSID(rsD!modelkategory)) & "', " & _
                               "'" & IIf(Trim(txt(8).Text) <> "", Trim(txt(8).Text) & "-" & Trim(txt(9).Text) & "-" & Trim(txt(10).Text), "") & "','" & Format(rsD("masaberlakuawal"), "YYYY-MM-DD") & "','" & Format(rsD("masaberlakuakhir"), "YYYY-MM-DD") & "', " & _
                               "'" & Trim(Left(rsD("DealerPenjual"), 4)) & "','" & IIf(Opt(0).value = True, 1, 0) & "','" & IIf(Opt(2).value = True, 1, 0) & "','" & Trim(s2d(txt(11).Text)) & "',Getdate(),'" & Trim(UserLog) & "', '" & Trim(KodeReQ) & "')"
48                  End If
49                  AdaBegin = False: KetEro = ""
50                  Call objHT.OpenTransactionBengkel
51                  Call objHT.OpenTransactionHO
52                  AdaBegin = True
53                  Call objHT.InsertData_Bengkel(sSQL)
54                  sSQL = ""
55                  If Opt(2).value = True Then
56                      sSQL = "INSERT INTO OSS..Asuransi (NoSPK,NamaPemesan,Alamat,Kota,telp_rumah,HP1,HP2," & _
                               "NamaTertanggung,telp_rumahTT,AlamatTT,KotaTT,HP1TT,HP2TT," & _
                               "Kode_type,Kode_warna,NamaSales,noRangka,noMesin,thnUnit," & _
                               "kodeDealer,banjir,Gempa,SRCC,Teroris," & _
                               "TanggungP3,CelakaDiri,userEntry,tglEntry," & _
                               "areaAsuransi,tenorAsuransi,KodeAsuransi,JenisAsuransi,JenisAsuransi2,JenisAsuransi3,JenisAsuransi4,JenisAsuransi5,JenisAsuransi6,JenisAsuransi7,JenisAsuransi8,CelakaDiriPass) " & _
                               "VALUES ('" & NulSID(KodeReQ) & "','" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(txt(1).Text))) & "','','" & Trim(s2d(Nul(txt(2).Text))) & "','" & Trim(s2d(Nul(txt(3).Text))) & "','" & Trim(s2d(Nul(txt(4).Text))) & "'," & _
                               "'" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(txt(2).Text))) & "','" & Trim(s2d(Nul(rsD("AlamatTT")))) & "','" & Trim(s2d(Nul(rsD("KotaTT")))) & "','" & Trim(s2d(Nul(rsD("HP1TT")))) & "','" & Trim(s2d(Nul(rsD("HP2TT")))) & "'," & _
                               "'" & Trim(s2d(Nul(rsD("kodekategori")))) & "','" & Left(Trim(s2d(Nul(rsD("warna")))), 7) & "','" & Trim(s2d(Nul(rsD("namasales")))) & "','" & Trim(s2d(Nul(rsD("norangka")))) & "','" & Trim(s2d(Nul(rsD("nomesin")))) & "','" & Trim(s2d(Nul(rsD("kodetahun")))) & "'," & _
                               "'" & Left(Trim(s2d(Nul(rsD("DealerPenjual")))), 4) & "','" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & _
                               "" & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & ",'" & Trim(gSA) & "',GETDATE(), " & _
                               "'" & Trim(DcAsuransi.Columns(1).value) & "','" & CInt(cboTenor.Text) & "','" & Trim(DcAsuransi.Columns(2).value) & "','" & isXs(cbojns(0).Text) & "','" & isXs(cbojns(1).Text) & "','" & isXs(cbojns(2).Text) & "','" & isXs(cbojns(3).Text) & "','" & isXs(cbojns(4).Text) & "','" & isXs(cbojns(5).Text) & "','" & isXs(cbojns(6).Text) & "','" & isXs(cbojns(7).Text) & "'," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ")"
57                        Call objHT.InsertData_Ho(sSQL)
                          
58                        sSQL = "INSERT INTO [" & dbReplikasiHO & "]..SendMail_Hansarenkei "
59                        sSQL = sSQL & " (KodeReq,Aktifitas,KodeDealer,NamaDealer,Perusahaan,NamaCustomer,MDL,FullModel,"
60                        sSQL = sSQL & " NoRangka,Tujuan,CC,"
61                        sSQL = sSQL & " Pengirim,UserEntry,TglEntry,AppEntry,Tipe) values ("
62                        sSQL = sSQL & " '" & NulSID(KodeReQ) & "','NEW','" & Right(gDb, 2) & "','" & NulSID(Log.LogNamaNas) & "','" & NulSID(DcAsuransi.Columns(3).value) & "','" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(rsD("kodekategori")))) & "','" & Trim(s2d(Nul(txt(7).Text))) & "',"
63                        sSQL = sSQL & " '" & Trim(s2d(Nul(rsD("norangka")))) & "','" & toEmail & "','" & ccEmail & "',"
64                        sSQL = sSQL & " '" & pengirimEmail & "','" & gSA & "',GETDATE(),'SVC','Asuransi')"
65                        Call objHT.InsertData_Bengkel(sSQL)
66                  Else
                        
67                  End If
68                  Call objHT.CloseTransactionBengkel
69                  Call objHT.CloseTransactionHO
70                  AdaBegin = False
71                  MsgBox "Data Berhasil diproses!", vbInformation
72                  Unload Me
73              Case NulSID("Upgrade Customer")
74                  KodeReQ = "": Set rso = Nothing: Set rso = New ADODB.Recordset: rsD.MoveFirst
75                  sSQL = "SELECT TOP 1 KodeReq FROM FUAsuransi With (NoLock) WHERE SUBSTRING(KodeReq,1,11) = '" & Left(Trim(rsD("DealerPenjual")), 4) & "-UG" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & "'  ORDER BY KodeReq DESC" '2020-NW2010
76                  Set rso = objO.FetchData(sSQL)
77                  If rso.RecordCount > 0 Then
78                      KodeReQ = Left(Trim(rsD("DealerPenjual")), 4) & "-UG" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & Format(Right(rso("kodereq"), 4) + 1, "0000")
79                  Else
80                      KodeReQ = Left(Trim(rsD("DealerPenjual")), 4) & "-UG" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & "0001"
81                  End If
82                  If Opt(2).value = True Then
83                      sSQL = "INSERT INTO FUAsuransi (Aktifitas,Nama,Alamat,Telp,HP1,HP2,NoRangka,Model, " & _
                               "NoPolisi,NamaLeasing,KodeAsuransi,MasaberlakuAwal,MasaberlakuAkhir,Tenor, " & _
                               "JenisAsuransi1,JenisAsuransi2,JenisAsuransi3,JenisAsuransi4,JenisAsuransi5,JenisAsuransi6,JenisAsuransi7, " & _
                               "JenisAsuransi8,DealerPenjual,Terhubung,Deal,Alasan,TglFU,userFU,kodeReq, " & _
                               "Banjir,Gempa,SRCC,Teroris,TanggungP3,CelakaDiri,CelakaDiriPass) VALUES " & _
                               "('" & NulSID("upgrade") & "','" & s2d(Trim(txt(0).Text)) & "','" & s2d(Trim(txt(1).Text)) & "','" & s2d(Trim(txt(2).Text)) & "','" & s2d(Trim(txt(3).Text)) & "','" & s2d(Trim(txt(4).Text)) & "','" & Trim(txt(5).Text) & "','" & Left(NulSID(rsD!modelkategory), 25) & "', " & _
                               "'" & NulSID(rsD!NoPolisi) & "','" & Trim(rsD("kode_leasing")) & "'," & CInt(DcAsuransi.Columns(2).value) & ",'" & Format(rsD("masaberlakuawal"), "YYYY-MM-DD") & "','" & Format(rsD("masaberlakuakhir"), "YYYY-MM-DD") & "'," & CInt(cboTenor.Text) & ", " & _
                               "'" & s2d(Trim(isAs(cbojns(0)))) & "','" & s2d(Trim(isAs(cbojns(1)))) & "','" & s2d(Trim(isAs(cbojns(2)))) & "','" & s2d(Trim(isAs(cbojns(3)))) & "','" & isAs(s2d(Trim(cbojns(4)))) & "','" & isAs(s2d(Trim(cbojns(5)))) & "','" & s2d(Trim(isAs(cbojns(6)))) & "', " & _
                               "'" & s2d(Trim(isAs(cbojns(7)))) & "','" & Trim(Left(rsD("DealerPenjual"), 4)) & "','" & IIf(Opt(0).value = True, 1, 0) & "','" & IIf(Opt(2).value = True, 1, 0) & "','" & Trim(s2d(txt(11).Text)) & "',Getdate(),'" & Trim(gSA) & "', '" & Trim(KodeReQ) & "', " & _
                               "'" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ")"
84                  Else
85                      sSQL = "INSERT INTO FUAsuransi (Aktifitas,Nama,Alamat,Telp,HP1,HP2,NoRangka,Model," & _
                               "NoPolisi,MasaberlakuAwal,MasaberlakuAkhir," & _
                               "DealerPenjual,Terhubung,Deal,Alasan,TglFU,userFU,kodeReq) VALUES " & _
                               "('" & NulSID("upgrade") & "','" & s2d(Trim(txt(0).Text)) & "','" & s2d(Trim(txt(1).Text)) & "','" & s2d(Trim(txt(2).Text)) & "','" & Trim(txt(3).Text) & "','" & s2d(Trim(txt(4).Text)) & "','" & Trim(txt(5).Text) & "','" & s2d(NulSID(rsD!modelkategory)) & "', " & _
                               "'" & IIf(Trim(txt(8).Text) <> "", Trim(txt(8).Text) & "-" & Trim(txt(9).Text) & "-" & Trim(txt(10).Text), "") & "','" & Format(rsD("masaberlakuawal"), "YYYY-MM-DD") & "','" & Format(rsD("masaberlakuakhir"), "YYYY-MM-DD") & "', " & _
                               "'" & Trim(Left(rsD("DealerPenjual"), 4)) & "','" & IIf(Opt(0).value = True, 1, 0) & "','" & IIf(Opt(2).value = True, 1, 0) & "','" & Trim(s2d(txt(11).Text)) & "',Getdate(),'" & Trim(UserLog) & "', '" & Trim(KodeReQ) & "')"
86                  End If
87                  AdaBegin = False: KetEro = ""
88                  Call objHT.OpenTransactionBengkel
89                  Call objHT.OpenTransactionHO
90                  AdaBegin = True
91                  Call objHT.InsertData_Bengkel(sSQL)
92                  sSQL = ""
93                  If Opt(2).value = True Then
94                      sSQL = "INSERT INTO OSS..Asuransi (NoSPK,NamaPemesan,Alamat,Kota,telp_rumah,HP1,HP2," & _
                               "NamaTertanggung,telp_rumahTT,AlamatTT,KotaTT,HP1TT,HP2TT," & _
                               "Kode_type,Kode_warna,NamaSales,noRangka,noMesin,thnUnit," & _
                               "kodeDealer,banjir,Gempa,SRCC,Teroris," & _
                               "TanggungP3,CelakaDiri,userEntry,tglEntry," & _
                               "areaAsuransi,tenorAsuransi,KodeAsuransi,JenisAsuransi,JenisAsuransi2,JenisAsuransi3,JenisAsuransi4,JenisAsuransi5,JenisAsuransi6,JenisAsuransi7,JenisAsuransi8,CelakaDiriPass) " & _
                               "VALUES ('" & NulSID(KodeReQ) & "','" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(txt(1).Text))) & "','','" & Trim(s2d(Nul(txt(2).Text))) & "','" & Trim(s2d(Nul(txt(3).Text))) & "','" & Trim(s2d(Nul(txt(4).Text))) & "'," & _
                               "'" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(txt(2).Text))) & "','" & Trim(s2d(Nul(rsD("AlamatTT")))) & "','" & Trim(s2d(Nul(rsD("KotaTT")))) & "','" & Trim(s2d(Nul(rsD("HP1TT")))) & "','" & Trim(s2d(Nul(rsD("HP2TT")))) & "'," & _
                               "'" & Trim(s2d(Nul(rsD("kodekategori")))) & "','" & Left(Trim(s2d(Nul(rsD("warna")))), 7) & "','" & Trim(s2d(Nul(rsD("namasales")))) & "','" & Trim(s2d(Nul(rsD("norangka")))) & "','" & Trim(s2d(Nul(rsD("nomesin")))) & "','" & Trim(s2d(Nul(rsD("kodetahun")))) & "'," & _
                               "'" & Left(Trim(s2d(Nul(rsD("DealerPenjual")))), 4) & "','" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & _
                               "" & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & ",'" & Trim(gSA) & "',GETDATE(), " & _
                               "'" & Trim(DcAsuransi.Columns(1).value) & "','" & CInt(cboTenor.Text) & "','" & Trim(DcAsuransi.Columns(2).value) & "','" & isXs(cbojns(0).Text) & "','" & isXs(cbojns(1).Text) & "','" & isXs(cbojns(2).Text) & "','" & isXs(cbojns(3).Text) & "','" & isXs(cbojns(4).Text) & "','" & isXs(cbojns(5).Text) & "','" & isXs(cbojns(6).Text) & "','" & isXs(cbojns(7).Text) & "'," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ")"
95                        Call objHT.InsertData_Ho(sSQL)
                          
96                        sSQL = "INSERT INTO [OSS].[dbo].[UpgradeAsuransi]([KodeReq],[NamaPemesan],[Alamat],[telp_rumah],[HP1],[HP2],[NamaDebitur],[AlamatDebitur],[kotaDebitur],[telp_rumahDebitur],[HP1Debitur],[HP2Debitur],[Kode_type]" & _
                                 " ,[Kode_warna],[NamaSales],[noHPSales],[noRangka],[noMesin],[thnUnit],[kodeDealer],[noSpp],[noPO],[userDown],[tglDown],[namaLeasing],[wilayahLeasing],[KodeAsuransi],[jenisAsuransi]" & _
                                 " ,[jenisAsuransi2],[jenisAsuransi3],[jenisAsuransi4],[jenisAsuransi5],[jenisAsuransi6],[jenisAsuransi7],[jenisAsuransi8],[Banjir],[Gempa],[SRCC],[Teroris],[TanggungP3],[CelakaDiri]" & _
                                 " ,[CelakaDiriPass],[tenorAsuransi],[hargOTR],[userEntry],[tglEntry],[UserUploadLsg],[TglUploadLsg],[UserUploadAsr],[TglUploadAsr],[UserProsesLsg]" & _
                                 " ,[TglProsesLsg],[UserProsesAsr],[TglProsesAsr],[UserProsesNIS],[TglProsesNIS],[statusAsuransi],[Kota],[nospk]) VALUES ('" & KodeReQ & "','" & Trim(s2d(Nul(txt(0).Text))) & "'" & _
                                 " ,'" & Trim(s2d(Nul(txt(1).Text))) & "','','" & Trim(s2d(Nul(txt(3).Text))) & "','" & Trim(s2d(Nul(txt(4).Text))) & "','" & Left(NulSID(txt(0)), 40) & "','" & Left(NulSID(txt(1)), 100) & "'" & _
                                 " ,'','','" & Left(NulSID(txt(3)), 15) & "','" & Left(NulSID(txt(4)), 15) & "','" & Left(NulSID(rsD!namakategory), 50) & "','" & Left(NulSID(rsD!warna), 7) & "','" & Left(NulSID(rsD!namasales), 50) & "'" & _
                                 " ,'','" & Left(NulSID(rsD!noRangka), 25) & "','" & Left(NulSID(rsD!noMesin), 15) & "','" & Trim(s2d(Nul(rsD("kodetahun")))) & "','" & Log.kodeCabang & "','','" & Left(NulSID(KodeReQ), 30) & "'" & _
                                 " ,'','01-01-1900','','','" & Trim(DcAsuransi.Columns(2).value) & "','" & isXs(cbojns(0).Text) & "'" & _
                                 " ,'" & isXs(cbojns(1).Text) & "','" & isXs(cbojns(2).Text) & "','" & isXs(cbojns(3).Text) & "','" & isXs(cbojns(4).Text) & "','" & isXs(cbojns(5).Text) & "','" & isXs(cbojns(6).Text) & "'" & _
                                 " ,'" & isXs(cbojns(7).Text) & "','" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "" & _
                                 " ," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ",'" & CInt(cboTenor.Text) & "',0,'" & gSA & "',getdate(),''" & _
                                 " ,'1900-01-01','','1900-01-01','','1900-01-01',''" & _
                                 " ,'1900-01-01','','1900-01-01','','','" & NulSID(rsD!NoSPK) & "')"
97                        Call objHT.InsertData_Ho(sSQL)
                          
98                        sSQL = "INSERT INTO [" & dbReplikasiHO & "]..SendMail_Hansarenkei "
99                        sSQL = sSQL & " (KodeReq,Aktifitas,KodeDealer,NamaDealer,Perusahaan,NamaCustomer,MDL,FullModel,"
100                       sSQL = sSQL & " NoRangka,Tujuan,CC,"
101                       sSQL = sSQL & " Pengirim,UserEntry,TglEntry,AppEntry,Tipe) values ("
102                       sSQL = sSQL & " '" & NulSID(KodeReQ) & "','UPGRADE','" & Right(gDb, 2) & "','" & NulSID(Log.LogNamaNas) & "','" & NulSID(DcAsuransi.Columns(3).value) & "','" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(rsD("kodekategori")))) & "','" & Trim(s2d(Nul(txt(7).Text))) & "',"
103                       sSQL = sSQL & " '" & Trim(s2d(Nul(rsD("norangka")))) & "','" & toEmail & "','" & ccEmail & "',"
104                       sSQL = sSQL & " '" & pengirimEmail & "','" & gSA & "',GETDATE(),'SVC','Asuransi')"
105                       Call objHT.InsertData_Bengkel(sSQL)
106                 Else
                        
107                 End If
108                 Call objHT.CloseTransactionBengkel
109                 Call objHT.CloseTransactionHO
110                 AdaBegin = False
111                 MsgBox "Data Berhasil diproses!", vbInformation
112                 Unload Me
113             Case NulSID("Renew Customer")
114                 KodeReQ = "": Set rso = Nothing: Set rso = New ADODB.Recordset: rsD.MoveFirst
115                 sSQL = "SELECT TOP 1 KodeReq FROM FUAsuransi With (NoLock) WHERE SUBSTRING(KodeReq,1,11) = '" & Left(Trim(rsD("DealerPenjual")), 4) & "-RW" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & "' ORDER BY KODEREQ DESC" '2020-NW2010
116                 Set rso = objO.FetchData(sSQL)
117                 If rso.RecordCount > 0 Then
118                     KodeReQ = Left(Trim(rsD("DealerPenjual")), 4) & "-RW" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & Format(Right(rso("kodereq"), 4) + 1, "0000")
119                 Else
120                     KodeReQ = Left(Trim(rsD("DealerPenjual")), 4) & "-RW" & Right(Format(year(isnow), "0000"), 2) & Format(Month(isnow), "00") & "0001"
121                 End If
122                 If Opt(2).value = True Then
123                     sSQL = "INSERT INTO FUAsuransi (Aktifitas,Nama,Alamat,Telp,HP1,HP2,NoRangka,Model, " & _
                               "NoPolisi,NamaLeasing,KodeAsuransi,MasaberlakuAwal,MasaberlakuAkhir,Tenor, " & _
                               "JenisAsuransi1,JenisAsuransi2,JenisAsuransi3,JenisAsuransi4,JenisAsuransi5,JenisAsuransi6,JenisAsuransi7, " & _
                               "JenisAsuransi8,DealerPenjual,Terhubung,Deal,Alasan,TglFU,userFU,kodeReq, " & _
                               "Banjir,Gempa,SRCC,Teroris,TanggungP3,CelakaDiri,CelakaDiriPass) VALUES " & _
                               "('" & NulSID("Renew") & "','" & s2d(Trim(txt(0).Text)) & "','" & s2d(Trim(txt(1).Text)) & "','" & s2d(Trim(txt(2).Text)) & "','" & s2d(Trim(txt(3).Text)) & "','" & s2d(Trim(txt(4).Text)) & "','" & Trim(txt(5).Text) & "','" & s2d(NulSID(rsD!modelkategory)) & "', " & _
                               "'" & NulSID(rsD!NoPolisi) & "','" & Trim(rsD("kode_leasing")) & "'," & CInt(DcAsuransi.Columns(2).value) & ",'" & Format(rsD("masaberlakuawal"), "YYYY-MM-DD") & "','" & Format(rsD("masaberlakuakhir"), "YYYY-MM-DD") & "'," & CInt(cboTenor.Text) & ", " & _
                               "'" & s2d(Trim(isAs(cbojns(0)))) & "','" & s2d(Trim(isAs(cbojns(1)))) & "','" & s2d(Trim(isAs(cbojns(2)))) & "','" & s2d(Trim(isAs(cbojns(3)))) & "','" & isAs(s2d(Trim(cbojns(4)))) & "','" & isAs(s2d(Trim(cbojns(5)))) & "','" & s2d(Trim(isAs(cbojns(6)))) & "', " & _
                               "'" & s2d(Trim(isAs(cbojns(7)))) & "','" & Trim(Left(rsD("DealerPenjual"), 4)) & "','" & IIf(Opt(0).value = True, 1, 0) & "','" & IIf(Opt(2).value = True, 1, 0) & "','" & Trim(s2d(txt(11).Text)) & "',Getdate(),'" & Trim(gSA) & "', '" & Trim(KodeReQ) & "', " & _
                               "'" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & "," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ")"
124                 Else
125                     sSQL = "INSERT INTO FUAsuransi (Aktifitas,Nama,Alamat,Telp,HP1,HP2,NoRangka,Model," & _
                               "NoPolisi,MasaberlakuAwal,MasaberlakuAkhir," & _
                               "DealerPenjual,Terhubung,Deal,Alasan,TglFU,userFU,kodeReq) VALUES " & _
                               "('" & NulSID("Renew") & "','" & s2d(Trim(txt(0).Text)) & "','" & s2d(Trim(txt(1).Text)) & "','" & s2d(Trim(txt(2).Text)) & "','" & Trim(txt(3).Text) & "','" & s2d(Trim(txt(4).Text)) & "','" & Trim(txt(5).Text) & "','" & s2d(NulSID(rsD!modelkategory)) & "', " & _
                               "'" & IIf(Trim(txt(8).Text) <> "", Trim(txt(8).Text) & "-" & Trim(txt(9).Text) & "-" & Trim(txt(10).Text), "") & "','" & Format(rsD("masaberlakuawal"), "YYYY-MM-DD") & "','" & Format(rsD("masaberlakuakhir"), "YYYY-MM-DD") & "', " & _
                               "'" & Trim(Left(rsD("DealerPenjual"), 4)) & "','" & IIf(Opt(0).value = True, 1, 0) & "','" & IIf(Opt(2).value = True, 1, 0) & "','" & Trim(s2d(txt(11).Text)) & "',Getdate(),'" & Trim(UserLog) & "', '" & Trim(KodeReQ) & "')"
126                 End If
127                 AdaBegin = False: KetEro = ""
128                 Call objHT.OpenTransactionBengkel
129                 Call objHT.OpenTransactionHO
130                 AdaBegin = True
131                 Call objHT.InsertData_Bengkel(sSQL)
132                 sSQL = ""
133                 If Opt(2).value = True Then
134                     sSQL = "INSERT INTO OSS..Asuransi (NoSPK,NamaPemesan,Alamat,Kota,telp_rumah,HP1,HP2," & _
                               "NamaTertanggung,telp_rumahTT,AlamatTT,KotaTT,HP1TT,HP2TT," & _
                               "Kode_type,Kode_warna,NamaSales,noRangka,noMesin,thnUnit," & _
                               "kodeDealer,banjir,Gempa,SRCC,Teroris," & _
                               "TanggungP3,CelakaDiri,userEntry,tglEntry," & _
                               "areaAsuransi,tenorAsuransi,KodeAsuransi,JenisAsuransi,JenisAsuransi2,JenisAsuransi3,JenisAsuransi4,JenisAsuransi5,JenisAsuransi6,JenisAsuransi7,JenisAsuransi8,CelakaDiriPass) " & _
                               "VALUES ('" & NulSID(KodeReQ) & "','" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(txt(1).Text))) & "','','" & Trim(s2d(Nul(txt(2).Text))) & "','" & Trim(s2d(Nul(txt(3).Text))) & "','" & Trim(s2d(Nul(txt(4).Text))) & "'," & _
                               "'" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(txt(2).Text))) & "','" & Trim(s2d(Nul(rsD("AlamatTT")))) & "','" & Trim(s2d(Nul(rsD("KotaTT")))) & "','" & Trim(s2d(Nul(rsD("HP1TT")))) & "','" & Trim(s2d(Nul(rsD("HP2TT")))) & "'," & _
                               "'" & Trim(s2d(Nul(rsD("kodekategori")))) & "','" & Left(Trim(s2d(Nul(rsD("warna")))), 7) & "','" & Trim(s2d(Nul(rsD("namasales")))) & "','" & Trim(s2d(Nul(rsD("norangka")))) & "','" & Trim(s2d(Nul(rsD("nomesin")))) & "','" & Trim(s2d(Nul(rsD("kodetahun")))) & "'," & _
                               "'" & Left(Trim(s2d(Nul(rsD("DealerPenjual")))), 4) & "','" & Ch(0).value & "','" & Ch(2).value & "','" & Ch(1).value & "','" & Ch(3).value & "'," & _
                               "" & IIf(UCase(cboPos(0).Text) = "--PILIH--", 0, Replace(cboPos(0).Text, ".", "")) & "," & IIf(UCase(cboPos(1).Text) = "--PILIH--", 0, Replace(cboPos(1).Text, ".", "")) & ",'" & Trim(gSA) & "',GETDATE(), " & _
                               "'" & Trim(DcAsuransi.Columns(1).value) & "','" & CInt(cboTenor.Text) & "','" & Trim(DcAsuransi.Columns(2).value) & "','" & isXs(cbojns(0).Text) & "','" & isXs(cbojns(1).Text) & "','" & isXs(cbojns(2).Text) & "','" & isXs(cbojns(3).Text) & "','" & isXs(cbojns(4).Text) & "','" & isXs(cbojns(5).Text) & "','" & isXs(cbojns(6).Text) & "','" & isXs(cbojns(7).Text) & "'," & IIf(UCase(cboPos(2).Text) = "--PILIH--", 0, Replace(cboPos(2).Text, ".", "")) & ")"
135                       Call objHT.InsertData_Ho(sSQL)
                          
136                       sSQL = "INSERT INTO [" & dbReplikasiHO & "]..SendMail_Hansarenkei "
137                       sSQL = sSQL & " (KodeReq,Aktifitas,KodeDealer,NamaDealer,Perusahaan,NamaCustomer,MDL,FullModel,"
138                       sSQL = sSQL & " NoRangka,Tujuan,CC,"
139                       sSQL = sSQL & " Pengirim,UserEntry,TglEntry,AppEntry,Tipe) values ("
140                       sSQL = sSQL & " '" & NulSID(KodeReQ) & "','RENEW','" & Right(gDb, 2) & "','" & NulSID(Log.LogNamaNas) & "','" & NulSID(DcAsuransi.Columns(3).value) & "','" & Trim(s2d(Nul(txt(0).Text))) & "','" & Trim(s2d(Nul(rsD("kodekategori")))) & "','" & Trim(s2d(Nul(txt(7).Text))) & "',"
141                       sSQL = sSQL & " '" & Trim(s2d(Nul(rsD("norangka")))) & "','" & toEmail & "','" & ccEmail & "',"
142                       sSQL = sSQL & " '" & pengirimEmail & "','" & gSA & "',GETDATE(),'SVC','Asuransi')"
143                       Call objHT.InsertData_Bengkel(sSQL)
144                 Else
                        
145                 End If
146                 Call objHT.CloseTransactionBengkel
147                 Call objHT.CloseTransactionHO
148                 AdaBegin = False
149                 MsgBox "Data Berhasil diproses!", vbInformation
150                 Unload Me
151         End Select
152     Else
153           MsgBox "Data Yang Anda Inputkan Belum Lengkap !" & vbCrLf & isFalse, vbInformation
154     End If
155     Exit Sub
Erro:
156     If AdaBegin Then
157       Call objHT.CloseTransactionBengkel_err
158       Call objHT.CloseTransactionHO_err
159     End If
       
160       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSave_Click of Form frmRetensiAsuransiCall"

End Sub

Private Function isAs(vux As String) As String
    isAs = ""
    isAs = Replace(vux, "--pilih--", "")
End Function

Private Function isXs(vue As String) As String
    isXs = ""
    Select Case UCase(vue)
        Case "--PILIH--"
            isXs = "0"
        Case "ALL RISK"
            isXs = "2"
        Case "TOTAL LOST ONLY"
            isXs = "1"
    End Select
End Function

Private Function isValid() As Boolean
  isFalse = ""
  isValid = True
  If Opt(1).value = True Then         'Tidak Terhubung
      If txt(11).Text = "" Or Len(txt(11).Text) = 0 Then
          isValid = False
          isFalse = "[Alasan]"
          Exit Function
      End If
  ElseIf Opt(0).value = True Then     'Terhubung
      If Opt(3).value = True Then     'Tidak Deal
          If txt(11).Text = "" Or Len(txt(11).Text) = 0 Then
              isValid = False
              isFalse = "[Alasan]"
              Exit Function
          End If
      ElseIf Opt(2).value = True Then 'Deal
          If cboTenor.Text = "--pilih--" Then
              isValid = False
              isFalse = "[Tenor]"
              Exit Function
          End If
          If cboTenor.Text <> "--pilih--" Then
              For C = 0 To (CInt(cboTenor.Text) - 1)
                  If cbojns(C).Text = "--pilih--" Then
                      isValid = False
                      isFalse = "[Jenis Asuransi Tahun " & C + 1 & "]"
                      Exit Function
                  End If
              Next C
              For C = (CInt(cboTenor.Text)) To 7
                  If cbojns(C).Text <> "--pilih--" Then
                      isValid = False
                      isFalse = "[Jenis Asuransi Tahun " & C + 1 & " melebihi tenor]"
                      Exit Function
                  End If
              Next C
          End If
          '20201123 req mb thania untuk menutup validasi
'          If cboPos(0).Text = "--pilih--" Then
'              isValid = False
'              isFalse = "[Pihak ke-3]"
'              Exit Function
'          End If
'          If cboPos(1).Text = "--pilih--" Then
'              isValid = False
'              isFalse = "[Kecelakaan diri Driver]"
'              Exit Function
'          End If
'          If cboPos(2).Text = "--pilih--" Then
'              isValid = False
'              isFalse = "[Kecelakaan diri Penumpang]"
'              Exit Function
'          End If
          If DcAsuransi.Text = "" Then
              isValid = False
              isFalse = "[Asuransi]"
          End If
      Else
          If Opt(2).value = False And Opt(3).value = False Then
              isValid = False
              isFalse = "[Deal or Tidak]"
          End If
      End If
  End If
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub DcAsuransi_Click()
    'MsgBox DcAsuransi.Columns(0).value & " | " & DcAsuransi.Columns(1).value & " | " & DcAsuransi.Columns(2).value
End Sub

Private Sub Form_Load()
    Me.Caption = TitleBP
   ' Move 3240, objA.MyTop + 385
    Me.Top = 0
    Move 0, 0
    Me.Width = Layar.Width - 600
    Me.Height = Layar.Height - 800
    Call clearAll
    If idAsuransi > 0 Then
        'DcAsuransi.Text = objO.FetchDataHO_String("SELECT NamaAsuransi + ' ' + WilayahAsuransi as NamaAsuransi, WilayahAsuransi, KodeAsuransi FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 AND KODEASURANSI = '" & idAsuransi & "'")
        DcAsuransi.value = objO.FetchDataHO_String("SELECT NamaAsuransi + ' ' + WilayahAsuransi as NamaAsuransi, WilayahAsuransi, KodeAsuransi FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 AND KODEASURANSI = '" & idAsuransi & "'")
    End If
    FrTenor.Visible = False
    FrJenis.Visible = False
    FrThird.Visible = False
    Opt(2).Enabled = False
    Opt(3).Enabled = False
    If rsD.RecordCount > 0 Then
        rsD.MoveFirst
        Call BindToCust
    End If
End Sub

Private Sub BindToCust()
    txt(0).Text = Trim(Nul(rsD("nama_pemesan")))
    txt(1).Text = Trim(Nul(rsD("Alamat")))
    'txt(2).Text = Trim(Nul(rsd("telepon")))
    txt(3).Text = Trim(Nul(rsD("nohp1")))
    txt(4).Text = Trim(Nul(rsD("nohp2")))
    txt(5).Text = Trim(Nul(rsD("norangka")))
    txt(6).Text = Trim(Nul(rsD("nomesin")))
    txt(7).Text = Trim(Nul(rsD("namakategory")))
    txt(12).Text = Trim(Nul(rsD("warna")))
'    If Trim(Nul(rsd("nopolisi"))) <> "" Then
'        Call pecahNoPol(Trim(Nul(rsd("nopolisi"))), nDpn, nTgh, nBlk)
'        txt(8).Text = nDpn
'        txt(9).Text = nTgh
'        txt(10).Text = nBlk
'    Else
'        txt(8).Text = ""
'        txt(9).Text = ""
'        txt(10).Text = ""
'    End If
End Sub

Private Sub clearAll()
  For C = 0 To 9
      txt(C).Text = ""
  Next C
  For C = 0 To 3
      Opt(C).value = False
      Ch(C).value = 0
  Next C
  For C = 0 To 7
      cbojns(C).Clear
      cbojns(C).AddItem "--pilih--"            '(1)
      cbojns(C).AddItem "All Risk"            '(1)
      cbojns(C).AddItem "Total Lost Only"     '(2)
      cbojns(C).Text = cbojns(C).LIST(0)
  Next C
  With cboTenor
      .Clear
      .AddItem "--pilih--"
      .AddItem "1"
      .AddItem "2"
      .AddItem "3"
      .AddItem "4"
      .AddItem "5"
      .AddItem "6"
      .AddItem "7"
      .AddItem "8"
      .Text = .LIST(0)
  End With
    
    
  sSQL = "": sSQL = "select DISTINCT pertanggungan from simulasi_kredit..M_ThirdPartyLiability order by pertanggungan"
  Set Tri = Nothing: Set Tri = New ADODB.Recordset
  Set Tri = objO.FetchDataHO(sSQL)
  If Tri.RecordCount > 0 Then
      Tri.MoveFirst
      With cboPos(0)
          .Clear
          .AddItem "--pilih--"
          Do Until Tri.EOF
              .AddItem Format(Nul(Tri("pertanggungan")), "###,###,###,###")
              Tri.MoveNext
          Loop
          .Text = .LIST(0)
      End With
  End If
    
  sSQL = "": sSQL = "select DISTINCT pertanggungan from simulasi_kredit..M_PersonalAccident order by pertanggungan"
  Set Tri = Nothing: Set Tri = New ADODB.Recordset
  Set Tri = objO.FetchDataHO(sSQL)
  If Tri.RecordCount > 0 Then
      Tri.MoveFirst
      With cboPos(1)
          .Clear
          .AddItem "--pilih--"
          Do Until Tri.EOF
              .AddItem Format(Nul(Tri("pertanggungan")), "###,###,###,###")
              Tri.MoveNext
          Loop
          .Text = .LIST(0)
      End With
      Tri.MoveFirst
      With cboPos(2)
          .Clear
          .AddItem "--pilih--"
          Do Until Tri.EOF
              .AddItem Format(Nul(Tri("pertanggungan")), "###,###,###,###")
              Tri.MoveNext
          Loop
          .Text = .LIST(0)
      End With
  End If
    
  AdoAs.ConnectionString = SQLConBengkel99
  AdoAs.RecordSource = "SELECT NamaAsuransi + ' ' + WilayahAsuransi AS NamaAsuransi, WilayahAsuransi, KodeAsuransi,NamaAsuransi as nama FROM oss..mstAsuransi With (Nolock) WHERE oss = 1 ORDER BY NamaAsuransi,WilayahAsuransi "

  DcAsuransi.DataSourceList = AdoAs
  'DcAsuransi.DataFieldToDisplay = "NamaAsuransi"
  'DcAsuransi.DataFieldList = "KodeAsuransi"
  DcAsuransi.DataFieldList = "NamaAsuransi"
  DcAsuransi.Columns(0).Width = 4200
  DcAsuransi.Columns(1).Width = 3000
  DcAsuransi.Columns(2).Width = 1000
  DcAsuransi.Columns(3).Width = 1000
  DcAsuransi.Columns(0).Caption = "Nama Asuransi"
  DcAsuransi.Columns(1).Caption = "Wilayah Asuransi"
  DcAsuransi.Columns(2).Caption = "Kode"
  DcAsuransi.Columns(3).Visible = False

End Sub



Private Sub Opt_Click(Index As Integer)
    Select Case Index
        Case 0
            If Opt(Index).value = True Then
                Opt(2).Enabled = True
                Opt(3).Enabled = True
                If Opt(2).value = True Then
                    FrTenor.Visible = True
                    FrJenis.Visible = True
                    FrThird.Visible = True
                End If
                txt(11).Text = ""
            End If
        Case 1
            If Opt(Index).value = True Then
                Opt(2).Enabled = False
                Opt(3).Enabled = False
                FrTenor.Visible = False
                FrJenis.Visible = False
                FrThird.Visible = False
            End If
        Case 2
            If Opt(Index).value = True Then
                FrTenor.Visible = True
                FrJenis.Visible = True
                FrThird.Visible = True
                txt(11).Text = ""
            End If
        Case 3
            If Opt(Index).value = True Then
                FrTenor.Visible = False
                FrJenis.Visible = False
                FrThird.Visible = False
                txt(11).SetFocus
            End If
    End Select
End Sub
