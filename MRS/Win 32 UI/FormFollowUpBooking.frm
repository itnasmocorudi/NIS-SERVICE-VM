VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FormFollowUpBooking 
   BorderStyle     =   0  'None
   Caption         =   "ing ::."
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormFollowUpBooking.frx":0000
   ScaleHeight     =   8370
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScrollFUB 
      Height          =   8350
      LargeChange     =   400
      Left            =   11710
      Max             =   8800
      SmallChange     =   350
      TabIndex        =   14
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   480
      ScaleHeight     =   8655
      ScaleWidth      =   11295
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin TabDlg.SSTab SSFUB 
         Height          =   17355
         Left            =   -120
         TabIndex        =   1
         Top             =   -220
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   30612
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         BackColor       =   12577002
         TabCaption(0)   =   " "
         TabPicture(0)   =   "FormFollowUpBooking.frx":EFAF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shape4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblPart"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Shape5"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label19"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Shape2"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lblJml"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label3"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Line1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label2"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Line2"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblKet"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Line3"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Line4"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Shape1"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "ImgView"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label4"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Line5"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Line6"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Line7"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Shape3"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Line10"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Line11"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Line12"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "ShapeFu"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Shape7"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Line14"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Line16"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Line17"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Line18"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Line9"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Line20"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Line21"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Shape6"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "FrameKendaraan"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "FrPros"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "GridBookPart"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "GridFollowBook"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "DtRemain"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "ImageList"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "cboType"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "cmdExit"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "cmdProcess"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "CmdPros"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "Command1"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "Frame4"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "Frame3"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "Frame10"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "Frame9"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "Frame1"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "cmdTop"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "chAll"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "CmdXls"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).ControlCount=   53
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
            Picture         =   "FormFollowUpBooking.frx":EFCB
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox chAll 
            Caption         =   "Display All Booking"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   128
            Top             =   1400
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.CommandButton cmdTop 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Back To Top"
            Height          =   735
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FormFollowUpBooking.frx":F447
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   16340
            Width           =   11175
         End
         Begin VB.Frame Frame1 
            Caption         =   " BPK "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   240
            TabIndex        =   117
            Top             =   15000
            Width           =   4215
            Begin TDBDate6Ctl.TDBDate DTDO 
               Height          =   330
               Left            =   1680
               TabIndex        =   118
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormFollowUpBooking.frx":F7D2
               Caption         =   "FormFollowUpBooking.frx":F8FE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormFollowUpBooking.frx":F96A
               Keys            =   "FormFollowUpBooking.frx":F988
               Spin            =   "FormFollowUpBooking.frx":F9E6
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
               Enabled         =   0
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
            Begin TDBDate6Ctl.TDBDate DTDEC 
               Height          =   330
               Left            =   1680
               TabIndex        =   122
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormFollowUpBooking.frx":FA0E
               Caption         =   "FormFollowUpBooking.frx":FB3A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormFollowUpBooking.frx":FBA6
               Keys            =   "FormFollowUpBooking.frx":FBC4
               Spin            =   "FormFollowUpBooking.frx":FC22
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
               Enabled         =   0
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
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl DEC"
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
               Left            =   240
               TabIndex        =   123
               Top             =   720
               Width           =   765
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl DO / BPK"
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
               Left            =   240
               TabIndex        =   121
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label16 
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
               Left            =   1560
               TabIndex        =   120
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label15 
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
               Left            =   1560
               TabIndex        =   119
               Top             =   735
               Width           =   75
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   " CR 1 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   4440
            TabIndex        =   110
            Top             =   15000
            Width           =   3495
            Begin VB.TextBox txtKMCR1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   111
               Top             =   720
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate DTCR1 
               Height          =   330
               Left            =   1680
               TabIndex        =   112
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormFollowUpBooking.frx":FC4A
               Caption         =   "FormFollowUpBooking.frx":FD76
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormFollowUpBooking.frx":FDE2
               Keys            =   "FormFollowUpBooking.frx":FE00
               Spin            =   "FormFollowUpBooking.frx":FE5E
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
               Enabled         =   0
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
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Validasi"
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
               Left            =   240
               TabIndex        =   116
               Top             =   360
               Width           =   1110
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angka KM"
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
               Left            =   240
               TabIndex        =   115
               Top             =   720
               Width           =   960
            End
            Begin VB.Label Label12 
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
               Left            =   1440
               TabIndex        =   114
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label11 
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
               Left            =   1440
               TabIndex        =   113
               Top             =   735
               Width           =   75
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   " CR 2 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   7920
            TabIndex        =   103
            Top             =   15000
            Width           =   3495
            Begin VB.TextBox txtKMCR2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   104
               Top             =   720
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate DTCR2 
               Height          =   330
               Left            =   1680
               TabIndex        =   105
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormFollowUpBooking.frx":FE86
               Caption         =   "FormFollowUpBooking.frx":FFB2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormFollowUpBooking.frx":1001E
               Keys            =   "FormFollowUpBooking.frx":1003C
               Spin            =   "FormFollowUpBooking.frx":1009A
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
               Enabled         =   0
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
            Begin VB.Label Label10 
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
               Left            =   1440
               TabIndex        =   109
               Top             =   735
               Width           =   75
            End
            Begin VB.Label Label9 
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
               Left            =   1440
               TabIndex        =   108
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angka KM"
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
               Left            =   240
               TabIndex        =   107
               Top             =   720
               Width           =   960
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Validasi"
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
               Left            =   240
               TabIndex        =   106
               Top             =   360
               Width           =   1110
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " ALAMAT SURAT MENYURAT "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2175
            Left            =   240
            TabIndex        =   54
            Top             =   9000
            Width           =   11175
            Begin VB.TextBox txtPropinsiCust2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4800
               MaxLength       =   25
               TabIndex        =   62
               Top             =   1680
               Width           =   2055
            End
            Begin VB.TextBox txtKotaCust2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   25
               TabIndex        =   61
               Top             =   1680
               Width           =   1935
            End
            Begin VB.TextBox txtKodePosCust2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   5
               TabIndex        =   60
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox txtAlamatCust2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   59
               Top             =   720
               Width           =   5175
            End
            Begin VB.TextBox txtNamaCust2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   58
               Top             =   360
               Width           =   3255
            End
            Begin VB.TextBox txtTelpCDMA 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8520
               MaxLength       =   15
               TabIndex        =   57
               Top             =   1680
               Width           =   2295
            End
            Begin VB.TextBox txtTelpGSM 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8520
               MaxLength       =   15
               TabIndex        =   56
               Top             =   1320
               Width           =   2295
            End
            Begin VB.TextBox txtTelpRumah 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8520
               MaxLength       =   15
               TabIndex        =   55
               Top             =   960
               Width           =   2295
            End
            Begin VB.Line Line8 
               X1              =   7440
               X2              =   10800
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label Label57 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
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
               Left            =   4680
               TabIndex        =   80
               Top             =   1680
               Width           =   60
            End
            Begin VB.Label Label56 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Propinsi"
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
               Left            =   3840
               TabIndex        =   79
               Top             =   1680
               Width           =   780
            End
            Begin VB.Label Label51 
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
               Left            =   1560
               TabIndex        =   78
               Top             =   1320
               Width           =   75
            End
            Begin VB.Label Label49 
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
               Left            =   1560
               TabIndex        =   77
               Top             =   1680
               Width           =   75
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "KodePos"
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
               Left            =   240
               TabIndex        =   76
               Top             =   1320
               Width           =   840
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kota"
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
               Left            =   240
               TabIndex        =   75
               Top             =   1680
               Width           =   465
            End
            Begin VB.Label Label46 
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
               Left            =   1560
               TabIndex        =   74
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label44 
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
               Left            =   1560
               TabIndex        =   73
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama "
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
               Left            =   240
               TabIndex        =   72
               Top             =   360
               Width           =   585
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat"
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
               Left            =   240
               TabIndex        =   71
               Top             =   720
               Width           =   690
            End
            Begin VB.Label Label32 
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
               Left            =   8280
               TabIndex        =   70
               Top             =   600
               Width           =   75
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Telepon"
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
               Left            =   7440
               TabIndex        =   69
               Top             =   600
               Width           =   795
            End
            Begin VB.Label Label29 
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
               Left            =   8280
               TabIndex        =   68
               Top             =   960
               Width           =   75
            End
            Begin VB.Label Label28 
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
               Left            =   8280
               TabIndex        =   67
               Top             =   1320
               Width           =   75
            End
            Begin VB.Label Label27 
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
               Left            =   8280
               TabIndex        =   66
               Top             =   1680
               Width           =   75
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rumah"
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
               Left            =   7440
               TabIndex        =   65
               Top             =   960
               Width           =   660
            End
            Begin VB.Label Label23 
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
               Left            =   7440
               TabIndex        =   64
               Top             =   1320
               Width           =   435
            End
            Begin VB.Label Label22 
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
               Left            =   7440
               TabIndex        =   63
               Top             =   1680
               Width           =   600
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " FAKTUR DATA STNK "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2055
            Left            =   240
            TabIndex        =   28
            Top             =   11160
            Width           =   11175
            Begin VB.TextBox txtNamaSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   36
               Top             =   360
               Width           =   4455
            End
            Begin VB.TextBox txtAlamatSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   1680
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   35
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox txtKotaSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   34
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtKodePosSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4920
               MaxLength       =   30
               TabIndex        =   33
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox txtTempatLahir 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   32
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtPekerjaan 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   31
               Top             =   1080
               Width           =   3135
            End
            Begin VB.TextBox txtHoby 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   30
               Top             =   1440
               Width           =   3135
            End
            Begin VB.TextBox txtAgama 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   29
               Top             =   720
               Width           =   1695
            End
            Begin TDBDate6Ctl.TDBDate DTLahir 
               Height          =   330
               Left            =   9600
               TabIndex        =   37
               Top             =   360
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   582
               Calendar        =   "FormFollowUpBooking.frx":100C2
               Caption         =   "FormFollowUpBooking.frx":101EE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormFollowUpBooking.frx":1025A
               Keys            =   "FormFollowUpBooking.frx":10278
               Spin            =   "FormFollowUpBooking.frx":102D6
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
               Enabled         =   0
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
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama STNK"
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
               Left            =   240
               TabIndex        =   53
               Top             =   360
               Width           =   1080
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat "
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
               Left            =   240
               TabIndex        =   52
               Top             =   720
               Width           =   750
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kota"
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
               Left            =   240
               TabIndex        =   51
               Top             =   1440
               Width           =   465
            End
            Begin VB.Label Label36 
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
               Left            =   1560
               TabIndex        =   50
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label37 
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
               Left            =   1560
               TabIndex        =   49
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Pos   :"
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
               Left            =   3720
               TabIndex        =   48
               Top             =   1440
               Width           =   1140
            End
            Begin VB.Label Label39 
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
               Left            =   1560
               TabIndex        =   47
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pekerjaan"
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
               Left            =   6480
               TabIndex        =   46
               Top             =   1080
               Width           =   930
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agama"
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
               Left            =   6480
               TabIndex        =   45
               Top             =   705
               Width           =   645
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TTL"
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
               Left            =   6480
               TabIndex        =   44
               Top             =   360
               Width           =   345
            End
            Begin VB.Label Label63 
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
               Left            =   8280
               TabIndex        =   43
               Top             =   1200
               Width           =   75
            End
            Begin VB.Label Label64 
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
               Left            =   7680
               TabIndex        =   42
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label65 
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
               Left            =   7680
               TabIndex        =   41
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hoby"
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
               Left            =   6480
               TabIndex        =   40
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label67 
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
               Left            =   7680
               TabIndex        =   39
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label Label68 
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
               Left            =   7680
               TabIndex        =   38
               Top             =   1080
               Width           =   75
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Msg SparePart"
            Height          =   1215
            Left            =   9840
            Picture         =   "FormFollowUpBooking.frx":102FE
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   6720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton CmdPros 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Follow Up"
            Height          =   1215
            Left            =   360
            Picture         =   "FormFollowUpBooking.frx":10A9F
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   6720
            Width           =   1215
         End
         Begin VB.CommandButton cmdProcess 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pr&ocess"
            Height          =   732
            Left            =   5280
            Picture         =   "FormFollowUpBooking.frx":1115F
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   480
            Width           =   732
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
            Left            =   10320
            Picture         =   "FormFollowUpBooking.frx":117CC
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cboType 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   615
            Width           =   2055
         End
         Begin MSComctlLib.ImageList ImageList 
            Left            =   7800
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   145
            ImageHeight     =   78
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormFollowUpBooking.frx":11F5C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormFollowUpBooking.frx":13362
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormFollowUpBooking.frx":14698
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormFollowUpBooking.frx":156F3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormFollowUpBooking.frx":16AA7
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate DtRemain 
            Height          =   315
            Left            =   2880
            TabIndex        =   5
            Top             =   270
            Width           =   2055
            _Version        =   65536
            _ExtentX        =   3625
            _ExtentY        =   556
            Calendar        =   "FormFollowUpBooking.frx":17CB9
            Caption         =   "FormFollowUpBooking.frx":17DE5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormFollowUpBooking.frx":17E51
            Keys            =   "FormFollowUpBooking.frx":17E6F
            Spin            =   "FormFollowUpBooking.frx":17ECD
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
         Begin TrueOleDBGrid70.TDBGrid GridFollowBook 
            Height          =   2535
            Left            =   120
            TabIndex        =   6
            Top             =   2160
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   4471
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Info"
            Columns(0).DataField=   "Deliv"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "N"
            Columns(1).DataField=   "FU"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "No Booking"
            Columns(2).DataField=   "NoFak"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Order Book"
            Columns(3).DataField=   "orderBook"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tanggal"
            Columns(4).DataField=   "Tanggal"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Jam"
            Columns(5).DataField=   "Jam"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "No Polisi"
            Columns(6).DataField=   "No_Polisi"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "No Rangka"
            Columns(7).DataField=   "No_Chassis"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Nama Pelanggan"
            Columns(8).DataField=   "Customer"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Alamat Pelanggan"
            Columns(9).DataField=   "Cust_Adres"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Tlp PSTN"
            Columns(10).DataField=   "Telpon"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Tlp GSM"
            Columns(11).DataField=   "GSM"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Tlp CDMA"
            Columns(12).DataField=   "CDMA"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Catatan"
            Columns(13).DataField=   "CAT"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Model"
            Columns(14).DataField=   "FMDL"
            Columns(14).DataWidth=   1500
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   15
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=15"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1402"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1296"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2196"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2090"
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(15)=   "Column(3).Width=3281"
            Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3175"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=1931"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1826"
            Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(23)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1667"
            Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(27)=   "Column(6).Width=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2540"
            Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(31)=   "Column(7).Width=4419"
            Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=4313"
            Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(35)=   "Column(8).Width=5292"
            Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=5186"
            Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(39)=   "Column(9).Width=7064"
            Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=6959"
            Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(43)=   "Column(10).Width=3254"
            Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=3149"
            Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(47)=   "Column(11).Width=3254"
            Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=3149"
            Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(51)=   "Column(12).Width=3254"
            Splits(0)._ColumnProps(52)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(12)._WidthInPix=3149"
            Splits(0)._ColumnProps(54)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(55)=   "Column(13).Width=3281"
            Splits(0)._ColumnProps(56)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(13)._WidthInPix=3175"
            Splits(0)._ColumnProps(58)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(59)=   "Column(14).Width=3281"
            Splits(0)._ColumnProps(60)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(61)=   "Column(14)._WidthInPix=3175"
            Splits(0)._ColumnProps(62)=   "Column(14).Order=15"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=86,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=94,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=82,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=79,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=80,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=81,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=78,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=62,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=28,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=54,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=51,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=52,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=53,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=50,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=47,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=48,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=49,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=90,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
            _StyleDefs(96)  =   "Named:id=33:Normal"
            _StyleDefs(97)  =   ":id=33,.parent=0"
            _StyleDefs(98)  =   "Named:id=34:Heading"
            _StyleDefs(99)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(100) =   ":id=34,.wraptext=-1"
            _StyleDefs(101) =   "Named:id=35:Footing"
            _StyleDefs(102) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(103) =   "Named:id=36:Selected"
            _StyleDefs(104) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(105) =   "Named:id=37:Caption"
            _StyleDefs(106) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(107) =   "Named:id=38:HighlightRow"
            _StyleDefs(108) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(109) =   "Named:id=39:EvenRow"
            _StyleDefs(110) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(111) =   "Named:id=40:OddRow"
            _StyleDefs(112) =   ":id=40,.parent=33"
            _StyleDefs(113) =   "Named:id=41:RecordSelector"
            _StyleDefs(114) =   ":id=41,.parent=34"
            _StyleDefs(115) =   "Named:id=42:FilterBar"
            _StyleDefs(116) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid GridBookPart 
            Height          =   1335
            Left            =   120
            TabIndex        =   13
            Top             =   5040
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   2355
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Kode Part"
            Columns(0).DataField=   "KodeItem"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nama Part"
            Columns(1).DataField=   "NamaBarang"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Qty Order"
            Columns(2).DataField=   "Qty_Pesan"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Qty Ready"
            Columns(3).DataField=   "Qty_SP"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Keterangan"
            Columns(4).DataField=   "Keterangan"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Jumlah Qty "
            Columns(5).DataField=   "Kelengkapan"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).AllowSizing=   -1  'True
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3043"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2937"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=5292"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5186"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2143"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2037"
            Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=2037"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1931"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=3281"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=3175"
            Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(23)=   "Column(5).Width=3228"
            Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=3122"
            Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=82,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=78,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.PictureBox FrPros 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1890
            Left            =   2040
            ScaleHeight     =   1890
            ScaleWidth      =   7455
            TabIndex        =   16
            Top             =   6480
            Visible         =   0   'False
            Width           =   7455
            Begin VB.OptionButton OptPros 
               Caption         =   "Call Again"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   1080
               TabIndex        =   127
               Top             =   1560
               Width           =   1575
            End
            Begin VB.CommandButton cmdBooking 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Booking"
               Height          =   732
               Left            =   2880
               Picture         =   "FormFollowUpBooking.frx":17EF5
               Style           =   1  'Graphical
               TabIndex        =   125
               Top             =   600
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.CommandButton cmdSave 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               Height          =   735
               Left            =   3600
               Picture         =   "FormFollowUpBooking.frx":1854A
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   600
               Width           =   735
            End
            Begin VB.PictureBox Picture3 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   4920
               ScaleHeight     =   255
               ScaleWidth      =   2295
               TabIndex        =   23
               Top             =   1320
               Width           =   2295
               Begin VB.Label lblTime 
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
                  Left            =   120
                  TabIndex        =   24
                  Top             =   0
                  Width           =   2055
               End
            End
            Begin VB.OptionButton OptPros 
               Caption         =   "Batal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1080
               TabIndex        =   22
               Top             =   1200
               Width           =   1575
            End
            Begin VB.OptionButton OptPros 
               Caption         =   "Reschdule"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1080
               TabIndex        =   21
               Top             =   840
               Width           =   1575
            End
            Begin VB.OptionButton OptPros 
               Caption         =   "OK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   1080
               TabIndex        =   19
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtAlasan 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   340
               Left            =   1080
               TabIndex        =   18
               Top             =   90
               Width           =   6135
            End
            Begin VB.Image ImgPros 
               Height          =   855
               Left            =   4920
               Top             =   480
               Width           =   2295
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Proses :"
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
               Left            =   140
               TabIndex        =   20
               Top             =   105
               Width           =   855
            End
            Begin VB.Line Line13 
               X1              =   120
               X2              =   960
               Y1              =   390
               Y2              =   390
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Proses :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.Frame FrameKendaraan 
            Caption         =   " DATA KENDARAAN "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1935
            Left            =   240
            TabIndex        =   81
            Top             =   13200
            Width           =   11175
            Begin VB.TextBox txtTahun 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8760
               MaxLength       =   4
               TabIndex        =   89
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtNoRangka 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   88
               Top             =   1080
               Width           =   3015
            End
            Begin VB.TextBox txtNoMesin 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   87
               Top             =   1440
               Width           =   3015
            End
            Begin VB.TextBox txtWarna 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   86
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txtType 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   85
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox lblWarna 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4080
               MaxLength       =   30
               TabIndex        =   84
               Top             =   720
               Width           =   2895
            End
            Begin VB.TextBox lblType 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4080
               MaxLength       =   30
               TabIndex        =   83
               Top             =   360
               Width           =   2895
            End
            Begin VB.TextBox txtNoPol 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8760
               MaxLength       =   10
               TabIndex        =   82
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label Label74 
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
               Left            =   8640
               TabIndex        =   102
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tahun"
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
               Left            =   7320
               TabIndex        =   101
               Top             =   360
               Width           =   585
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Rangka"
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
               Left            =   240
               TabIndex        =   100
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Warna"
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
               Left            =   240
               TabIndex        =   99
               Top             =   720
               Width           =   570
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
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
               Left            =   240
               TabIndex        =   98
               Top             =   345
               Width           =   465
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Mesin"
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
               Left            =   240
               TabIndex        =   97
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label Label52 
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
               Left            =   1560
               TabIndex        =   96
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label54 
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
               Left            =   1560
               TabIndex        =   95
               Top             =   720
               Width           =   75
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
               Left            =   1560
               TabIndex        =   94
               Top             =   1080
               Width           =   75
            End
            Begin VB.Label Label55 
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
               Left            =   1560
               TabIndex        =   93
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label lblMerk 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   2520
               TabIndex        =   92
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label90 
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
               Left            =   8640
               TabIndex        =   91
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nomor Polisi"
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
               Left            =   7320
               TabIndex        =   90
               Top             =   720
               Width           =   1245
            End
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   9240
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Line Line21 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   3960
            Y1              =   8880
            Y2              =   8880
         End
         Begin VB.Line Line20 
            BorderColor     =   &H000000FF&
            X1              =   1440
            X2              =   3960
            Y1              =   8760
            Y2              =   8760
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000000FF&
            X1              =   2280
            X2              =   3960
            Y1              =   8640
            Y2              =   8640
         End
         Begin VB.Line Line18 
            BorderColor     =   &H000000FF&
            X1              =   2400
            X2              =   11400
            Y1              =   16315
            Y2              =   16315
         End
         Begin VB.Line Line17 
            BorderColor     =   &H000000FF&
            X1              =   1320
            X2              =   10320
            Y1              =   16265
            Y2              =   16265
         End
         Begin VB.Line Line16 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   9120
            Y1              =   16215
            Y2              =   16215
         End
         Begin VB.Line Line14 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   11400
            Y1              =   13100
            Y2              =   13100
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   1455
            Left            =   9720
            Shape           =   4  'Rounded Rectangle
            Top             =   6600
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Shape ShapeFu 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   1455
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   6600
            Width           =   1455
         End
         Begin VB.Line Line12 
            BorderColor     =   &H000000FF&
            X1              =   2160
            X2              =   3960
            Y1              =   4770
            Y2              =   4770
         End
         Begin VB.Line Line11 
            BorderColor     =   &H000000FF&
            X1              =   1440
            X2              =   3960
            Y1              =   4860
            Y2              =   4860
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            X1              =   720
            X2              =   3960
            Y1              =   4965
            Y2              =   4965
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   5160
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   2400
            X2              =   4080
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   1560
            X2              =   4080
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            X1              =   360
            X2              =   4080
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LIST FOLLOW UP BOOKING"
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
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   11295
         End
         Begin VB.Image ImgView 
            Height          =   1095
            Left            =   6240
            Top             =   240
            Width           =   2295
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   10200
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   5040
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   5040
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblKet 
            Caption         =   "Keterangan"
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
            Left            =   360
            TabIndex        =   11
            Top             =   1035
            Width           =   4695
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   5040
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label2 
            Caption         =   "Type Follow Up"
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
            Left            =   360
            TabIndex        =   10
            Top             =   675
            Width           =   2415
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   5040
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Tanggal Reminder"
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
            Left            =   360
            TabIndex        =   9
            Top             =   315
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Data  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   9240
            TabIndex        =   8
            Top             =   1860
            Width           =   1455
         End
         Begin VB.Label lblJml 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   10440
            TabIndex        =   7
            Top             =   1860
            Width           =   855
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0FF&
            Height          =   495
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   1680
            Width           =   11235
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "    DATA FOLLOW UP BOOKING"
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
            Left            =   120
            TabIndex        =   126
            Top             =   8640
            Width           =   11295
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0FF&
            Height          =   495
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   8520
            Width           =   11235
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LIST PART ORDER - No. ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1080
            TabIndex        =   15
            Top             =   4755
            Width           =   10335
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0FF&
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   4680
            Width           =   11220
         End
      End
   End
   Begin VB.Line Line15 
      X1              =   480
      X2              =   11520
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "FormFollowUpBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private OnLoad As Boolean
Private rsFol As ADODB.Recordset
Private objFol As New ClsFollowUpBooking
Private rsPart As ADODB.Recordset
Private rsFUB As ADODB.Recordset
Private isIndex As Integer

Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String


Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub cboType_Click()
   Select Case cboType.Text
      Case " H - 1"
         Me.lblKet.Caption = "Keterangan     : Booking 1 Hari Ke Depan !"
         Me.ImgView.Picture = Me.ImageList.ListImages(1).Picture
         chAll.Visible = True
      Case " J - 1"
         Me.lblKet.Caption = "Keterangan     : Booking 1 Jam Ke Depan !"
         Me.ImgView.Picture = Me.ImageList.ListImages(2).Picture
         chAll.Visible = False
      Case Else
         Me.lblKet.Caption = "Keterangan"
         chAll.Visible = False
   End Select
End Sub

Private Sub cmdBooking_Click()
If GridFollowBook.Columns(0).value = "BTL" Then
    If Trim(GridFollowBook.Columns(3).value) < Trim(Format(isnow, "DD-MM-YY")) Then
        Exit Sub
    End If
End If
Dim noPol As String
    With frmBooking
        noPol = Trim(GridFollowBook.Columns(6))
        Dim nDpn1 As String
        Dim nTgh1 As String
        Dim nBlk1 As String
        Call pecahNoPol(noPol, nDpn1, nTgh1, nBlk1)
        .Text(0) = Trim(nDpn1)
        .Text(56) = Trim(nTgh1)
        ''2016-02-24 || user tidak bisa melakukan reschedule booking
        .passingX = False
        
        ''.passingX = True
        .Text(57) = Trim(nBlk1)
        .show
    End With
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdTop_Click()
   Me.VScrollFUB.value = 1
End Sub

Private Sub Form_Load()
    '' 20200512DEV
    If Log.disableExportMra = 1 Then
        If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
            'cmdCSV.Enabled = True
            CmdXls.Enabled = True
        Else
            'cmdCSV.Enabled = False
            CmdXls.Enabled = False
        End If
    End If
   If OnLoad = False Then Exit Sub
   Call ListAllCombo
   Call LoadForm
End Sub

Private Sub ListAllCombo()
   With Me
      .cboType.Clear
      .cboType.AddItem " H - 1"
      .cboType.AddItem " J - 1"
   End With
End Sub

Private Sub LoadForm()
   CmdXls.Enabled = False
   Me.DtRemain.value = Date
End Sub


Private Sub cmdProcess_Click()
1        If Me.DtRemain.value = Null Or Me.cboType.Text = "" Then Exit Sub
2        CmdXls.Enabled = False
3        Select Case cboType.Text
            Dim Er As New ErrInfo
            Case " H - 1"
4              Set rsFol = objFol.FetchDataFollowUpHari(DtRemain.value, IIf(chAll.value = Checked, True, False))
5              If rsFol.RecordCount = 0 Then
6                 Er.HeadCap = "Follow'Up Booking !"
7                 Er.InfoCap = "Data Booking untuk tanggal " & DateAdd("d", 1, DtRemain.value) & ", tidak ada !"
8                 Beep
9                 Er.show (vbModal)
10                Set Me.GridFollowBook.DataSource = Nothing
11                lblJml.Caption = "0"
12             Else
                  '' 20200512DEV
13                If Log.disableExportMra = 1 Then
14                    If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
                          'cmdCSV.Enabled = True
15                        CmdXls.Enabled = True
16                    Else
                          'cmdCSV.Enabled = False
17                        CmdXls.Enabled = False
18                    End If
19                Else
20                    CmdXls.Enabled = True
21                End If
                  ''CmdXls.Enabled = True
22                Set Me.GridFollowBook.DataSource = rsFol
23                lblJml.Caption = rsFol.RecordCount
24             End If
25          Case " J - 1"
26             Set rsFol = objFol.FetchDataFollowUpJam(DtRemain.value, Val(Replace(Trim(Format(isnow, "HH:NN")), ":", "")), Val(Replace(Trim(Format(DateAdd("h", 1, Now), "HH:NN")), ":", "")))
27             If rsFol.RecordCount = 0 Then
28                Er.HeadCap = "Follow'Up Booking !"
29                Er.InfoCap = "Data Booking untuk tanggal " & DtRemain.value & " Pada Jam : " & Trim(Format(isnow, "HH:NN")) & " - " & Trim(Format(DateAdd("h", 1, Now), "HH:NN")) & " , tidak ada !"
30                Beep
31                Er.show (vbModal)
32                Set Me.GridFollowBook.DataSource = Nothing
33                lblJml.Caption = "0"
34             Else
                  '' 20200512DEV
35                If Log.disableExportMra = 1 Then
36                    If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
                          'cmdCSV.Enabled = True
37                        CmdXls.Enabled = True
38                    Else
                          'cmdCSV.Enabled = False
39                        CmdXls.Enabled = False
40                    End If
41                Else
42                    CmdXls.Enabled = True
43                End If
                  ''CmdXls.Enabled = True
44                Set Me.GridFollowBook.DataSource = rsFol
45                lblJml.Caption = rsFol.RecordCount
46             End If
47       End Select
End Sub

Private Sub GridFollowBook_FilterChange()
On Error GoTo err
   Set cols = Me.GridFollowBook.Columns
   Dim C As Integer
   
   C = GridFollowBook.col
   GridFollowBook.HoldFields
   rsFol.filter = getFilter()
   If rsFol.EOF = True Then
      For Each col In GridFollowBook.Columns
         col.FilterText = ""
      Next col
      rsFol.filter = ""
   End If
   GridFollowBook.col = C
   GridFollowBook.EditActive = True
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
         tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub GridFollowBook_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 16 Then rsFol.Sort = GridFollowBook.Columns(ColIndex).DataField & " " & Urut
4        GridFollowBook.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormFollowUpBooking")

End Sub

Private Sub GridFollowBook_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   FrPros.Visible = False
   CmdPros.Visible = True
   ShapeFu.Visible = True
   lblPart.Caption = "LIST PART ORDER - No." & GridFollowBook.Columns(2).value
   Call ViewDataSPKItem
   Set GridBookPart.DataSource = Nothing
   Set rsPart = objFol.FetchDataPart(GridFollowBook.Columns(2).value)
   If GridFollowBook.Columns(0).value = "OK" Then Call CmdPros_Click
   If rsPart.RecordCount = 0 Then Exit Sub
   Set GridBookPart.DataSource = rsPart
End Sub

Private Sub ViewDataSPKItem()
    Dim rsS As ADODB.Recordset
    Set rsS = objFol.FetchItemSPK(GridFollowBook.Columns(6).value, GridFollowBook.Columns(5).value)
    With Me
        If rsS.RecordCount <> 0 Then
            .txtNamaSTNK.Text = Nul(rsS.Fields(22).value)
            .txtAlamatSTNK.Text = Nul(rsS.Fields(23).value)
            .txtKodePosSTNK.Text = Nul(rsS.Fields(25).value)
            .txtKotaSTNK.Text = Nul(rsS.Fields(24).value)
            .txtTempatLahir.Text = Nul(rsS.Fields(26).value)
            .DTLahir.Text = TglNul(rsS.Fields(27).value)
            .txtAgama.Text = Nul(rsS.Fields(28).value)
            .txtHoby.Text = Nul(rsS.Fields(30).value)
            .txtPekerjaan = Nul(rsS.Fields(29).value)
            .txtNamaCust2.Text = Nul(rsS.Fields(6).value)
            .txtAlamatCust2.Text = Nul(rsS.Fields(7).value)
            .txtKodePosCust2.Text = Nul(rsS.Fields(10).value)
            .txtKotaCust2.Text = Nul(rsS.Fields(8).value)
            .txtPropinsiCust2.Text = Nul(rsS.Fields(9).value)
            .txtTelpCDMA.Text = Nul(rsS.Fields(16).value)
            .txtTelpGSM.Text = Nul(rsS.Fields(15).value)
            .txtTelpRumah.Text = Nul(rsS.Fields(14).value)
            .txtType.Text = Nul(rsS.Fields(17).value)
            .lblType.Text = Nul(rsS.Fields(18).value)
            .txtWarna.Text = Nul(rsS.Fields(19).value)
            .lblWarna.Text = Nul(rsS.Fields(20).value)
            .txtNoRangka.Text = Nul(rsS.Fields(13).value)
            .txtNoMesin.Text = Nul(rsS.Fields(12).value)
            .txtTahun.Text = Nul(rsS.Fields(33).value)
            .DtDEC.Text = TglNul(rsS.Fields(1).value)
            .DtDO.Text = TglNul(rsS.Fields(1).value)
            .DTCR1.Text = TglNul(rsS.Fields(2).value)
            .DTCR2.Text = TglNul(rsS.Fields(4).value)
            .txtKMCR1.Text = Nul(rsS.Fields(3).value)
            .txtKMCR2.Text = Nul(rsS.Fields(5).value)
        Else
            .txtNamaSTNK.Text = ""
            .txtAlamatSTNK.Text = ""
            .txtKodePosSTNK.Text = ""
            .txtKotaSTNK.Text = ""
            .txtTempatLahir.Text = ""
            .DTLahir.Text = "__/__/____"
            .txtAgama.Text = ""
            .txtHoby.Text = ""
            .txtPekerjaan = ""
            .txtNamaCust2.Text = ""
            .txtAlamatCust2.Text = ""
            .txtKodePosCust2.Text = ""
            .txtKotaCust2.Text = ""
            .txtPropinsiCust2.Text = ""
            .txtTelpCDMA.Text = ""
            .txtTelpGSM.Text = ""
            .txtTelpRumah.Text = ""
            .txtType.Text = ""
            .lblType.Text = ""
            .txtWarna.Text = ""
            .lblWarna.Text = ""
            .txtNoRangka.Text = ""
            .txtNoMesin.Text = ""
            .txtTahun.Text = ""
            .DtDEC.Text = "__/__/____"
            .DtDO.Text = "__/__/____"
            .DTCR1.Text = "__/__/____"
            .DTCR2.Text = "__/__/____"
            .txtKMCR1.Text = ""
            .txtKMCR2.Text = ""
        End If
        .txtNamaCust2.Text = GridFollowBook.Columns(7).value
        .txtAlamatCust2.Text = GridFollowBook.Columns(8).value
        .txtNoPol.Text = GridFollowBook.Columns(5).value
        .txtTelpCDMA.Text = GridFollowBook.Columns(11).value
        .txtTelpGSM.Text = GridFollowBook.Columns(10).value
        .txtTelpRumah.Text = GridFollowBook.Columns(9).value
    End With
End Sub

Private Sub CmdPros_Click()
On Error GoTo ErrNull
   If rsFol.RecordCount = 0 Then Exit Sub
   FrPros.Visible = True
   CmdPros.Visible = False
   ShapeFu.Visible = False
   Call View_M_Process(GridFollowBook.Columns(2).value)
ErrNull:
Exit Sub
End Sub

Private Sub View_M_Process(NoFak As String)
   Set rsFUB = objFol.Get_FU_Booking(NoFak)
   With Me
      If rsFUB.RecordCount = 0 Then ' Belum di proses
         .OptPros(0).value = False
         .OptPros(1).value = False
         .OptPros(2).value = False
         .txtAlasan.Text = ""
         .ImgPros.Picture = Nothing
      Else                          ' Sudah di Proses
         .OptPros(CInt(rsFUB.Fields(1).value)).value = True
         .txtAlasan.Text = rsFUB.Fields(3).value
         .lblTime.Caption = Trim(rsFUB.Fields(2).value)
         .ImgPros.Picture = Nothing
      End If
   End With
   
   If Me.OptPros(0).value = True Then
         Me.ImgPros.Picture = Me.ImageList.ListImages(3).Picture
   ElseIf Me.OptPros(1).value = True Then
         Me.ImgPros.Picture = Me.ImageList.ListImages(4).Picture
   ElseIf Me.OptPros(2).value = True Then
         Me.ImgPros.Picture = Me.ImageList.ListImages(5).Picture
   End If
End Sub


Private Sub OptPros_Click(Index As Integer)
   Select Case Index
      Case 0
         Me.cmdBooking.Visible = False
         Me.ImgPros.Picture = Me.ImageList.ListImages(3).Picture
      Case 1
         Me.cmdBooking.Visible = True
         Me.ImgPros.Picture = Me.ImageList.ListImages(4).Picture
      Case 2
         Me.cmdBooking.Visible = False
         Me.ImgPros.Picture = Me.ImageList.ListImages(5).Picture
      Case 3
         Me.cmdBooking.Visible = False
         Me.ImgPros.Picture = Nothing 'Me.ImageList.ListImages(5).Picture
   End Select
   isIndex = Index
End Sub

Private Sub cmdSave_Click()
   Dim Er As New ErrInfo
   If OptPros(0).value = False And OptPros(1).value = False And OptPros(2).value = False And OptPros(3).value = False Then
        Er.HeadCap = "Follow Up Proses !"
        Er.InfoCap = "Proses Belum Dipilih !"
        Beep
        Er.show (vbModal)
      Exit Sub
   End If
   Call objFol.AddNew_FUBOOK(GridFollowBook.Columns(2).value, isIndex, txtAlasan.Text)
   Call objFol.Update_DelivBOOK(isIndex, GridFollowBook.Columns(2).value)
        Er.HeadCap = "Follow'Up Proses !"
        Er.InfoCap = "Proses telah tersimpan !"
        Beep
        Er.show (vbModal)
   lblTime.Caption = Now
End Sub


Private Sub VScrollFUB_Change()
    Me.SSFUB.Top = -(VScrollFUB.value)
End Sub
Private Sub VScrollFUB_Scroll()
    Me.SSFUB.Top = -(VScrollFUB.value)
End Sub

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsFol.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
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
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
16        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
17        sSQL = sSQL & " values ('" & gSA & "','FormFollowUpBooking',getdate(),'" & Format(DtRemain.value, "MM/DD/YYYY") & "','" & Format(DtRemain.value, "MM/DD/YYYY") & "','" & cboType.Text & "','" & rsFol.RecordCount & "','" & onVal & "','" & ComptName & "') "
18        Call objO.InsertData(sSQL)
          
19         GridFollowBook.Columns(9).Visible = False
20         GridFollowBook.ExportToFile onVal, True
21         GridFollowBook.Columns(9).Visible = True
22                Er.HeadCap = "Export to excel !"
23                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
24                Beep
25                Er.show (vbModal)
                  'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
26         Exit Sub
Err_Export:
27          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
28          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub
