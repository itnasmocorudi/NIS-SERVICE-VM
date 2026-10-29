VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmAturMech 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jam Kerja Mekanik"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5385
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OK"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cboShift 
         Height          =   315
         ItemData        =   "frmAturMech.frx":0000
         Left            =   1320
         List            =   "frmAturMech.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   735
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
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
         Mask            =   ">&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   4200
         TabIndex        =   2
         Top             =   622
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Selesai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3195
         TabIndex        =   10
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Mulai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3330
         TabIndex        =   9
         Top             =   278
         Width           =   690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   7
         Top             =   285
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   6
         Top             =   660
         Width           =   330
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   480
      Left            =   1560
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   847
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin TrueOleDBGrid70.TDBGrid ListMech 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6588
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "mech"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "name"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   4
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "masuk"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmAturMech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : -
'=======================================

Dim rsTemp As New ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim temp As String

Private Sub hapusForm()
    
    text(0) = IIf(Weekday(gTglServer) = 1, 6, Weekday(gTglServer) - 2) & WeekdayName(IIf(Weekday(gTglServer) = 1, 8, Weekday(gTglServer)) - 1)
    '(Weekday(gTglServer) - 1) & WeekdayName(Weekday(gTglServer) - 1)
    text(1) = "00:00"
    text(1).Mask = "99:99"
    text(2) = "00:00"
    text(2).Mask = "99:99"
    
    ' ISI COMBO
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT * FROM MECHSHIFT ORDER BY SHIFT")
    If Not rsTemp.EOF Then
        cboShift.clear
        Do While Not rsTemp.EOF
            cboShift.AddItem rsTemp("SHIFT")
            rsTemp.MoveNext
        Loop
    End If
    rsTemp.Close
End Sub

Private Function createSublet() As Boolean
    On Error GoTo Vb_Error
    
    If rsSublet.State Then rsSublet.Close
    With rsSublet
        .Fields.Append "KODEMECH", adVarChar, 25
        .Fields.Append "NAMAMECH", adVarChar, 30
        .Fields.Append "MASUK", adBoolean
        .Fields.Append "HARI", adVarChar, 10
        .Fields.Append "TGL", adVarChar, 10
        .Fields.Append "SHIFT", adVarChar, 1
        .Fields.Append "JAMMULAI", adVarChar, 5
        .Fields.Append "JAMSELESAI", adVarChar, 5
        .Fields.Append "MENIT", adInteger
        .Open
    End With
    If rsSublet.State Then createSublet = True
    Exit Function
    
Vb_Error:
    MsgBox ErrMessage(Erl, Err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub refreshGrid()
    
    If rsTemp.State Then rsTemp.Close
    ' LIHAT LAGI
    Set rsTemp = oConn.Execute("SELECT distinct A.mech, name, isnull(masuk,0) as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                             & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                             & "FROM MEKANIK A LEFT JOIN T_MEKANIK B ON A.MECH = B.MECH AND TANGGAL = GETDATE() AND HARI = '" & IIf(Weekday(gTglServer) = 1, 6, Weekday(gTglServer) - 2) & WeekdayName(IIf(Weekday(gTglServer) = 1, 8, Weekday(gTglServer)) - 1) & "' ORDER BY a.MECH")
    '(Weekday(gTglServer) - 1) & WeekdayName(Weekday(gTglServer) - 1)
    If Not rsTemp.EOF Then
        If createSublet Then
            Do While Not rsTemp.EOF
                With rsSublet
                    .AddNew
                    !KODEMECH = rsTemp("MECH")
                    !NAMAMECH = rsTemp("NAME")
                    !MASUK = rsTemp("MASUK")
                    !hari = rsTemp("HARI")
                    !Tgl = rsTemp("TANGGAL")
                    !Shift = rsTemp("SHIFT")
                    !jamMulai = rsTemp("JAMMULAI")
                    !jamSelesai = rsTemp("JAMSELESAI")
                    !menit = rsTemp("MENITKERJA")
                    .Update
                    .MoveNext
                End With
                rsTemp.MoveNext
            Loop
        End If
    End If
    ListMech.DataSource = rsSublet
    Call headColumns
End Sub

Private Sub headColumns()
    
    With ListMech
        .Columns(0).Caption = "KODE"
        .Columns(0).DataField = "KODEMECH"
        .Columns(0).Width = 800
        .Columns(1).Caption = "NAMA"
        .Columns(1).DataField = "NAMAMECH"
        .Columns(1).Width = 2000
        .Columns(2).Caption = "ADA"
        .Columns(2).DataField = "MASUK"
        .Columns(2).Width = 800
        .Columns(2).Alignment = dbgCenter
    End With
End Sub

Private Sub cboShift_Change()
    cboShift_Click
End Sub

Private Sub cboShift_Click()
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT * FROM MECHSHIFT WHERE SHIFT = '" & cboShift & "'")
    If Not rsTemp.EOF Then
        text(1) = rsTemp("JAMMULAI")
        text(2) = rsTemp("JAMSELESAI")
    End If
    rsTemp.Close
End Sub

Private Sub cboShift_KeyUp(KeyCode As Integer, Shift As Integer)
'    If rsTemp.State Then rsTemp.Close
'    Set rsTemp = oConn.Execute("SELECT * FROM MECHSHIFT WHERE SHIFT = '" & cboShift & "'")
'    If Not rsTemp.EOF Then
'        Text(1) = rsTemp("JAMMULAI")
'        Text(2) = rsTemp("JAMSELESAI")
'    End If
'    rsTemp.Close
    cboShift_Click
End Sub

Private Sub cmdOK_Click()
    
    If cboShift <> "" Then
        If Not rsSublet.EOF Then
            rsSublet.MoveFirst
            rsSublet.Find "KODEMECH = '" & temp & "'", , adSearchForward, 1
            If Not rsSublet.EOF Then
                rsSublet!hari = text(0)
                rsSublet!Tgl = Format(gTglServer, "DD/MM/YYYY")
                rsSublet!Shift = CDbl(cboShift)
                rsSublet!jamMulai = text(1)
                rsSublet!jamSelesai = text(2)
                rsSublet!menit = CDbl(hitMenit(rsSublet("JAMMULAI"), rsSublet("JAMSELESAI")))
                rsSublet.Update
            End If
        End If
        Frame2.Visible = False
    Else
        MsgBox "Pilih Shift !!!"
        Exit Sub
    End If
End Sub

Private Sub cmdSimpan_Click()
    ' SIMPAN
    If Not rsSublet.State = adStateClosed Then
        rsSublet.MoveFirst
        If Not rsSublet.EOF Then
            Do While Not rsSublet.EOF
                If rsTemp.State Then rsTemp.Close
                Set rsTemp = oConn.Execute("SELECT * FROM T_MEKANIK WHERE MECH = '" & rsSublet("KODEMECH") & "' AND HARI = '" & rsSublet("HARI") & "' AND TANGGAL = '" & rsSublet("TGL") & "'")
                If Not rsTemp.EOF Then
                    If rsSublet("HARI") = "" Then
                        If rsTemp("TERPAKAI") > 0 Then
                            oConn.Execute ("UPDATE T_MEKANIK SET MENITKERJA = 0, MASUK = 0 WHERE MECH = '" & rsSublet("KODEMECH") & "' AND TANGGAL = GETDATE()")
                        Else
                            oConn.Execute ("DELETE T_MEKANIK WHERE MECH = '" & rsSublet("KODEMECH") & "' AND TANGGAL = GETDATE()")
                        End If
                    Else
                        oConn.Execute ("UPDATE T_MEKANIK SET TANGGAL = '" & rsSublet("TGL") & "', SHIFT = '" & IIf(rsSublet("SHIFT") = "", 0, rsSublet("SHIFT")) & "', jammulai =  '" & rsSublet("JAMMULAI") & "', jamselesai = '" & rsSublet("JAMSELESAI") & "', masuk = '" & IIf(rsSublet("MASUK"), 1, 0) & "', menitkerja = '" & rsSublet("MENIT") & "' WHERE MECH = '" & rsSublet("KODEMECH") & "' AND HARI = '" & rsSublet("HARI") & "'")
                    End If
                Else
                    oConn.Execute ("INSERT INTO T_MEKANIK VALUES ('" & rsSublet("KODEMECH") & "', '" & rsSublet("HARI") & "', '" & rsSublet("TGL") & "', '" & IIf(rsSublet("SHIFT") = "", 0, rsSublet("SHIFT")) & "', '" & rsSublet("JAMMULAI") & "', '" & rsSublet("JAMSELESAI") & "', '" & rsSublet("MENIT") & "', 0, '" & IIf(rsSublet("MASUK"), 1, 0) & "' )")
                End If
                
                ' BELUM KELAR
                Dim rsHlp As New ADODB.Recordset
                If rsHlp.State Then rsHlp.Close
                Set rsHlp = oConn.Execute("SELECT * FROM JOBDTA WHERE [DATE] = GETDATE() AND MEKANIK1 + ' ' + MEKANIK2 + ' ' + MEKANIK3 + ' ' + MEKANIK4 + ' ' + MEKANIK5 + ' ' + MEKANIK6 + ' ' + MEKANIK7 LIKE '" & rsSublet("KODEMECH") & "' AND ([TIME] < '" & IIf(Trim(rsSublet("JAMMULAI")) <> "", Trim(rsSublet("JAMMULAI")), "00:00") & "' OR [TIME] > '" & IIf(Trim(rsSublet("JAMSELESAI")) <> "", Trim(rsSublet("JAMSELESAI")), "23:59") & "')")
                If Not rsHlp.EOF Then
                    Do While Not rsHlp.EOF
                        MsgBox "WO dengan No. " & rsHlp("NO_WO") & " dikerjakan oleh mekanik " & rsSublet("KODEMECH") & vbCr & "Ganti dengan mekanik lain. "
                        rsHlp.MoveNext
                    Loop
                End If
                rsHlp.Close
                
                ' MEKANIK TIDAK MASUK, PADAHAL SUDAH DIBLOK UTK OH JOB.
                If Not rsSublet("MASUK") Then
                    If rsHlp.State Then rsHlp.Close
                    Set rsHlp = oConn.Execute("SELECT * FROM STALL WHERE FLAG = 'P' OR FLAG = 'S' AND MECH = '" & rsSublet("KODEMECH") & "'")
                    If Not rsHlp.EOF Then
                        Do While Not rsHlp.EOF
                            MsgBox "Unit dengan No Polisi " & Left(rsHlp("KETERANGAN"), InStr(rsHlp("KETERANGAN"), "|") - 1) & " di " & rsHlp("NOSTALL") & " dikerjakan oleh mekanik " & rsSublet("KODEMECH") & vbCr & "Ganti dengan mekanik lain."
                            rsHlp.MoveNext
                        Loop
                    End If
                    rsHlp.Close
                End If
                rsSublet.MoveNext
            Loop
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call hapusForm
    hariNow = WeekdayName(IIf(Weekday(gTglServer) = 1, 7, Weekday(gTglServer) - 1)) & ", " & Format(gTglServer, "DD - MMMM - YYYY")
    Call refreshGrid
End Sub

Private Sub ListMech_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If ListMech.Columns(2) <> "" Then
        If ListMech.Columns(2) <> 0 Then
            temp = ListMech.Columns(0)
            Frame2.Visible = True
            rsSublet.MoveFirst
            rsSublet.Find "KODEMECH = '" & temp & "'", , adSearchForward, 1
            If Not rsSublet.EOF Then
                If rsSublet!Shift <> "" Then
                    cboShift = IIf(rsSublet!Shift <> "0", rsSublet!Shift, 1)
                Else
                    cboShift = IIf(rsSublet!Shift <> "", rsSublet!Shift, 1)
                End If
                'Text(1) = IIf(Trim(rsSublet!jamMulai) = "", 0, rsSublet!jamMulai)
                'Text(2) = IIf(Trim(rsSublet!jamSelesai) = "", 0, rsSublet!jamSelesai)
            End If
        Else
            Frame2.Visible = False
        End If
    End If
End Sub

Private Sub text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
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

Private Sub ListMech_AfterColUpdate(ByVal ColIndex As Integer)
Dim rsHlp As New ADODB.Recordset
    
    If ColIndex = 2 Then
        temp = ListMech.Columns(0)
        If ListMech.Columns(2) = 0 Then
            
            If rsHlp.State Then rsHlp.Close
'            Set rsHlp = oConn.Execute("SELECT * FROM JOBDTA WHERE [DATE] = '" & Format(gTglServer, "MM/DD/YYYY") & "' AND MEKANIK1 + ' ' +MEKANIK2+' '+MEKANIK3+' '+MEKANIK4+' '+MEKANIK5+' '+MEKANIK6+' '+MEKANIK7 LIKE '%" & temp & "%'")
            Set rsHlp = oConn.Execute("SELECT DISTINCT NOORDER FROM PROGRESS WHERE TANGGAL = GETDATE() AND MECH LIKE '%" & temp & "%'")
            If Not rsHlp.EOF Then
                Do While Not rsHlp.EOF
                    MsgBox "Mekanik ini mengerjakan unit Hari ini. Tidak bisa dihapus" & vbCr & "Ubah dulu WO : " & rsHlp("NOORDER")
                    rsHlp.MoveNext
                Loop
                ListMech.Columns(2) = 1
                Exit Sub
            End If
            rsHlp.Close
            
            Frame2.Visible = False
            rsSublet.MoveFirst
            rsSublet.Find "KODEMECH = '" & temp & "'", , adSearchForward, 1
            If Not rsSublet.EOF Then
                rsSublet!hari = ""
                rsSublet!Tgl = ""
                rsSublet!Shift = ""
                rsSublet!jamMulai = "00:00"
                rsSublet!jamSelesai = "00:00"
                rsSublet!menit = 0
                rsSublet.Update
            End If
            temp = ""
            Call hapusForm
        Else
            Frame2.Visible = True
            cboShift.ListIndex = 0
'            rsSublet.MoveFirst
'            rsSublet.Find "KODEMECH = '" & temp & "'", , adSearchForward, 1
'            If Not rsSublet.EOF Then
'                cboShift = IIf(rsSublet!Shift = "", 1, rsSublet!Shift)
'                Text(1) = IIf(rsSublet!jamMulai = "", "0", rsSublet!jamMulai)
'                Text(2) = IIf(rsSublet!jamSelesai = "", "0", rsSublet!jamSelesai)
'            End If
        End If
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)

    Select Case Index
        Case 1
            If Replace(text(Index), ":", "") = "" Then
                text(Index) = "00:00"
            ElseIf CDbl(Replace(text(Index), ":", "")) >= 2400 Then
                text(Index) = "00:00"
'            Else
'                rsSublet.MoveFirst
'                rsSublet.Find "KODEMECH = '" & temp & "'", , adSearchForward, 1
'                If Not rsSublet.EOF Then
'                    rsSublet!HARI = Text(0)
'                    rsSublet!Tgl = Format(gTglServer, "DD/MM/YYYY")
'                    rsSublet!Shift = CDbl(cboShift)
'                    rsSublet!jamMulai = Text(Index)
'                    rsSublet.Update
'                End If
            End If
            
        Case 2
            If Replace(text(Index), ":", "") = "" Then
                text(Index) = "00:00"
            ElseIf CDbl(Replace(text(Index), ":", "")) > 2400 Then
                text(Index) = "00:00"
            ElseIf CDbl(Replace(text(Index), ":", "")) < CDbl(Replace(text(1), ":", "")) Then
                text(Index) = Format(CDbl(Left(text(1), 2)) + 1, "0#") & ":" & Format(Right(text(1), 2), "0#")
            Else
                rsSublet.MoveFirst
                rsSublet.Find "KODEMECH = '" & temp & "'", , adSearchForward, 1
                If Not rsSublet.EOF Then
                    rsSublet!hari = text(0)
                    rsSublet!Tgl = Format(gTglServer, "DD/MM/YYYY")
                    rsSublet!Shift = CDbl(cboShift)
                    rsSublet!jamMulai = text(1)
                    rsSublet!jamSelesai = text(Index)
                    rsSublet!menit = CDbl(hitMenit(rsSublet("JAMMULAI"), rsSublet("JAMSELESAI")))
                    rsSublet.Update
                End If
            End If
    End Select
End Sub
