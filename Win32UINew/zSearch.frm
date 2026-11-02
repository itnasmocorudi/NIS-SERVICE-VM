VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form zSearch 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10095
   Icon            =   "zSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox xxx 
      Height          =   375
      Left            =   9000
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Pilih 
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
      Left            =   0
      TabIndex        =   8
      Top             =   6885
      Width           =   4485
   End
   Begin VB.ComboBox cboS 
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
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox cari 
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
      Height          =   320
      IMEMode         =   3  'DISABLE
      Left            =   3200
      MaxLength       =   30
      TabIndex        =   1
      Top             =   600
      Width           =   6885
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   5580
      Left            =   0
      TabIndex        =   2
      Top             =   1000
      Width           =   10060
      _ExtentX        =   17754
      _ExtentY        =   9843
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
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
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin ucXPButton.XPButton cmdOpen 
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   6720
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&d"
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
      MICON           =   "zSearch.frx":29C12
      PICN            =   "zSearch.frx":29C2E
      PICH            =   "zSearch.frx":2A0EF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdCancel 
      Height          =   495
      Left            =   8880
      TabIndex        =   4
      Top             =   6720
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&c"
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
      MICON           =   "zSearch.frx":2A645
      PICN            =   "zSearch.frx":2A661
      PICH            =   "zSearch.frx":2ABA5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblJml 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9240
      TabIndex        =   12
      Top             =   60
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   10080
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected :"
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
      Left            =   0
      TabIndex        =   9
      Top             =   6630
      Width           =   960
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Text :"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   1290
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search By :"
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
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   10080
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   0
      X2              =   10080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: LIST  ::."
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
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   45
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "zSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsZ         As ADODB.Recordset
Private nCol        As Integer

Private MySearch    As String
Private MyRs        As ADODB.Recordset
Private isPassing   As Boolean

Private Urut        As String
Private n           As Integer
Private param       As String
Dim i As Integer
Private isModel     As String

Public Property Let SetSearch(value As String)
    MySearch = value
End Property
Public Property Let SetModel(value As String)
    isModel = value
End Property
Public Property Get rs() As ADODB.Recordset
    Set rs = MyRs
End Property
Public Property Get IsPass() As Boolean
    IsPass = isPassing
End Property

Private Sub cari_Change()
        If cboS.text = "" Then Exit Sub
        cari.text = Replace(cari.text, "'", "`")
        param = " WHERE " & cboS.text & " Like '%" & cari.text & "%'"
        If cari.text = "" Or cari.text = "'" Then
            param = " WHERE " & cboS.text & " Like 'Antztein'"
        End If
        Select Case ShowForm
            Case "Rujukan"
                Call ViewFieldRujukan_Search(param & " And UserDel = ''")
            Case "Model"
                Call ViewFieldModel_Search(param)
            Case "Warna"
                Call ViewFieldWarna_Search(param)
            Case "Custdat"
                Call ViewFieldCustdat_Search(param)
        End Select
        On Error Resume Next
        lblJml.Caption = rsZ.RecordCount
End Sub

Private Sub cboS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo ErrNull
    isPassing = False
    Set MyRs = rsZ
    If MyRs.RecordCount <> 0 Then
        isPassing = True
    End If
    Unload Me
    Exit Sub
ErrNull:
    isPassing = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    isPassing = False
    Unload Me
End Sub

Private Sub Command1_Click()
    MsgBox Grid.Columns(CInt(xxx.text)).Width
End Sub


'=========================================================================================================
'SORT
'=========================================================================================================
Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= nCol Then rsZ.Sort = "[" & Grid.Columns(ColIndex).DataField & "]" & " " & Urut
4        Grid.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsZ.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter ")
End Sub
'=========================================================================================================
' 1111111111111111111111111111111111111111111111111111111111111111111111
Private Sub Form_Activate()
    Me.Caption = "NIS SERVICE"
    Select Case ShowForm
        Case "Rujukan"
            Ket.Caption = ".:: List Rujukan Heavy Repair ::."
            Call ViewFieldRujukan
        Case "Model"
            Ket.Caption = ".:: List Model ::."
            Call ViewFieldModel
        Case "Warna"
            Ket.Caption = ".:: List Warna ::."
            Call ViewFieldWarna
        Case "Custdat"
            Ket.Caption = ".:: List Customer ::."
            Call ViewFieldCustomer
    End Select
    cboS.SetFocus
End Sub
'222222222222222222222222222222222222222222222222222222222222222222222222222222222222222
Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo ErrNull
    If rsZ.RecordCount > 0 Then
        Pilih.text = Trim(rsZ.Fields(0).value)
    Else
        Pilih.text = ""
    End If
ErrNull:
End Sub
'=========================================================================================================
'3333333333333333333333333333333333333333333333333333333333333333333333333333
Private Sub ViewFieldRujukan()
    On Error GoTo err
    Set rsZ = objO.FetchData("Select * from RUJUKAN WITH (NOLOCK) where NoRujuk = 'ANTZ'")
    Call SetField(rsZ, CInt(rsZ.Fields.Count))
err:
End Sub
Private Sub ViewFieldModel()
    On Error GoTo err
    
    If Log.RPLHistori = 0 Then
        Set rsZ = objO.FetchData("select MDL, MODEL, [DESC] as Nama from Model WITH (NOLOCK) Where MDL = 'ANTZ'")
    Else
        Set rsZ = objO.FetchData("select MDL, MODEL, [DESC] as Nama from [" & dbReplikasiHO & "]..Model WITH (NOLOCK) Where MDL = 'ANTZ'")
    End If
Call SetField(rsZ, CInt(rsZ.Fields.Count))
err:
End Sub
Private Sub ViewFieldWarna()
    On Error GoTo err
    Set rsZ = objO.FetchDataSPK("Select KodeWarna, NamaInd as Nama from INWarnaEXT WITH (NOLOCK) Where KodeWarna = 'ANTZ'")
Call SetField(rsZ, CInt(rsZ.Fields.Count))
err:
End Sub
Private Sub ViewFieldCustomer()
    On Error GoTo err
    Set rsZ = objO.FetchData("select * from Custdat WITH (NOLOCK) Where NO_Chassis = 'ANTZ'")
Call SetField(rsZ, CInt(rsZ.Fields.Count) - 1)
err:
End Sub
'====================
'Set Field Set Field
Private Sub SetField(nRS As ADODB.Recordset, nFilter As Integer)
    With Me.cboS
        .Clear
        For i = 0 To nFilter
           .AddItem nRS.Fields(i).Name
        Next i
    End With
End Sub


'444444444444444444444444444444444444444444444444444444444444444444444444444444444444
Private Sub ViewFieldRujukan_Search(nParam As String)
    On Error GoTo err
    Set rsZ = objO.FetchData("Select top 25 * from RUJUKAN WITH (NOLOCK) " & nParam)
    Set Grid.DataSource = rsZ
    If rsZ.RecordCount > 0 Then
        rsZ.MoveFirst
        Pilih.text = Trim(rsZ.Fields(0).value)
    Else
        Pilih.text = ""
    End If
err:
End Sub
Private Sub ViewFieldModel_Search(nParam As String)
    On Error GoTo err
    
    If Log.RPLHistori = 0 Then
        Set rsZ = objO.FetchData("select top 25 MDL, MODEL, [DESC] as Nama from Model WITH (NOLOCK) " & nParam)
    Else
        Set rsZ = objO.FetchData("select top 25 MDL, MODEL, [DESC] as Nama from [" & dbReplikasiHO & "]..Model WITH (NOLOCK) " & nParam)
    End If
    Set Grid.DataSource = rsZ
    If rsZ.RecordCount > 0 Then
        rsZ.MoveFirst
        Pilih.text = Trim(rsZ.Fields(0).value)
    Else
        Pilih.text = ""
    End If
err:
End Sub

Private Sub ViewFieldWarna_Search(nParam As String)
    On Error GoTo err
    Set rsZ = objO.FetchDataSPK("Select top 25 KodeWarna, NamaInd as Nama from INWarnaEXT WITH (NOLOCK) " & nParam)
    Set Grid.DataSource = rsZ
    If rsZ.RecordCount > 0 Then
        rsZ.MoveFirst
        Pilih.text = Trim(rsZ.Fields(0).value)
    Else
        Pilih.text = ""
    End If
err:
End Sub
Private Sub ViewFieldCustdat_Search(nParam As String)
    On Error GoTo err
    Set rsZ = objO.FetchData("select top 25 * from Custdat WITH (NOLOCK) " & nParam)
    Set Grid.DataSource = rsZ
    If rsZ.RecordCount > 0 Then
        rsZ.MoveFirst
        Pilih.text = Trim(rsZ.Fields(0).value)
    Else
        Pilih.text = ""
    End If
err:
End Sub
'==========================================================================================================

Private Sub Pilih_GotFocus()
    Sendkeys "{tab}"
End Sub
Private Sub Grid_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{tab}"
End Sub
