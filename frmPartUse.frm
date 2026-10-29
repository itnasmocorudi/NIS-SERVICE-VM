VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmPartUse 
   Caption         =   "Part"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Keluar"
      Height          =   390
      Left            =   6750
      TabIndex        =   0
      Top             =   3885
      Width           =   1890
   End
   Begin TrueDBGrid70.TDBGrid TDBGrid1 
      Height          =   3450
      Left            =   15
      TabIndex        =   1
      Top             =   30
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   6085
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Part"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Qty"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Satuan"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Sub Total"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2884"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2805"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=5133"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5054"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=1138"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1058"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2223"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2143"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2778"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2699"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      AllowArrows     =   0   'False
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
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
      _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   6855
      TabIndex        =   2
      Top             =   3525
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   ">"
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   75
      TabIndex        =   4
      Top             =   3585
      Width           =   6030
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   255
      Left            =   5790
      TabIndex        =   3
      Top             =   3585
      Width           =   945
   End
End
Attribute VB_Name = "frmPartUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyArray As New XArrayDB
Dim NoWO As String
Dim vNamaTabel As String
Dim LOGONSTR2 As String

Private Sub Form_Activate()
    Dim i As Byte
    Dim vTotal As Double
    Dim vHitung As Double
    Dim JmlBrs As Byte
    Dim vNo_Faktur As String
    Dim vTgl_Faktur As String
    
    vNamaTabel = "PART" & DriveSerialNumber("c:")
    
    MyArray.ReDim -1, -1, -1, -1
    
    TDBGrid1.Columns(0).Value = False
    NoWO = Trim(frmOrder.Text(26))
    
    On Error GoTo Salah
    If gSudahInv Then
       ' LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "SIDATA" & Trim(gBlnBedaPart) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
       ''LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
    Else
       ' LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "INDATA" & Trim(gGenerasi) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
       ''LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
    End If
    
    If Get_ConnectionAcces(NisDb, NisUSER, NisPass, NisKomp) Then
    ''oConnAcces.ConnectionString = LOGONSTR2
    ''oConnAcces.CommandTimeout = 120
    ''oConnAcces.ConnectionTimeout = 120
    ''oConnAcces.Open
    
    On Error GoTo Salahlagi
        
        i = 0
        If gSudahInv Then
            sSQL = "Select count(*) AS JMLBaris from SISAL AS A INNER JOIN SIHDR AS B ON B.NoFaktur= A.NoFaktur where left(B.NoSo, 7) = '" & NoWO & "' AND KODELGN='" & Trim(frmOrder.Text(57)) & "'"
            oRs.Open sSQL, oConnAcces, adOpenStatic
            If Not oRs.BOF Then
                If Val(oRs("JMLBaris")) = 0 Then
                    JmlBrs = 0
                Else
                    JmlBrs = Val(oRs("JMLBaris")) - 1
                End If
            End If
            oRs.Close
            MyArray.ReDim 0, JmlBrs, 0, 4
            
            sSQL = "Select distinct B.NoSo from SISAL AS A INNER JOIN SIHDR AS B ON B.NoFaktur= A.NoFaktur where left(B.NoSo, 7) = '" & NoWO & "' AND KODELGN='" & Trim(frmOrder.Text(57)) & "'"
            oRs1.Open sSQL, oConnAcces, adOpenStatic
            If Not oRs1.BOF Then
                Do While Not oRs1.EOF
                    sSQL = "Select B.NoFaktur, B.TglFaktur, B.TypePpn, B.ExtDisPersen, A.qty, A.ItemDisPsn, A.HargaJual, A.Keterangan, A.KodeItem " _
                         & " from SISAL AS A INNER JOIN SIHDR AS B ON B.NoFaktur= A.NoFaktur where B.NoSo = '" & oRs1("NoSo") & "' AND KODELGN='" & Trim(frmOrder.Text(57)) & "'"
                    oRs.Open sSQL, oConnAcces, adOpenStatic
                    If Not oRs.BOF Then
                        vNo_Faktur = oRs("NoFaktur")
                        vTgl_Faktur = oRs("TglFaktur")
                        Do While Not oRs.EOF
                            MyArray.Value(i, 0) = Trim(oRs("KodeItem"))
                            MyArray.Value(i, 1) = Trim(oRs("Keterangan"))
                            MyArray.Value(i, 2) = Format(oRs("qty"), "#,##0.00")
                            MyArray.Value(i, 3) = Format(oRs("HargaJual"), "#,##0.00")
                            vHitung = (Koma2Titik(oRs("Qty")) * Koma2Titik(oRs("HargaJual")))
                            vHitung = vHitung - (Koma2Titik(oRs("ItemDisPsn")) * vHitung)
                            vHitung = vHitung - (Koma2Titik(oRs("ExtDisPersen")) * vHitung)
                            MyArray.Value(i, 4) = Format(vHitung, "#,##0.00")
                            vTotal = vTotal + vHitung
                            oRs.MoveNext
                            i = i + 1
                        Loop
                    End If
                    oRs.Close
                    oRs1.MoveNext
                Loop
            End If
            oRs1.Close
        Else
            sSQL = "SELECT DISTINCT Inmst.NoItem, Inmst.NamaBarang, [Inmst].[QtySaldoAwal]+[Inmst].[QtyTerimabln]+[Inmst].[QtyKeluarBln] AS Qty, Inmst.HargaJualA, [Inmst].[HargaJualA]*([Inmst].[QtySaldoAwal]+[Inmst].[QtyTerimabln]+[Inmst].[QtyKeluarBln]) AS Total from INMST"
            sSQL = sSQL & " WHERE (([Inmst].[QtySaldoAwal]+[Inmst].[QtyTerimabln]+[Inmst].[QtyKeluarBln])<>0) AND "
            sSQL = sSQL & " (Inmst.KodeGdg='" & gGudang & "') AND (left(Inmst.NoRack,7)='" & NoWO & "')"
            'MsgBox sSQL

            oRs.Open sSQL, oConnAcces, adOpenStatic
            If Not oRs.BOF Then
                MyArray.ReDim 0, oRs.RecordCount - 1, 0, 4
                Do While Not oRs.EOF
                    MyArray.Value(i, 0) = Trim(oRs("NoItem"))
                    MyArray.Value(i, 1) = Trim(oRs("NamaBarang"))
                    MyArray.Value(i, 2) = Format(oRs("qty"), "#,##0.00##")
                    MyArray.Value(i, 3) = Format(oRs("HargaJualA"), "#,##0.00##")
                    MyArray.Value(i, 4) = Format(oRs("Total"), "#,##0.00##")
                    vTotal = vTotal + oRs("Total")
                    oRs.MoveNext
                    i = i + 1
                Loop
            End If
            oRs.Close
        End If
        
        Text(1) = Format(vTotal, "#,##0.00")
        TDBGrid1.Columns(4).Alignment = dbgRight
        TDBGrid1.Columns(3).Alignment = dbgRight
        TDBGrid1.Columns(2).Alignment = dbgRight
        TDBGrid1.Columns(4).AllowFocus = False
        TDBGrid1.Columns(3).AllowFocus = False
        TDBGrid1.Columns(2).AllowFocus = False
        TDBGrid1.Columns(1).AllowFocus = False
        TDBGrid1.Columns(0).AllowFocus = False
        TDBGrid1.Array = MyArray
        TDBGrid1.ReBind
    
    End If
    oConnAcces.Close
    If gSudahInv Then
        Label1.Caption = "Part sudah ditutup dgn No Invoice : " & vNo_Faktur & " Tgl : " & vTgl_Faktur
    Else
        Label1.Caption = "Part belum ditutup"
    End If
    
    Exit Sub
Salah:
    MsgBox "Tidak memakai part!"
    Exit Sub
Salahlagi:
    MsgBox "Error pengambilan data!"
    Exit Sub
End Sub

Private Sub Form_Load()
    Call CenterForm(Me, True)
End Sub

Private Sub TblKeluar_Click()
    On Error Resume Next
    Unload Me
    oConnAcces.Close
End Sub

Private Sub TDBGrid1_Validate(Cancel As Boolean)
    TDBGrid1.Update
End Sub
