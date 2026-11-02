VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapBulanMointorUnitRCPol 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapBulanMointorUnitRCPol.dsx":0000
End
Attribute VB_Name = "RptLapBulanMointorUnitRCPol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vOldNoPol As String
Dim vNoPol As String

Private Sub Detail_Format()
    vNoPol = Field29.DataValue
    If vOldNoPol <> vNoPol Then
        Line3.Visible = True
        Urut.Visible = True
        Urut = Str(Val(Urut) + 1)
    Else
        Line3.Visible = False
        Field29.DataValue = ""
        Field28.DataValue = ""
        Urut.Visible = False
    End If
    vOldNoPol = vNoPol
    
    If Field50.DataValue = "01/01/3000" Then
        Field50.DataValue = ""
    End If
    If Field49.DataValue = "01/01/3000" Then
        Field49.DataValue = ""
    End If
    If Field48.DataValue = "01/01/3000" Then
        Field48.DataValue = ""
    End If
    If Field47.DataValue = "01/01/3000" Then
        Field47.DataValue = ""
    End If
    If Field46.DataValue = "01/01/3000" Then
        Field46.DataValue = ""
    End If
    If Field45.DataValue = "01/01/3000" Then
        Field45.DataValue = ""
    End If
    If Field44.DataValue = "01/01/3000" Then
        Field44.DataValue = ""
    End If
    If Field43.DataValue = "01/01/3000" Then
        Field43.DataValue = ""
    End If
    If Field42.DataValue = "01/01/3000" Then
        Field42.DataValue = ""
    End If
    If Field41.DataValue = "01/01/3000" Then
        Field41.DataValue = ""
    End If
    If Field40.DataValue = "01/01/3000" Then
        Field40.DataValue = ""
    End If
    If Field39.DataValue = "01/01/3000" Then
        Field39.DataValue = ""
    End If
    If Field38.DataValue = "01/01/3000" Then
        Field38.DataValue = ""
    End If
    If Field37.DataValue = "01/01/3000" Then
        Field37.DataValue = ""
    End If
    If Field36.DataValue = "01/01/3000" Then
        Field36.DataValue = ""
    End If
    If Field35.DataValue = "01/01/3000" Then
        Field35.DataValue = ""
    End If
    If Field35.DataValue = "01/01/3000" Then
        Field35.DataValue = ""
    End If
    If Field34.DataValue = "01/01/3000" Then
        Field34.DataValue = ""
    End If
    If Field33.DataValue = "01/01/3000" Then
        Field33.DataValue = ""
    End If
    If Field32.DataValue = "01/01/3000" Then
        Field32.DataValue = ""
    End If
    If Field31.DataValue = "01/01/3000" Then
        Field31.DataValue = ""
    End If
    If Field30.DataValue = "01/01/3000" Then
        Field30.DataValue = ""
    End If
End Sub

Private Sub PageHeader_Format()
    TglService = "TANGGAL : " & gTglTemp
    KodeSA = gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
End Sub
