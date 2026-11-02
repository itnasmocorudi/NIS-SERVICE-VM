VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapBulanMointorUnitRC 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapBulanMointorUnitRC.dsx":0000
End
Attribute VB_Name = "RptLapBulanMointorUnitRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vGTotal As Double
Dim vTotal As Double

Private Sub Detail_Format()
    
    vTotal = Val(Field30.DataValue) + Val(Field31.DataValue) + Val(Field32.DataValue) _
           + Val(Field33.DataValue) + Val(Field34.DataValue) + Val(Field35.DataValue) _
           + Val(Field36.DataValue) + Val(Field37.DataValue) + Val(Field38.DataValue) _
           + Val(Field39.DataValue) + Val(Field40.DataValue) + Val(Field41.DataValue) _
           + Val(Field42.DataValue) + Val(Field43.DataValue) + Val(Field44.DataValue) _
           + Val(Field45.DataValue) + Val(Field46.DataValue) + Val(Field47.DataValue) _
           + Val(Field48.DataValue) + Val(Field49.DataValue) + Val(Field50.DataValue)
    If vTotal > 0 Or Val(Field74.Text) > 0 Then
        Detail.Visible = True
        Urut = Str(Val(Urut) + 1)
        Field51.Text = vTotal
        vGTotal = vGTotal + vTotal
    Else
        Detail.Visible = False
    End If
End Sub

Private Sub GroupFooter2_Format()
    Field73.Text = vGTotal
End Sub

Private Sub PageHeader_Format()
    TglService = "TANGGAL : " & gTglTemp
    KodeSA = gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
End Sub
