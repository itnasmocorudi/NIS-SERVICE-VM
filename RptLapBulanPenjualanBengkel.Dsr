VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapBulanPenjualanBengkel 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapBulanPenjualanBengkel.dsx":0000
End
Attribute VB_Name = "RptLapBulanPenjualanBengkel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTotOrder As Double
Dim vGTotOrder As Double

Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
'    vTotOrder = Val(Field140.DataValue) + Val(Field141.DataValue) _
              + Val(Field142.DataValue) + Val(Field143.DataValue) _
              + Val(Field144.DataValue) + Val(Field145.DataValue) _
              + Val(Field146.DataValue) + Val(Field147.DataValue) _
              + Val(Field172.DataValue)
'    Field130.Text = vTotOrder
'    vGTotOrder = vGTotOrder + vTotOrder
End Sub

Private Sub GroupFooter2_Format()
'    Field149.Text = vGTotOrder
End Sub

Private Sub GroupHeader3_Format()
    If Field168.DataValue = 0 Then
        GroupHeader3.Visible = False
    Else
        GroupHeader3.Visible = True
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
