VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapBulanOrderBengkel 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapBulanOrderBengkel.dsx":0000
End
Attribute VB_Name = "RptLapBulanOrderBengkel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
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
