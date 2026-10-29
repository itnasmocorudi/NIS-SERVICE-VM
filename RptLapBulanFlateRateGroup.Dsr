VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapBulanFlateRateGroup 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapBulanFlateRateGroup.dsx":0000
End
Attribute VB_Name = "RptLapBulanFlateRateGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
    Field88.DataValue = Str(gReturn * Field64.DataValue)
End Sub

Private Sub GroupFooter2_Format()
    Field90.DataValue = Str(gReturn * Field78.DataValue)
End Sub

Private Sub GroupFooter3_Format()
    Field89.DataValue = Str(gReturn * Field71.DataValue)
End Sub

Private Sub PageHeader_Format()
    Field83.Text = gTglTemp
    Field84.Text = gTglTemp1
    Field85.Text = gTglTemp2
    KodeSA = gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
End Sub
