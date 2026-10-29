VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptInfoBook 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   33443
   _ExtentY        =   19315
   SectionData     =   "RptInfoBook.dsx":0000
End
Attribute VB_Name = "RptInfoBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vBtl As Integer
Dim vBok As Integer
Dim vOke As Integer

Private Sub Detail_Format()
    Urut = str(Val(Urut) + 1)
    Select Case Field41.DataValue
        Case "OKE"
            vOke = vOke + 1
        Case "BTL"
            vBtl = vBtl + 1
        Case "BOK"
            vBok = vBok + 1
    End Select
End Sub

Private Sub GroupFooter2_Format()
    Field42.DataValue = vOke + vBtl + vBok
    Field43.DataValue = vBok
    Field44.DataValue = vOke
    Field45.DataValue = vBtl
End Sub

Private Sub PageHeader_Format()
    TglService = gKetTemp & " : " & gTglTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
    vBtl = 0
    vBok = 0
    vOke = 0
End Sub
