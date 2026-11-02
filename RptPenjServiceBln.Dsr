VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptPenjServiceBln 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   12568
   SectionData     =   "RptPenjServiceBln.dsx":0000
End
Attribute VB_Name = "RptPenjServiceBln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
End Sub

Private Sub PageHeader_Format()
    If Me.pageNumber = 1 Then
        DIPERIKSA.Visible = True
    Else
        DIPERIKSA.Visible = False
    End If
    TglService = gTanggal & " " & Format(Time, "HH:NN")
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = "Tanggal : " & gTglTemp
End Sub
