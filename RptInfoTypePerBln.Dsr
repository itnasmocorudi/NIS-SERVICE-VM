VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptInfoTypePerBln 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "RptInfoTypePerBln.dsx":0000
End
Attribute VB_Name = "RptInfoTypePerBln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = str(Val(Urut) + 1)
End Sub

Private Sub PageHeader_Format()
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
    Label32.Caption = gTglTemp
End Sub
