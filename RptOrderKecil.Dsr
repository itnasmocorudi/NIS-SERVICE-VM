VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptOrderKecil 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptOrderKecil.dsx":0000
End
Attribute VB_Name = "RptOrderKecil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = str(Val(Urut) + 1)
End Sub

Private Sub GroupFooter2_Format()
    LblTotal = "TOTAL INV " & gKetTemp & " : "
End Sub

Private Sub PageHeader_Format()
    Perush = gPerush
    Alamat = gAlamat
End Sub
