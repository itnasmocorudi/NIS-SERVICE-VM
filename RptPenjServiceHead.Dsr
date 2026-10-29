VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptPenjServiceHead 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptPenjServiceHead.dsx":0000
End
Attribute VB_Name = "RptPenjServiceHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
End Sub

Private Sub GroupFooter2_Format()
    LblTotal = "TOTAL INV " & gKetTemp & " : "
    Label16 = "TOTAL " & gKetTemp & " Rp : "
End Sub

Private Sub GroupHeader2_Format()
    If Status <> "" Then
        GroupFooter2.Visible = False
    Else
        GroupFooter2.Visible = True
    End If
End Sub

Private Sub PageHeader_Format()
    If Me.pageNumber = 1 Then
        DIPERIKSA.Visible = True
        DISETUJUI.Visible = True
    Else
        DIPERIKSA.Visible = False
        DISETUJUI.Visible = False
    End If
    TglService = "TANGGAL : " & gTglTemp & " " & gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
End Sub
