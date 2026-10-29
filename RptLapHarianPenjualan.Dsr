VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapHarianPenjualan 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapHarianPenjualan.dsx":0000
End
Attribute VB_Name = "RptLapHarianPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
    If Field182.DataValue <> "" Then
        Jrate.DataValue = ""
        Field82.DataValue = ""
        Field83.DataValue = ""
        Field84.DataValue = ""
        Field85.DataValue = ""
        Field86.DataValue = ""
        Field87.DataValue = ""
        Field88.DataValue = ""
        Field89.DataValue = ""
        Field90.DataValue = ""
        Field91.DataValue = ""
        Field162.DataValue = ""
        Field163.DataValue = ""
    End If
End Sub

Private Sub GroupFooter2_Format()
    Label56.Caption = Trim(gTglTemp) & " : " & Trim(Val(Field181.Text)) & " dengan jumlah SO/Mobil terinvoice : " & Trim(Val(Field180.Text))
End Sub

Private Sub GroupHeader2_Format()
    If Field182.DataValue = "BTL" Then
        GroupFooter2.Visible = False
    Else
        GroupFooter2.Visible = True
    End If
End Sub

Private Sub GroupHeader3_Format()
    If Field182.DataValue = "BTL" Then
        GroupHeader3.Visible = True
    Else
        GroupHeader3.Visible = False
    End If
End Sub

Private Sub GroupHeader4_Format()
    If Field129.DataValue = 0 Then
        GroupHeader4.Visible = False
    Else
        GroupHeader4.Visible = True
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
