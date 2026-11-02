VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapHarianSOMasuk 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   20267
   SectionData     =   "RptLapHarianSOMasuk.dsx":0000
End
Attribute VB_Name = "RptLapHarianSOMasuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vOrder As Integer
Dim vJOrder As Integer
Dim vRate As Double
Dim vJRate As Double

Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
    vRate = Val(Field44.DataValue) + Val(Field45.DataValue) + Val(Field46.DataValue) _
          + Val(Field47.DataValue) + Val(Field48.DataValue) + Val(Field49.DataValue) _
          + Val(Field50.DataValue) + Val(Field51.DataValue)
    
    vOrder = 0
    If Field42.Text <> "" Then
        vOrder = vOrder + 1
    End If
    If Field52.Text <> "" Then
        vOrder = vOrder + 1
    End If
    If Field53.Text <> "" Then
'        Detail.Height = Detail.Height - 180
'    Else
        vOrder = vOrder + 1
    End If
    If Field54.Text = "" Then
        Detail.Height = Detail.Height - 180
    Else
        vOrder = vOrder + 1
    End If
    If Field55.Text = "" Then
        Detail.Height = Detail.Height - 180
    Else
        vOrder = vOrder + 1
    End If
    If Field56.Text = "" Then
        Detail.Height = Detail.Height - 180
    Else
        vOrder = vOrder + 1
    End If
    If Field57.Text = "" Then
        Detail.Height = Detail.Height - 180
    Else
        vOrder = vOrder + 1
    End If
    JOrd.Text = vOrder
    vJOrder = vJOrder + vOrder
    Jrate.Text = vRate
    vJRate = vJRate + vRate
End Sub

Private Sub GroupFooter2_Format()
    Field72.Text = vJOrder
    Field73.Text = vJRate
End Sub

Private Sub PageHeader_Format()
    TglService = "TANGGAL : " & gTglTemp & " " & gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
End Sub
