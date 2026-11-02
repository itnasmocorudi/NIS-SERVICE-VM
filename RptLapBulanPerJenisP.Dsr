VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapBulanPerJenisP 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapBulanPerJenisP.dsx":0000
End
Attribute VB_Name = "RptLapBulanPerJenisP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
    Field38.Visible = True
    Field39.Visible = True
    Field40.Visible = True
    Field41.Visible = True
    Field42.Visible = True
    Field43.Visible = True
    Field44.Visible = True
    
    Field45.Visible = True
    Field46.Visible = True
    Field47.Visible = True
    Field48.Visible = True
    Field49.Visible = True
    Field50.Visible = True
    Field51.Visible = True
    
    If IsNull(Field38.DataValue) Then
        Field38.Visible = False
        Field45.Visible = False
    End If
    If IsNull(Field39.DataValue) Then
        Field39.Visible = False
        Field46.Visible = False
    End If
    If IsNull(Field40.DataValue) Then
        Field40.Visible = False
        Field47.Visible = False
    End If
    If IsNull(Field41.DataValue) Then
        Field41.Visible = False
        Field48.Visible = False
    End If
    If IsNull(Field42.DataValue) Then
        Field42.Visible = False
        Field49.Visible = False
    End If
    If IsNull(Field43.DataValue) Then
        Field43.Visible = False
        Field50.Visible = False
    End If
    If IsNull(Field44.DataValue) Then
        Field44.Visible = False
        Field51.Visible = False
    End If
    If Field39.Text = "" Then
        Detail.Height = Detail.Height - 180
    End If
    If Field40.Text = "" Then
        Detail.Height = Detail.Height - 180
    End If
    If Field41.Text = "" Then
        Detail.Height = Detail.Height - 180
    End If
    If Field42.Text = "" Then
        Detail.Height = Detail.Height - 180
    End If
    If Field43.Text = "" Then
        Detail.Height = Detail.Height - 180
    End If
    If Field44.Text = "" Then
        Detail.Height = Detail.Height - 180
    End If
    If Field36.DataValue = "01/01/3000" Then
        Field36.DataValue = "  /  /    "
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
