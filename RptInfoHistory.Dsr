VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptInfoHistory 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   33655
   _ExtentY        =   19315
   SectionData     =   "RptInfoHistory.dsx":0000
End
Attribute VB_Name = "RptInfoHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = str(Val(Urut) + 1)
    If Field53.text = "" Then
        Detail.Height = Detail.Height - 270
    End If
    If Field51.text = "" Then
        Detail.Height = Detail.Height - 270
    End If
    If Field49.text = "" Then
        Detail.Height = Detail.Height - 270
    End If
    If Field47.text = "" Then
        Detail.Height = Detail.Height - 270
    End If
    If Field45.text = "" Then
        Detail.Height = Detail.Height - 270
    End If
    If Field43.text = "" Then
        Detail.Height = Detail.Height - 270
    End If
'    If Field40.Text = "" Then
'        Detail.Height = Detail.Height - 270
'    End If
    Line3.Y1 = Detail.Height - 10
    Line3.Y2 = Detail.Height - 10
End Sub

Private Sub PageHeader_Format()
'    TglService = "TANGGAL : " & gTglTemp & " " & gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
End Sub
