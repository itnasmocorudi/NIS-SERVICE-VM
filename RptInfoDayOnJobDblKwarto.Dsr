VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptInfoDayOnJobDblKwarto 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptInfoDayOnJobDblKwarto.dsx":0000
End
Attribute VB_Name = "RptInfoDayOnJobDblKwarto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
    Urut = str(Val(Urut) + 1)
    Select Case Trim(Field59.Text)
        Case " "
            Field59.Text = "TUNGGU"
        Case "P"
            Field59.Text = "PROSES"
        Case "B"
            Field59.Text = "PAUSE"
        Case "S"
            Field59.Text = "SELESAI"
        Case "K"
            Field59.Text = "KELUAR"
        Case Else
            Field59.Text = ""
    End Select
    Select Case Trim(Field58.Text)
        Case "BTL"
            Field58.Text = "BATAL"
        Case Else
            Field58.Text = "OKE"
    End Select
End Sub

Private Sub GroupFooter2_Format()
    LblTotal = "TOTAL INV " & gKetTemp & " : "
End Sub

Private Sub GroupFooter3_Format()
    If Field40.Text = "" Then
        GroupFooter3.Visible = False
    Else
        GroupFooter3.Visible = True
    End If
End Sub

Private Sub GroupFooter4_Format()
    If Field44.Text = "" Then
        GroupFooter4.Visible = False
    Else
        GroupFooter4.Visible = True
    End If
End Sub

Private Sub GroupFooter5_Format()
    If Field42.Text = "" Then
        GroupFooter5.Visible = False
    Else
        GroupFooter5.Visible = True
    End If
End Sub

Private Sub GroupFooter6_Format()
    If Field46.Text = "" Then
        GroupFooter6.Visible = False
    Else
        GroupFooter6.Visible = True
    End If
End Sub

Private Sub GroupFooter7_Format()
    If Field48.Text = "" Then
        GroupFooter7.Visible = False
    Else
        GroupFooter7.Visible = True
    End If
End Sub

Private Sub GroupFooter8_Format()
    If Field52.Text = "" Then
        GroupFooter8.Visible = False
    Else
        GroupFooter8.Visible = True
    End If
End Sub

Private Sub GroupFooter9_Format()
    If Field50.Text = "" Then
        GroupFooter9.Visible = False
    Else
        GroupFooter9.Visible = True
    End If
End Sub

Private Sub PageHeader_Format()
    TglService = "TANGGAL : " & gTglTemp & " " & gKetTemp
    Perush = gPerush
    Alamat = gAlamat
    Tanggal = gTanggal & " " & Format(Time, "HH:NN")
'    LblPage = "Halaman : " & Str(Me.pageNumber) & " dari " & Str(Me.Printer.)
End Sub
