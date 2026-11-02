VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} RptLapHarianPenjualanSA 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "RptLapHarianPenjualanSA.dsx":0000
End
Attribute VB_Name = "RptLapHarianPenjualanSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vJOrder As Integer
Dim vSubJOrder As Integer
Dim vRate As Double
Dim vJRate As Double
Dim vSubJRate As Double
Dim vBOR As Double
Dim vTBOR As Double
Dim vELC As Double
Dim vTELC As Double
Dim vERC As Double
Dim vTERC As Double
Dim vGRP As Double
Dim vTGRP As Double
Dim vIRC As Double
Dim vTIRC As Double
Dim vLUB As Double
Dim vTLUB As Double
Dim vTWC As Double
Dim vTTWC As Double
Dim vWAS As Double
Dim vTWAS As Double
Dim vPDI As Double
Dim vTPDI As Double

Private Sub HitungPerKelompok(KlpKerja As String, Nilai As Double)
    Select Case Trim(KlpKerja)
        Case "BOR"
            Field75.DataValue = Field75.DataValue + Nilai
            vBOR = vBOR + Nilai
            vTBOR = vTBOR + Nilai
        Case "RTJ"
            Field76.DataValue = Field76.DataValue + Nilai
            vELC = vELC + Nilai
            vTELC = vTELC + Nilai
        Case "ERC"
            Field77.DataValue = Field77.DataValue + Nilai
            vERC = vERC + Nilai
            vTERC = vTERC + Nilai
        Case "GRP"
            Field78.DataValue = Field78.DataValue + Nilai
            vGRP = vGRP + Nilai
            vTGRP = vTGRP + Nilai
        Case "IRC"
            Field79.DataValue = Field79.DataValue + Nilai
            vIRC = vIRC + Nilai
            vTIRC = vTIRC + Nilai
        Case "LUB"
            Field80.DataValue = Field80.DataValue + Nilai
            vLUB = vLUB + Nilai
            vTLUB = vTLUB + Nilai
        Case "TWC"
            Field81.DataValue = Field81.DataValue + Nilai
            vTWC = vTWC + Nilai
            vTTWC = vTTWC + Nilai
        Case "WAS"
            Field82.DataValue = Field82.DataValue + Nilai
            vWAS = vWAS + Nilai
            vTWAS = vTWAS + Nilai
        Case "PDI"
            Field145.DataValue = Field145.DataValue + Nilai
            vPDI = vPDI + Nilai
            vTPDI = vTPDI + Nilai
    End Select
End Sub

Private Sub Detail_Format()
    Urut = Str(Val(Urut) + 1)
    vRate = Val(Field44.DataValue) + Val(Field45.DataValue) + Val(Field46.DataValue) _
          + Val(Field47.DataValue) + Val(Field48.DataValue) + Val(Field49.DataValue) _
          + Val(Field50.DataValue) + Val(Field51.DataValue)
    
    vJOrder = vJOrder + JOrd.DataValue
    vSubJOrder = vSubJOrder + JOrd.DataValue
    Jrate.DataValue = vRate
    vJRate = vJRate + vRate
    vSubJRate = vSubJRate + vRate
    Field75.DataValue = 0
    Field76.DataValue = 0
    Field77.DataValue = 0
    Field78.DataValue = 0
    Field79.DataValue = 0
    Field80.DataValue = 0
    Field81.DataValue = 0
    Field82.DataValue = 0
    Field145.DataValue = 0
    Call HitungPerKelompok(Field93.Text, Val(Field100.Text))
    Call HitungPerKelompok(Field94.Text, Val(Field101.Text))
    Call HitungPerKelompok(Field95.Text, Val(Field102.Text))
    Call HitungPerKelompok(Field96.Text, Val(Field103.Text))
    Call HitungPerKelompok(Field97.Text, Val(Field104.Text))
    Call HitungPerKelompok(Field98.Text, Val(Field105.Text))
    Call HitungPerKelompok(Field99.Text, Val(Field106.Text))
End Sub

Private Sub GroupFooter2_Format()
    Field107.DataValue = vTBOR
    Field108.DataValue = vTELC
    Field109.DataValue = vTERC
    Field110.DataValue = vTGRP
    Field111.DataValue = vTIRC
    Field112.DataValue = vTLUB
    Field113.DataValue = vTTWC
    Field114.DataValue = vTWAS
    Field147.DataValue = vTPDI
    vBOR = 0
    vELC = 0
    vERC = 0
    vGRP = 0
    vIRC = 0
    vLUB = 0
    vTWC = 0
    vWAS = 0
    vTBOR = 0
    vTELC = 0
    vTERC = 0
    vTGRP = 0
    vTIRC = 0
    vTLUB = 0
    vTTWC = 0
    vTWAS = 0
    vTPDI = 0
    Field72.DataValue = vJOrder
    Field73.DataValue = vJRate
End Sub

Private Sub GroupFooter3_Format()
    Field124.DataValue = vSubJOrder
    Field125.DataValue = vSubJRate
    Field126.DataValue = vBOR
    Field127.DataValue = vELC
    Field128.DataValue = vERC
    Field129.DataValue = vGRP
    Field130.DataValue = vIRC
    Field131.DataValue = vLUB
    Field132.DataValue = vTWC
    Field133.DataValue = vWAS
    Field146.DataValue = vPDI
    vBOR = 0
    vELC = 0
    vERC = 0
    vGRP = 0
    vIRC = 0
    vLUB = 0
    vTWC = 0
    vWAS = 0
    vPDI = 0
    Label48.Caption = "Sub Total " & Field143.Text & " :"
    vSubJRate = 0
    vSubJOrder = 0
End Sub

Private Sub GroupHeader4_Format()
    If Field144.DataValue = 0 Then
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
