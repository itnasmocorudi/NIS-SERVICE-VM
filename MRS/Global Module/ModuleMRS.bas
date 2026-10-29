Attribute VB_Name = "ModuleMRS"
'Option Explicit
'Private Const PathMrs32 As String = "C:\WINDOWS\system32\Mrs32.Ant"
'Public SQLCon        As String
'Public SQLConSPK     As String
'
'Type LoginServer
'     LogServ         As String
'     LogUser         As String
'     LogPass         As String
'     LogNas          As String
'     LogAlamat       As String
'     LogBengkel      As String
'     LogSPK          As String
'End Type
'Public Log          As LoginServer
'
'Public Const AccessCon As String = "Provider = Microsoft.JET.OLEDB.4.0;" & _
'"Data Source =  C:\Inetpub\wwwroot\sms\dbSMS.mdb"
'
'Public IsNow As Date
'
'Sub Main()
'   Dim IsText
'   Dim A As Integer
'   Open PathMrs32 For Input As #1
'   Do Until EOF(1)
'   Line Input #1, IsText
'      A = A + 1
'      Select Case A
'         Case 1: Log.LogServ = IsText
'         Case 2: Log.LogUser = IsText
'         Case 3: Log.LogPass = IsText
'         Case 4: Log.LogNas = IsText
'         Case 5: Log.LogAlamat = IsText
'         Case 6: Log.LogBengkel = IsText
'         Case 7: Log.LogSPK = IsText
'      End Select
'      If A = 7 Then Exit Do
'   Loop
'   Close #1
'
'   Log.LogServ = Crypt(Log.LogServ, "antztein", False)
'      Log.LogServ = Crypt(Log.LogServ, "tandayu", False)
'   Log.LogUser = Crypt(Log.LogUser, "antztein", False)
'      Log.LogUser = Crypt(Log.LogUser, "tandayu", False)
'   Log.LogPass = Crypt(Log.LogPass, "antztein", False)
'      Log.LogPass = Crypt(Log.LogPass, "tandayu", False)
'   Log.LogNas = Crypt(Log.LogNas, "antztein", False)
'      Log.LogNas = Crypt(Log.LogNas, "tandayu", False)
'   Log.LogAlamat = Crypt(Log.LogAlamat, "antztein", False)
'      Log.LogAlamat = Crypt(Log.LogAlamat, "tandayu", False)
'   Log.LogBengkel = Crypt(Log.LogBengkel, "antztein", False)
'      Log.LogBengkel = Crypt(Log.LogBengkel, "tandayu", False)
'   Log.LogSPK = Crypt(Log.LogSPK, "antztein", False)
'      Log.LogSPK = Crypt(Log.LogSPK, "tandayu", False)
'   SQLCon = "Provider = SQLOLEDB.1; User ID=" & Log.LogUser & "; Password=" & Log.LogPass & "; Initial Catalog=" & Log.LogBengkel & "; Data Source =" & Log.LogServ
'   SQLConSPK = "Provider = SQLOLEDB.1; User ID=" & Log.LogUser & "; Password=" & Log.LogPass & "; Initial Catalog=" & Log.LogSPK & "; Data Source =" & Log.LogServ
'
'   Call SetTimerNow
'   Dim f As New FormMainMRS
'   f.Show
'End Sub
'
'Private Sub SetTimerNow()
'   Dim mConn As New ADODB.Connection
'   Dim rs As New ADODB.Recordset
'   mConn.Open (SQLCon)
'   rs.CursorLocation = adUseClient
'   rs.Open "Select GetDate()", mConn, adOpenStatic, adLockReadOnly, adCmdText
'   IsNow = rs.Fields(0).value
'End Sub
'
'Public Function Crypt(strOriginal As String, strPassword As String, encrypting As Boolean) As String
'
'Dim valPass As Long, valString As Long
'Dim X As Integer
'    For X = 1 To Len(strPassword)
'        valPass = valPass + Asc(Mid$(strPassword, X, 1)) + 13
'    Next X
'    For X = 1 To Len(strOriginal)
'        If encrypting = True Then
'            valString = Asc(Mid$(strOriginal, X, 1)) + valPass + (X ^ 2)
'            Do While valString > 255
'                valString = valString - 255
'            Loop
'        Else
'            valString = Asc(Mid$(strOriginal, X, 1)) - valPass - (X ^ 2)
'            Do While valString <= 0
'                valString = valString + 255
'            Loop
'        End If
'        Crypt = Crypt & Chr(valString)
'    Next X
'End Function
'
'Public Sub EnterPress(ByVal Key As Integer)
'      If Key = 13 Then SendKeys "{tab}"
'End Sub
'
'Function FormatTgl(ddate As Date) As String
'   FormatTgl = Format(ddate, "mm/dd/yyyy")
'End Function
'
'Function NullText(value As Variant) As Variant
'   NullText = IIf(IsNull(value), "", value)
'End Function
'
'Function NullTgl(Tgl As Variant) As Variant
'   If IsNull(Tgl) Or Tgl = "" Then
'      NullTgl = "Null"
'   Else
'      NullTgl = "'" & FormatTgl(CDate(Tgl)) & "'"
'   End If
'End Function
'
