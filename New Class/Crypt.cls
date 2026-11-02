VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Crypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Crypt(strOriginal As String, strPassword As String, encrypting As Boolean) As String
                        
Dim valPass As Long, valString As Long
    
    For X = 1 To Len(strPassword)
        valPass = valPass + Asc(Mid$(strPassword, X, 1)) + 13
    Next X
    For X = 1 To Len(strOriginal)
        If encrypting = True Then
            valString = Asc(Mid$(strOriginal, X, 1)) + valPass + (X ^ 2)
            Do While valString > 255
                valString = valString - 255
            Loop
        Else
            valString = Asc(Mid$(strOriginal, X, 1)) - valPass - (X ^ 2)
            Do While valString <= 0
                valString = valString + 255
            Loop
        End If
        Crypt = Crypt & Chr(valString)
    Next X
End Function
