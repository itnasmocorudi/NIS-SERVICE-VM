Attribute VB_Name = "Sound"
Option Explicit
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const wav_0 = "0.wav"
Const wav_1 = "1.wav"
Const wav_2 = "2.wav"
Const wav_3 = "3.wav"
Const wav_4 = "4.wav"
Const wav_5 = "5.wav"
Const wav_6 = "6.wav"
Const wav_7 = "7.wav"
Const wav_8 = "8.wav"
Const wav_9 = "9.wav"
Const wav_10 = "10.wav"
Const wav_11 = "11.wav"
Const wav_seratus = "100.wav"
Const wav_seribu = "1000.wav"
Const wav_puluh = "puluh.wav"
Const wav_belas = "belas.wav"
Const wav_ribu = "ribu.wav"
Const wav_juta = "juta.wav"
Const wav_ratus = "ratus.wav"

Function Puluh(X As Double) As String
On Error Resume Next
Dim Bil() As String
ReDim Bil(9)

Bil(0) = wav_0
Bil(1) = wav_1
Bil(2) = wav_2
Bil(3) = wav_3
Bil(4) = wav_4
Bil(5) = wav_5
Bil(6) = wav_6
Bil(7) = wav_7
Bil(8) = wav_8
Bil(9) = wav_9

If X < 10 Then
    Puluh = Bil(X)
ElseIf X = 10 Then
    Puluh = wav_10
ElseIf Trim(X) = 11 Then
    Puluh = wav_11
ElseIf Left(X, 1) = 1 And Right(X, 1) > 1 Then
    Puluh = Bil(Val(Right(X, 1))) & ";" & wav_belas
ElseIf X > 19 Then
    Puluh = Bil(Val(Left(X, 1))) & ";" & wav_puluh & ";" & IIf(Right(X, 1) = "0", "", Bil(Val(Right(X, 1))))
End If
End Function

Function Terbilang(X As Double) As String
On Error Resume Next
If X < 1 Then
        Terbilang = ""
ElseIf X < 100 Then
        Terbilang = Puluh(X)
ElseIf X < 1000 Then
        If Int(X / 100) = 1 Then
            Terbilang = wav_seratus & ";" & Terbilang(X Mod 100)
        Else
            Terbilang = Terbilang(Int(X / 100)) & ";" & wav_ratus & ";" & Terbilang(X Mod 100)
        End If
ElseIf X < 1000000 Then
        If Int(X / 1000) = 1 Then
            Terbilang = wav_seribu & ";" & Terbilang(X Mod 1000)
        Else
            Terbilang = Terbilang(Int(X / 1000)) & ";" & wav_ribu & ";" & Terbilang(X Mod 1000)
        End If
ElseIf X < 1000000000 Then
            Terbilang = Terbilang(Int(X / 1000000)) & ";" & wav_juta & ";" & Terbilang(X Mod 1000000)
End If
End Function
