Attribute VB_Name = "ITTS_DLL"
Option Explicit

Declare Sub IndoTTS_Say Lib "ITTS_DLL.dll" Alias "_IndoTTS_Say@4" (ByVal FName As String)
Declare Sub IndoTTS_SayToFile Lib "ITTS_DLL.dll" Alias "_IndoTTS_SayToFile@8" (ByVal FName As String, ByVal OutputFile As String)
Declare Sub IndoTTS_Stop Lib "ITTS_DLL.dll" Alias "_IndoTTS_Stop@0" ()
Declare Sub IndoTTS_ProsodyON Lib "ITTS_DLL.dll" Alias "_IndoTTS_ProsodyON@0" ()
Declare Sub IndoTTS_ProsodyOFF Lib "ITTS_DLL.dll" Alias "_IndoTTS_ProsodyOFF@0" ()
Declare Sub IndoTTS_SpeakON Lib "ITTS_DLL.dll" Alias "_IndoTTS_SpeakON@0" ()
Declare Sub IndoTTS_SpeakOFF Lib "ITTS_DLL.dll" Alias "_IndoTTS_SpeakOFF@0" ()
Declare Sub IndoTTS_SetPitchRatio Lib "ITTS_DLL.dll" Alias "_IndoTTS_SetPitchRatio@4" (ByVal pr As Single)
Declare Function IndoTTS_isPlaying Lib "ITTS_DLL.dll" Alias "_IndoTTS_isPlaying@0" () As Boolean
