Attribute VB_Name = "ApiModule"
'on top api
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWysiwyg = 2
End Enum

Public Declare Function LockWindowUpdate Lib "user32" (ByVal HWNDLOCK As Long) As Long
Public isBookByMrs As Boolean

'drag api
Option Explicit
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long



'======================================================================================
'.:: ========== Its Child Or Not ======================================================================================
'.:: ========== ===== Begin ====
'======================================================================================
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Sub TOChild(Hwnd_Parent As Long, Hwnd_Child As Long)
    SetParent Hwnd_Child, Hwnd_Parent
End Sub

Sub NotChild(Hwnd_Child As Long)
    SetParent Hwnd_Child, 0&
End Sub
'======================================================================================
'.:: ========== ===== End =====  ======================================================================================
'.:: ========== Its Child Or Not ======================================================================================


Public Sub FormDrag(TheForm As Form)
    On Error Resume Next
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

Function RepRp(nominal) As String
    RepRp = (Replace(Trim(nominal), ",", "."))
End Function

Function RepRpInv(nominal) As String
    RepRpInv = (Replace(Trim(nominal), ".", ","))
End Function
