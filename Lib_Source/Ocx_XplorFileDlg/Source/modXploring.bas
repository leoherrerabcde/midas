Attribute VB_Name = "modXploring"
'---------------------------------------------------------------------------------------
' Module    : modXploring
' Author    : Leo Herrera
' Date      : 20/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Public Declare Function SetErrorMode _
    Lib "kernel32" ( _
    ByVal wMode As Long) As Long

Public Declare Sub InitCommonControls Lib "Comctl32" ()

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
                ByVal hwnd As Long, ByVal nindex As Long, ByVal dwnewlong As Long) As Long

Public Declare Function InvalidateRect Lib "user32" _
                (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Enum enuTBType
    enuTB_FLAT = 1
    enuTB_STANDARD = 2
End Enum

Private Const GCL_HBRBACKGROUND = (-10)

Public Sub CambiarFondoToolbar(TB As Object, PNewBack As Long, pType As enuTBType)
Dim lTBWnd      As Long

    Select Case pType
        
        Case enuTB_FLAT
            DeleteObject SetClassLong(TB.hwnd, GCL_HBRBACKGROUND, PNewBack)
        
        Case enuTB_STANDARD
            lTBWnd = FindWindowEx(TB.hwnd, 0, "msvb_lib_toolbar", vbNullString)
            DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, PNewBack)
    End Select
End Sub



