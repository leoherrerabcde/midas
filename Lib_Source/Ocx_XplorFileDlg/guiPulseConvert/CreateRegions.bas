Attribute VB_Name = "CreateRegions"
' ===============================================================
' API declares for creating regions and setting a window's region
' ===============================================================

'' Point
'Type POINTAPI
'    x As Long
'    y As Long
'End Type

' Change region of a window:
Declare Function SetWindowRgn Lib "user32" _
    (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Precanned region creation functions:
Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, _
    ByVal Y3 As Long) As Long

' Polygon region creation functions:
Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, lpPolyCounts As Long, _
    ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

' Polygon type:
Public Const WINDING = 2
' Region combination:
Declare Function CombineRgn Lib "gdi32" _
    (ByVal hDestRgn As Long, _
    ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
    ByVal nCombineMode As Long) As Long
    
' Region combination types:
    Public Const RGN_AND = 1
    Public Const RGN_COPY = 5
    Public Const RGN_DIFF = 4
    Public Const RGN_MAX = RGN_COPY
    Public Const RGN_MIN = RGN_AND
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3
' Region combination return values:
    Public Const COMPLEXREGION = 3
    Public Const SIMPLEREGION = 2
    Public Const NULLREGION = 1

' GDI Clear up:
Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long




Sub MakeRoundRect( _
    ByVal lHwnd As Long, _
    ByVal lWidth As Long, ByVal lHeight As Long, _
    ByVal lRound As Long)

Dim hRgn As Long
    
    'hRgn = CreateEllipticRgn(0, 0, lWidth, lHeight)
    
    hRgn = CreateRoundRectRgn(0, 0, lWidth, lHeight, lRound, lRound)
    
    ' Change the region:
    SetWindowRgn lHwnd, hRgn, 1
    'DeleteObject hRgn
End Sub



