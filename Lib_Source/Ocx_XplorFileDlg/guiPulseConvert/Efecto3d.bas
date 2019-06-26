Attribute VB_Name = "Efecto3D_BAS"
Attribute VB_Description = "Módulo para Efectos 3D"
Option Explicit
'--------------------------------------------------
'Efecto 3D (nueva versión)              ( 5/Nov/94)
' Usando Container en lugar de Parent   ( 3/Sep/96)
'--------------------------------------------------
Global Const E3D_INSET = 1
Global Const E3D_RAISED = 2

Declare Function SetWindowWord Lib "User32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Sub Efecto3DN(QueEfecto As Integer, QueContenedor As Control, Optional QueControl)
    '------------------------------------------------------
    'Explicación de los parámetros a pasar:
    ' QueEfecto     Puede tener los valores: E3D_INSET o E3D_RAISED (tipo botón)
    ' ( 9/Nov/94)   Si QueEfecto > 10 DrawWidth=2...
    ' QueContenedor Si el contenedor es una Forma, se pondrá el mismo control,
    '               sino se pone el control contenedor de QueControl
    ' QueControl    Control al que se le hará el efecto 3D
    ' (10/Nov/95)   QueControl es opcional, usandose Quecontenedor
    '------------------------------------------------------
    
    Dim X As Long, Y As Long
    Dim CurrentX As Integer, CurrentY As Integer
    Dim Color_Gris As Long, Color_Blanco As Long
    Dim Ltmp As Long
    Dim Bevel As Integer
    
    If IsMissing(QueControl) Then
        Set QueControl = QueContenedor
    End If
    Color_Gris = RGB(92, 92, 92)
    Color_Blanco = RGB(255, 255, 255)
    
    'Ancho de la línea
    Bevel = 1
    Do While QueEfecto > 10
        QueEfecto = QueEfecto - 10
        Bevel = Bevel + 1
    Loop
        
    If QueEfecto = E3D_RAISED Then      'Estilo Command
        Ltmp = Color_Gris
        Color_Gris = Color_Blanco
        Color_Blanco = Ltmp
    End If

    X = Screen.TwipsPerPixelX
    Y = Screen.TwipsPerPixelY
    
    CurrentX = QueControl.Left - X
    CurrentY = QueControl.Top + QueControl.Height
    'Si se dibuja un Frame...                   (13/Nov/94)
    If TypeOf QueControl Is Frame Then
        Y = Y - 120
    End If
    
    If QueContenedor Is QueControl Then
        With QueControl
            .Container.DrawWidth = Bevel
            .Container.Line (CurrentX, CurrentY)-(CurrentX, CurrentY), Color_Gris
            .Container.Line -Step(0, -(.Height + Y)), Color_Gris
            .Container.Line -Step(.Width + X, 0), Color_Gris
            .Container.Line -Step(0, .Height + Y), Color_Blanco
            .Container.Line -Step(-(.Width + X), 0), Color_Blanco
        End With
    Else
        QueContenedor.DrawWidth = Bevel
        QueContenedor.Line (CurrentX, CurrentY)-(CurrentX, CurrentY), Color_Gris
        QueContenedor.Line -Step(0, -(QueControl.Height + Y)), Color_Gris
        QueContenedor.Line -Step(QueControl.Width + X, 0), Color_Gris
        QueContenedor.Line -Step(0, QueControl.Height + Y), Color_Blanco
        QueContenedor.Line -Step(-(QueControl.Width + X), 0), Color_Blanco
    End If

End Sub

