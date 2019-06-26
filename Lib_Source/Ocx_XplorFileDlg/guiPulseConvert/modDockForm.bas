Attribute VB_Name = "modDockForm"
Option Explicit

Public Enum DockFormPositionConstant
    PosLeft = 1
    PosRight
    PosOver
    PosUnder
End Enum

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'
'------------------------------------------------------------------------------
' APIS para incluir las ventanas en un PictureBox
'------------------------------------------------------------------------------
'
' Para hacer ventanas hijas
'Dim prevParent As Long
Public Declare Function SetParent Lib "user32" _
    (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'
' Para mostrar una ventana según el handle (hwnd)
' ShowWindow() Commands
Public Enum eShowWindow
    HIDE_eSW = 0&
    SHOWNORMAL_eSW = 1&
    NORMAL_eSW = 1&
    SHOWMINIMIZED_eSW = 2&
    SHOWMAXIMIZED_eSW = 3&
    MAXIMIZE_eSW = 3&
    SHOWNOACTIVATE_eSW = 4&
    SHOW_eSW = 5&
    MINIMIZE_eSW = 6&
    SHOWMINNOACTIVE_eSW = 7&
    SHOWNA_eSW = 8&
    RESTORE_eSW = 9&
    SHOWDEFAULT_eSW = 10&
    MAX_eSW = 10&
End Enum
'Private Declare Function ShowWindow Lib "user32" _
'    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As eShowWindow) As Long
'
' Para posicionar una ventana según su hWnd
Public Declare Function MoveWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'
' Para saber si una ventana es hija de otra
Public Declare Function IsChild Lib "user32" _
    (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
'
'' Para posicionar una ventana y asignar el ZOrder, etc.
'public Declare Function SetWindowPos Lib "user32" _
'    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
'    ByVal x As Long, ByVal y As Long, _
'    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Para cambiar el tamaño de una ventana y asignar los valores máximos y mínimos del tamaño
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECTAPI
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECTAPI
End Type
'public Declare Function SetWindowPlacement Lib "user32" _
    (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowPlacement Lib "user32" _
    (ByVal hWnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

' Mostrar el formulario indicado, dentro de picDock
Public Sub dockFormPict(ByVal formhWnd As Long, _
                     ByVal picDock As PictureBox, _
                     Optional ByVal ajustar As Boolean = True)
    ' Hacer el formulario indicado, un hijo del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Call SetParent(formhWnd, picDock.hWnd)
    posDockFormPict formhWnd, picDock, ajustar
    Call ShowWindow(formhWnd, NORMAL_eSW)
End Sub

Public Sub dockForm(ByVal formhWnd As Long, _
                     ByVal picDock As Form, _
                     Optional ByVal ajustar As Boolean = True)
    ' Hacer el formulario indicado, un hijo del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Call SetParent(formhWnd, picDock.hWnd)
    posDockForm formhWnd, picDock, ajustar
    Call ShowWindow(formhWnd, NORMAL_eSW)
End Sub

' Posicionar el formulario indicado dentro de picDock
' Posicionar el formulario indicado dentro de picDock
Private Sub posDockFormPict(ByVal formhWnd As Long, _
                        ByVal picDock As PictureBox, _
                        Optional ByVal ajustar As Boolean = True)
    ' Posicionar el formulario indicado en las coordenadas del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Dim nWidth As Long, nHeight As Long
    Dim wndPl As WINDOWPLACEMENT
    '
    If ajustar Then
        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
    Else
        ' el tamaño del formulario que se va a posicionar
        Call GetWindowPlacement(formhWnd, wndPl)
        With wndPl.rcNormalPosition
            nWidth = .Right - .Left
            nHeight = .Bottom - .Top
        End With
    End If
    Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)
End Sub

Public Sub posDockForm(ByVal formhWnd As Long, _
                        ByVal picDock As Form, _
                        Optional ByVal ajustar As Boolean = True)
    ' Posicionar el formulario indicado en las coordenadas del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Dim nWidth As Long, nHeight As Long
    Dim wndPl As WINDOWPLACEMENT
    '
    If ajustar Then
        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
    Else
        ' el tamaño del formulario que se va a posicionar
        Call GetWindowPlacement(formhWnd, wndPl)
        With wndPl.rcNormalPosition
            nWidth = .Right - .Left
            nHeight = .Bottom - .Top
        End With
    End If
    'Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)
    Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)
End Sub


Public Sub posDockFormBeSide(ByVal formhWnd As Long, _
                        ByVal frmRefhWnd As Long, _
                        ByVal Pos As DockFormPositionConstant)
    ' Posicionar el formulario indicado en las coordenadas del picDock
    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
    ' si Ajustar es False, se quedará con el tamaño actual.
    Dim nWidth              As Long
    Dim nHeight             As Long
    Dim wndPl               As WINDOWPLACEMENT
    Dim PosX                As Long
    Dim PosY                As Long
    Dim RefWidth            As Long
    Dim RefHeight           As Long
    
'    If ajustar Then
'        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
'        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
'    Else
        ' el tamaño del formulario que se va a posicionar
        Call GetWindowPlacement(formhWnd, wndPl)
        With wndPl.rcNormalPosition
            nWidth = .Right - .Left
            nHeight = .Bottom - .Top
        End With
'    End If
    Call GetWindowPlacement(frmRefhWnd, wndPl)
    With wndPl.rcNormalPosition
        RefWidth = .Right - .Left
        RefHeight = .Bottom - .Top
        Select Case Pos

            Case PosLeft
                PosX = .Left - RefWidth
                PosY = .Top
            
            Case PosRight
                PosX = .Right
                PosY = .Top

            Case PosOver
                PosX = .Left
                PosY = .Top - RefHeight

            Case PosUnder
                PosX = .Left
                PosY = .Bottom

        End Select
    End With

    'Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)
    Call MoveWindow(formhWnd, PosX, PosY, nWidth, nHeight, True)

End Sub

Sub SetDockFormPosition(lvFrm As Form, pickForm As Form, lv_Left As Single, lv_Top As Single)

    With lvFrm
        .Left = lv_Left
        .Top = lv_Top
        posDockForm lvFrm.hWnd, pickForm, False
    End With
    
End Sub



