Attribute VB_Name = "modHook"
'------------------------------------------------------------------
'M�dulo para subclasificaci�n (subclassing)             (26/Jun/98)
'
'
'�Guillermo 'guille' Som, 1998
'------------------------------------------------------------------
Option Explicit

'Para almacenar el form de llamada y el hWnd del form
Private elForm As Form
Private elhWnd As Long

Public PrevWndProc As Long
Public Const GWL_WNDPROC As Long = (-4&)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, _
    ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long


Public Function WndProc(ByVal hWnd As Long, ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WndProc = CallWindowProc(PrevWndProc, hWnd, uMSG, wParam, lParam)
    'Los mensajes de Windows llegar�n aqu�
    'Lo que hay que hacer es "capturar" los que se necesiten,
    'en este caso se devuelven los mensajes al form, usando para
    'ello un procedimiento p�blico llamado miMSG con los
    'siguientes par�metros:
    'ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long
    'la copia del form se har� al crear el Hook, es importante que
    's�lo se subclasifiquen ventanas cuando no halla ninguna activa
    '(de esto se encarga HookForm y unHookForm)
    '
    'Nos aseguramos que el form a�n est� disponible
'    If Not elForm Is Nothing Then
'        elForm.miMSG uMSG, wParam, lParam
'    End If

    frmWindowList.miMSG uMSG, wParam, lParam
    
    
End Function

Public Sub HookForm(ByVal unForm As Form)
    'unForm ser� el form de llamada,
    'para llamar a este procedimiento: HookForm Me
    '
    'Si a�n exist�a una subclasificaci�n
    If Not elForm Is Nothing Then
        unHookForm
    End If
    Set elForm = unForm
    elhWnd = unForm.hWnd
    PrevWndProc = SetWindowLong(elhWnd, GWL_WNDPROC, AddressOf WndProc)
    'Es importante recordar que se debe llamar a unHookForm antes
    'de cerrar el form... sobre todo si se usa en el IDE
End Sub


Public Sub HookApp(ByVal hWnd As Long)
    'unForm ser� el form de llamada,
    'para llamar a este procedimiento: HookForm Me
    '
    'Si a�n exist�a una subclasificaci�n
'    If Not elForm Is Nothing Then
'        unHookForm
'    End If
'    Set elForm = unForm
'    elhWnd = unForm.hWnd

    If elhWnd Then
        unHookApp
    End If
    
    elhWnd = hWnd
    PrevWndProc = SetWindowLong(elhWnd, GWL_WNDPROC, AddressOf WndProc)
    'Es importante recordar que se debe llamar a unHookForm antes
    'de cerrar el form... sobre todo si se usa en el IDE
End Sub

Public Sub unHookForm()
    Dim Ret As Long
    'Para llamar a este procedimiento: unHookForm
    '
    'Siempre se debe llamar primero a HookForm y despu�s se llama
    'a este otro para dejar de interceptar los mensajes de Windows
    'Si haces pruebas en el IDE, no te olvides de llamar a este
    'procedimiento, cerrando la aplicaci�n con el bot�n "Stop"
    'no se llamar� a este procedimiento.
    '
    'Si el valor de elhWnd es cero es que no se ha usado
    If elhWnd <> 0 Then
        Ret = SetWindowLong(elhWnd, GWL_WNDPROC, PrevWndProc)
    End If
    'Quitamos la referencia al form
    Set elForm = Nothing
    'Asignamos el valor cero a elhWnd
    elhWnd = 0
End Sub

Public Sub unHookApp()
    
    Dim Ret As Long
    'Para llamar a este procedimiento: unHookForm
    '
    'Siempre se debe llamar primero a HookForm y despu�s se llama
    'a este otro para dejar de interceptar los mensajes de Windows
    'Si haces pruebas en el IDE, no te olvides de llamar a este
    'procedimiento, cerrando la aplicaci�n con el bot�n "Stop"
    'no se llamar� a este procedimiento.
    '
    'Si el valor de elhWnd es cero es que no se ha usado
    If elhWnd <> 0 Then
        Ret = SetWindowLong(elhWnd, GWL_WNDPROC, PrevWndProc)
    End If
    'Quitamos la referencia al form
'    Set elForm = Nothing
    'Asignamos el valor cero a elhWnd
    elhWnd = 0
End Sub


