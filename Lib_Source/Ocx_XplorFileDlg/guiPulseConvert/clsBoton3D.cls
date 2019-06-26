VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoton3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsBoton3D
' Author    : lherrera
' Date      : 27/01/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private WithEvents m_Timer_Clear        As Timer
Attribute m_Timer_Clear.VB_VarHelpID = -1
'Private WithEvents m_Image              As Image
Private m_Boton                         As Control
Private m_bEfectoActivado               As Boolean
Private m_Container                     As Control
Private m_hWnd                          As Long

Private Function Is_The_Same(LV_C1 As Control, LV_C2 As Control) As Boolean

    Is_The_Same = False
    
    If LV_C1.Top <> LV_C2.Top Then
        Exit Function
    End If
    
    If LV_C1.Left <> LV_C2.Left Then
        Exit Function
    End If
    
    Is_The_Same = True
    
End Function

Public Sub Iniciar(lvControl As Control, lvContainer As Control, lvTimer As Timer)

    If m_Boton Is Nothing Then
    Else
        If Is_The_Same(m_Boton, lvControl) = False Then
            m_Container.Cls
            m_bEfectoActivado = False
            Set m_Boton = Nothing
            Set m_Container = Nothing
            m_hWnd = Get_Cursor_Pos
        Else
            'm_Timer_Clear.Enabled = False
            'm_Timer_Clear.Enabled = True
            Exit Sub
        End If
    End If
    m_hWnd = Get_Cursor_Pos
    Set m_Boton = lvControl
    Set m_Container = lvContainer
    If m_Timer_Clear Is Nothing Then
        Set m_Timer_Clear = lvTimer
    End If
    m_Timer_Clear.Interval = 200
    m_Timer_Clear.Enabled = True
    Efecto3DN E3D_RAISED, lvContainer, m_Boton
    m_bEfectoActivado = True

End Sub

'Private Sub m_Image_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    With m_Image
''        If PV_nBotonFileOp Then
''            If PV_nBotonFileOp <> Index + 1 Then
''                .Container.Cls      'Borrar el efecto anterior
''            Else
''                'Si estamos en el mismo botón, salir
''                Exit Sub
''            End If
''        End If
''        PV_nBotonFileOp = Index + 1
'        'Dibujar el efecto "botón"
'        If m_Timer_Clear Is Nothing Then
'
'        Else
'            m_Timer_Clear.Enabled = False
'            m_Timer_Clear.Enabled = True
'        End If
'        If m_bEfectoActivado = False Then
'            Efecto3DN E3D_RAISED, .Container, m_Image
'            m_bEfectoActivado = True
'        End If
'    End With
'
'End Sub

Private Sub m_Timer_Clear_Timer()

    If m_hWnd <> Get_Cursor_Pos Then
        m_Timer_Clear.Enabled = False
        m_Container.Cls
        m_bEfectoActivado = False
        Set m_Boton = Nothing
        Set m_Container = Nothing
    End If
    
End Sub
