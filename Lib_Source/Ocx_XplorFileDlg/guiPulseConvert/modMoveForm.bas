Attribute VB_Name = "modMoveForm"
'---------------------------------------------------------------------------------------
' Module    : modMoveForm
' Author    : lherrera
' Date      : 26/11/2010
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

'--------------------------------------------------------------------
'NOTAS:
'Listado a insertar en un módulo (.bas)
'si se quiere poner en un formulario (.frm)
'declarar la función como Private y quitar el Global de las constantes
'--------------------------------------------------------------------
'Constantes y declaración de función:
'
'Constantes para SendMessage
Global Const WM_LBUTTONUP = &H202
Global Const WM_SYSCOMMAND = &H112
Global Const SC_MOVE = &HF010
Global Const MOUSE_MOVE = &HF012

#If Win32 Then
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
#Else
    Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
#End If
'
'
'Este código se pondrá en el Control_MouseDown...
'



Public Sub Form_Move(lvFrm As Form, Pos_X As Long, Pos_Y As Long)

    lvFrm.Left = Pos_X
    lvFrm.Top = Pos_Y
    
End Sub

Public Function Form_Move_RefTo(ByRef lvFrm As Form, _
            ByRef lvFrmRef As Form, _
            ByVal lvOp As refoOperationsConstants, _
            Optional ByVal offSet_X As Long = 0, _
            Optional ByVal offSet_Y As Long = 0)
            
    If lvOp And refoToBorder Then
        If lvOp And refoLeft Then
            lvFrm.Left = lvFrmRef.Left - lvFrm.Width - offSet_X
        End If
        If lvOp And refoRight Then
            lvFrm.Left = lvFrmRef.Left + lvFrmRef.Width + offSet_X
        End If
        If lvOp And refoTop Then
            lvFrm.Top = lvFrmRef.Top - lvFrm.Height - offSet_Y
        End If
        If lvOp And refoBottom Then
            lvFrm.Top = lvFrmRef.Top + lvFrmRef.Height + offSet_Y
        End If
    End If
    
    If lvOp And refoAlignToTop Then
        If lvOp And refoHorizontal Then
            lvFrm.Left = lvFrmRef.Left + offSet_X
        End If
        If lvOp And refoVertical Then
            lvFrm.Top = lvFrmRef.Top + offSet_Y
        End If
    End If
    
    If lvOp And refoAlignToCenter Then
        If lvOp And refoHorizontal Then
            lvFrm.Left = lvFrmRef.Left + lvFrmRef.Width / 2 + offSet_X
        End If
        If lvOp And refoVertical Then
            lvFrm.Top = lvFrmRef.Top + lvFrmRef.Height / 2 + offSet_Y
        End If
    End If
    
    If lvOp And refoAlignToBottom Then
        If lvOp And refoHorizontal Then
            lvFrm.Left = lvFrmRef.Left + lvFrmRef.Width - lvFrm.Width - offSet_X
        End If
        If lvOp And refoVertical Then
            lvFrm.Top = lvFrmRef.Top + lvFrmRef.Height - lvFrm.Height - offSet_Y
        End If
    End If

End Function

Public Sub Form_Move_BesideTo(lvFrm As Form, lvFrmRef As Form, _
            offSet_X As Long, _
            offSet_Y As Long)
            
    lvFrm.Left = lvFrmRef.Left + lvFrmRef.Width + offSet_X
    lvFrm.Top = lvFrmRef.Top + offSet_Y
    
End Sub

