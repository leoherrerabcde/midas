Attribute VB_Name = "modResizeControls"
'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : Leo Herrera
' Date      : 29/04/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Const CTE_WIDTH_MARGEN = 60

Public Enum refoOperationsConstants        ' Resize Form Operation Constants
    refoLeft = 1
    refoWidth = 2
    refoRight = 2
    refoHorizontal = 3
    refoTop = 4
    refoHeight = 8
    refoBottom = 8
    refoVertical = 12
    refoToBorder = 16
    refoAlignToCenter = 32
    refoAlignToTop = 64
    refoAlignToBottom = 128
End Enum


Public Sub Move_Control(LV_Ctrl As Control, _
    Pos_X As Long, _
    Pos_Y As Long)

    LV_Ctrl.Top = Pos_Y
    LV_Ctrl.Left = Pos_X
    
End Sub

Public Sub Resize_Control_UpTo(ByRef LV_Ctrl As Control, _
    ByRef LV_Ctrl_Upto As Control, _
    ByVal lv_Op As refoOperationsConstants, _
    Optional ByVal lv_Margen As Long = 0)

    If lv_Op And refoToBorder Then
        If (lv_Op And refoHeight) Or (lv_Op And refoVertical) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Height = LV_Ctrl_Upto.Top - LV_Ctrl.Top
            Else
                LV_Ctrl.Height = LV_Ctrl_Upto.Top - LV_Ctrl.Top - lv_Margen
            End If
        End If
        If (lv_Op And refoWidth) Or (lv_Op And refoHorizontal) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Width = LV_Ctrl_Upto.Left - LV_Ctrl.Left
            Else
                LV_Ctrl.Width = LV_Ctrl_Upto.Left - LV_Ctrl.Left - lv_Margen
            End If
        End If
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure                   : Move_Control_Inside_Form
' Author                      : Leo Herrera
' Date                        : 31/05/2011
' Purpose                     : Mueve el control a la posiciòn indicada dentro del Formulario
' Input Arguments             :
' Output Arguments            :
' Used External Variables     :
' Modified External Variables :
' Modification Over the Inputs:
'---------------------------------------------------------------------------------------
'
Public Function Move_Control_Inside_Form(LV_Ctrl As Control, _
                    LV_Container As Form, _
                    lv_Op As refoOperationsConstants, _
                    Optional lv_Margen As Long = 0) As Boolean

Dim lvPos           As Long

    Move_Control_Inside_Form = True
    If lv_Op And refoLeft Then
        LV_Ctrl.Left = lv_Margen
    End If
    
    If lv_Op And refoRight Then
        lvPos = LV_Container.Width - LV_Ctrl.Width - lv_Margen
        If lvPos > 0 Then
            LV_Ctrl.Left = lvPos
        Else
            LV_Ctrl.Left = 0
            Move_Control_Inside_Form = False
        End If
    End If
    
    If lv_Op And refoTop Then
        LV_Ctrl.Top = lv_Margen
    End If
    
    If lv_Op And refoBottom Then
        lvPos = LV_Container.Height - LV_Ctrl.Height - lv_Margen
        If lvPos > 0 Then
            LV_Ctrl.Top = lvPos
        Else
            LV_Ctrl.Top = lvPos
        End If
    End If
    
End Function

Public Function Move_Control_NextTo(LV_Ctrl As Control, _
                    LV_Ctrl_Ref As Control, _
                    lv_Op As refoOperationsConstants, _
                    Optional lv_Margen As Long = 0) As Boolean

Dim lv_Pos           As Long

    Move_Control_NextTo = False
    If lv_Op And refoToBorder Then
        Move_Control_NextTo = True
        If lv_Op And refoLeft Then
            lv_Pos = LV_Ctrl_Ref.Left - LV_Ctrl.Width - lv_Margen
            If lv_Pos > 0 Then
                LV_Ctrl.Left = lv_Pos
            Else
                LV_Ctrl.Left = 0
                Move_Control_NextTo = False
            End If
        End If
        If lv_Op And refoTop Then
            lv_Pos = LV_Ctrl_Ref.Top - LV_Ctrl.Height - lv_Margen
            If lv_Pos > 0 Then
                LV_Ctrl.Top = lv_Pos
            Else
                LV_Ctrl.Top = 0
                Move_Control_NextTo = False
            End If
        End If
        If lv_Op And refoBottom Then
            lv_Pos = LV_Ctrl_Ref.Top + LV_Ctrl_Ref.Height + lv_Margen
            LV_Ctrl.Top = lv_Pos
        End If
        If lv_Op And refoRight Then
            lv_Pos = LV_Ctrl_Ref.Left + LV_Ctrl_Ref.Width + lv_Margen
            LV_Ctrl.Left = lv_Pos
        End If
    End If

End Function

Public Sub Resize_Control(LV_Ctrl As Control, _
                    LV_Container As Form, _
                    lv_Op As refoOperationsConstants, _
                    Optional lv_Margen As Long = 0)

Dim a As ContainedControls

    
    If lv_Op And refoToBorder Then
        If (lv_Op And refoLeft) Or (lv_Op And refoHorizontal) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Left = 0
            Else
                LV_Ctrl.Left = lv_Margen
            End If
        End If
        If (lv_Op And refoTop) Or (lv_Op And refoVertical) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Top = 0
            Else
                LV_Ctrl.Top = lv_Margen
            End If
        End If
        If (lv_Op And refoHeight) Or (lv_Op And refoVertical) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Height = LV_Container.Height - LV_Ctrl.Top
            Else
                LV_Ctrl.Height = LV_Container.Height - LV_Ctrl.Top - lv_Margen
            End If
        End If
        If (lv_Op And refoWidth) Or (lv_Op And refoHorizontal) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Width = LV_Container.Width - LV_Ctrl.Left
            Else
                LV_Ctrl.Width = LV_Container.Width - LV_Ctrl.Left - lv_Margen
            End If
        End If
    End If

End Sub



Public Sub Resize_Form(ByRef LV_Ctrl As Form, _
                    ByRef LV_Container As Form, _
                    ByVal lv_Op As refoOperationsConstants, _
                    Optional ByVal lv_Margen As Long = 0)

    If lv_Op And refoToBorder Then
        If (lv_Op And refoLeft) Or (lv_Op And refoHorizontal) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Left = 0
            Else
                LV_Ctrl.Left = lv_Margen
            End If
        End If
        If (lv_Op And refoTop) Or (lv_Op And refoVertical) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Top = 0
            Else
                LV_Ctrl.Top = lv_Margen
            End If
        End If
        If (lv_Op And refoHeight) Or (lv_Op And refoVertical) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Height = LV_Container.ScaleHeight - LV_Ctrl.Top
            Else
                LV_Ctrl.Height = LV_Container.ScaleHeight - LV_Ctrl.Top - lv_Margen
            End If
        End If
        If (lv_Op And refoWidth) Or (lv_Op And refoHorizontal) Then
            If lv_Margen = 0 Then
                LV_Ctrl.Width = LV_Container.ScaleWidth - LV_Ctrl.Left
            Else
                LV_Ctrl.Width = LV_Container.ScaleWidth - LV_Ctrl.Left - lv_Margen
            End If
        End If
    End If
End Sub



