VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetectClickOnDisableControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDetectClickOnDisableControl
' Author    : Leo Herrera
' Date      : 16/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private mMatch                      As Boolean
Private WithEvents mContainer       As Form
Attribute mContainer.VB_VarHelpID = -1
Private mControl                    As Control

Public Event OnDisableControlClicked()

Private Function VerifyPosition(X As Single, Y As Single) As Boolean

    VerifyPosition = False
    With mControl
        If .Left > X Then
            Exit Function
        End If
        If .Left + .Width < X Then
            Exit Function
        End If
        If .Top > Y Then
            Exit Function
        End If
        If .Top + .Height < Y Then
            Exit Function
        End If
    End With
    VerifyPosition = True
    
End Function

Public Sub SetControls(LV_Control As Control, LV_Container As Form)

    Set mContainer = LV_Container
    Set mControl = LV_Control
    
End Sub

Private Sub mContainer_Click()

    If mMatch = True Then
        RaiseEvent OnDisableControlClicked
    End If
    
End Sub

Private Sub mContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mControl.enable = False Then
        mMatch = VerifyPosition(X, Y)
    End If
    
End Sub

