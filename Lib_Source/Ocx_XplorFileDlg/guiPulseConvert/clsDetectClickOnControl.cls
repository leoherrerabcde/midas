VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetectClickOnControl"
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
Private mControl                    As Control
Private WithEvents mTimer           As Timer

Public Event OnDisableControlClicked()

Public Sub SetControls(LV_Control As Control, LV_Timer As Timer)

    Set mControl = LV_Control
    Set mTimer = LV_Timer
    mTimer.Enabled = True
    
End Sub

Private Sub mTimer_Timer()

    If mControl.Enabled = False Then
        If mControl.hWnd = Get_Cursor_Pos Then
            RaiseEvent OnDisableControlClicked
        End If
    End If
    
End Sub
