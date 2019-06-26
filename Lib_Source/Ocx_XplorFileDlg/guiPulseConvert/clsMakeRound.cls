VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMakeRound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMakeRound
' Author    : Leo Herrera
' Date      : 30/11/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private m_Round_Value           As Long

Public Sub SetRoundValue(NewValue As Long)
    
    m_Round_Value = NewValue
    
End Sub

Public Sub MakeRoundControl(LV_Control As Control, Optional lvRound As Long)

    With LV_Control
        If lvRound Then
            MakeRoundRect .hWnd, _
                          .Width \ Screen.TwipsPerPixelX, _
                        .Height \ Screen.TwipsPerPixelY, _
                         lvRound
        Else
            MakeRoundRect .hWnd, _
                          .Width \ Screen.TwipsPerPixelX, _
                        .Height \ Screen.TwipsPerPixelY, _
                         m_Round_Value
        End If
    End With
    
End Sub

Public Sub MakeRoundForm(LV_Form As Form, Optional lvRound As Long)

    With LV_Form
        If lvRound Then
            MakeRoundRect .hWnd, _
                          .Width \ Screen.TwipsPerPixelX, _
                        .Height \ Screen.TwipsPerPixelY, _
                         lvRound
        Else
            MakeRoundRect .hWnd, _
                          .Width \ Screen.TwipsPerPixelX, _
                        .Height \ Screen.TwipsPerPixelY, _
                         m_Round_Value
        End If
    End With
    
End Sub

Private Sub Class_Initialize()

    m_Round_Value = 20
    
End Sub
