VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockFormToPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDockFormToPict
' Author    : Leo Herrera
' Date      : 03/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private WithEvents m_FormDock           As Form
Private m_PictContainer       As PictureBox
Attribute m_PictContainer.VB_VarHelpID = -1

Public Sub PickForm(LV_Pict As PictureBox, LV_Form As Form)

    Set m_FormDock = LV_Form
    Set m_PictContainer = LV_Pict
    
    dockFormAndRound m_PictContainer, m_FormDock
    
End Sub

Sub Resize_Pict()

    m_FormDock.Height = m_PictContainer.Height
    m_FormDock.Width = m_PictContainer.Width
    
End Sub

Sub Close_Form()

    Unload m_FormDock
End Sub

Private Sub Class_Initialize()

    Set m_FormDock = Nothing
    
End Sub

Private Sub Class_Terminate()

'    If m_FormDock Is Not Nothing Then
'        Unload m_FormDock
'    End If
    
End Sub
