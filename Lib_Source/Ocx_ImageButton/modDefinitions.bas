Attribute VB_Name = "modDefinitions"
'---------------------------------------------------------------------------------------
' Module    : modDefinitions
' Author    : Leo Herrera
' Date      : 10/11/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


'Public Const CONST_ORIENT_PORTRAIT = 0
'Public Const CONST_ORIENT_LANDSCAPE = 1

Public Enum OrientationConstants
    orientPortrait = 0
    orientLandscape = 1
End Enum

Public Enum ResizeObjectConstant
    PictureButtonsObject = 0
    ContainerObject
    ControlObject
End Enum

Public Type ResizeStatus
    PictureButtonsResizeCounter     As Integer
    CointainerResizeCounter         As Integer
    ControlResizeCounter            As Integer
    Count                           As Integer
End Type

