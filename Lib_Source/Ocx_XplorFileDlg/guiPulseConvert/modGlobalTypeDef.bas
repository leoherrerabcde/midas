Attribute VB_Name = "modGlobalTypeDef"
'---------------------------------------------------------------------------------------
' Module    : modGlobalTypeDef
' Author    : Leo Herrera
' Date      : 08/04/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit




Public Type LogEventAction
    Mdi_Control_Index           As Integer
    Form_Index                  As Integer
    Form_Control_Index          As Integer
    Control_Event_Index         As Integer
    Arguments()                    As String
End Type

