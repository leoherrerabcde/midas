Attribute VB_Name = "modErrorListFunctions"
'---------------------------------------------------------------------------------------
' Module    : modErrorListFunctions
' Author    : Leo Herrera
' Date      : 14/01/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Const ERR_CODE__Frec_Error = 1
Public Const ERR_CODE__Amp_Error = 2
Public Const ERR_CODE__Pw_Error = 4
Public Const ERR_CODE__Neg_DToa = 8
Public Const ERR_CODE__Rel_Toa_Error = 16
Public Const ERR_CODE__Abs_Toa_Error = 32
Public Const ERR_CODE__FileTimeDesync = 64
'

Function GetErrorStr(lFilter As Long) As String

    Select Case lFilter
        Case Is = 1
            GetErrorStr = "Frecuencia Fuera de Rango"
        Case Is = 2
            GetErrorStr = "Amplitud Fuera de Rango"
        Case Is = 4
            GetErrorStr = "Ancho de Pulso Fuera de Rango"
        Case Is = 8
            GetErrorStr = "DToa Negatigo"
        Case Is = 16
            GetErrorStr = "Toa Relativo Decreciente"
        Case Is = 32
            GetErrorStr = "Toa Absoluto Decreciente"
        Case Is = 64
            GetErrorStr = "Desincronismo Tpo Archivo"
    End Select
    
End Function

Function GetErrorByFilter(lCode As Long, lFilter As Long) As String

    If (lCode And lFilter) = 0 Then
        GetErrorByFilter = ""
    Else
        GetErrorByFilter = GetErrorStr(lFilter)
    End If
    
End Function

Function GetErrorCode(lCode As Long) As String

    GetErrorCode = ""
    
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 1)
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 2)
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 4)
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 8)
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 16)
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 32)
    Concatenate_Error_Code GetErrorCode, GetErrorByFilter(lCode, 64)
    
End Function

Sub Concatenate_Error_Code(ByRef lsErrorDst As String, lsErrorSrc As String)

    If lsErrorSrc <> "" Then
        If lsErrorDst <> "" Then
            lsErrorDst = lsErrorDst & " + " & lsErrorSrc
        Else
            lsErrorDst = lsErrorSrc
        End If
    End If
    
End Sub
