Attribute VB_Name = "modArchivosIni"
'---------------------------------------------------------------------------------------
' Module    : modArchivosIni
' Author    : lherrera
' Date      : 19/01/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Dim m_Left As Single
Dim m_Top As Single
Dim m_Width As Single
Dim m_Height As Single

Dim Path_Archivo_Ini As String

'Función api que recupera un valor-dato de un archivo Ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long


'Lee un dato _
-----------------------------
'Recibe la ruta del archivo, la clave a leer y _
 el valor por defecto en caso de que la Key no exista
Function Leer_Ini(Path_INI As String, lsSection As String, Key As String, Default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(lsSection, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        Leer_Ini = Left$(bufer, Len_Value)

End Function

'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave

Function Grabar_Ini(Path_INI As String, lsSection As String, Key As String, Valor As Variant) As String

Dim lvErr       As Long
Dim lvStr       As String

    'lvStr = Err.Description
    
    lvErr = WritePrivateProfileString(lsSection, _
                                         Key, _
                                         Valor, _
                                         Path_INI)
                                         
    'lvStr = Err.Description
    
End Function

