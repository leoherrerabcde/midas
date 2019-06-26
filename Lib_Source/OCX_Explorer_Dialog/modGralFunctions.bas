Attribute VB_Name = "modGralFunctions"
'---------------------------------------------------------------------------------------
' Module    : Module2
' Author    : Leo Herrera
' Date      : 25/11/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

'Public Const GC_CONFIGURATION_SECTION = "Configuration"

Public GV_App_Title         As String

Function FindPath(ByVal lsPath As String, lsNew As String) As String

    Do
        FindPath = lsPath & "\" & lsNew
        If Is_Folder(FindPath) = True Then
            Exit Do
        ElseIf IsPathDrive(FindPath) = True Then
            FindPath = ""
            Exit Do
        End If
        lsPath = Retroceder_Path(lsPath)
    Loop
    
End Function

Function IsNothing(lvObj As Object) As Boolean

   On Error GoTo IsNothing_Error

    If lvObj Is Nothing Then
        IsNothing = True
    Else
        IsNothing = False
    End If

   On Error GoTo 0
   Exit Function

IsNothing_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsNothing of Módulo modGralFunctions"
    IsNothing = True
    
End Function

Function Is_Folder(lvPath As String) As Boolean

Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")
    Is_Folder = fso.folderexists(lvPath)
    
End Function

Function Create_Folder(lsFolder As String) As Boolean

    If lsFolder = "" Then
        Exit Function
    End If
    If Is_Folder(lsFolder) = True Then
        Create_Folder = False
    Else
        MkDir lsFolder
        Create_Folder = True
    End If
    
End Function

Function Get_Sub_Folder(ls_Folder() As String, lv_Root As String) As Integer

Dim LV_Path         As String
Dim i               As Integer
'Dim k               As Integer
Dim LV_Count        As Integer
Dim fso

    ReDim ls_Folder(0)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ls_Folder(0) = lv_Root
    'k = 0
    LV_Count = 1
    
    Do
        LV_Path = Dir(ls_Folder(i) & "\", vbDirectory)
        Do
            If LV_Path <> "." And LV_Path <> ".." Then
                If LV_Path = "" Then
                    Exit Do
                ElseIf fso.folderexists(ls_Folder(i) & "\" & LV_Path) = True Then
                    ReDim Preserve ls_Folder(LV_Count)
                    ls_Folder(LV_Count) = ls_Folder(i) & "\" & LV_Path
                    LV_Count = LV_Count + 1
                End If
            End If
            LV_Path = Dir()
        Loop Until LV_Path = ""
        i = i + 1
    Loop Until i >= LV_Count
    Get_Sub_Folder = LV_Count
    
End Function

Function File_Exist(lvFileName As String) As Boolean

Dim lvAttr As VbFileAttribute

    On Error GoTo File_Exist_Error

    lvAttr = GetAttr(lvFileName)

    File_Exist = True
    If lvAttr = vbArchive Then
        File_Exist = True
    End If
    
    On Error GoTo 0
    Exit Function

File_Exist_Error:

    File_Exist = False
    
End Function

Function IsPathDrive(lsPath As String) As Boolean

    If Len(lsPath) <= 3 Then
        IsPathDrive = True
    Else
        IsPathDrive = False
    End If

End Function

Function Retroceder_Path(ByVal lsPath As String) As String

Dim i           As Integer
Dim lsTemp()    As String
Dim l           As Long

    On Error GoTo ErrRootFolder
    
    lsPath = Trim$(lsPath)
    lsTemp = Split(lsPath, "\")
    
    l = UBound(lsTemp)
    If lsTemp(l) = "" Then
        l = l - 1
    End If
    
    ReDim Preserve lsTemp(l - 1)
        
    Retroceder_Path = Join(lsTemp, "\")

    Exit Function
    
ErrRootFolder:
    Retroceder_Path = ""
    On Error GoTo 0

End Function

Function Obtener_Configuracion(lsSection As String, lsKey As String, lsDefault As String) As String

    Obtener_Configuracion = GetSetting(GV_App_Title, lsSection, lsKey, lsDefault)
    SaveSetting GV_App_Title, lsSection, lsKey, Obtener_Configuracion
    
End Function

Sub Guardar_Configuracion(lsSection As String, lsKey As String, lsDefault As String)

    SaveSetting GV_App_Title, lsSection, lsKey, lsDefault

End Sub

'Sub SetSettingLongParameter(lsKey As String, lbValue As Long)
'
'    SaveSetting App.Title, GC_CONFIGURATION_SECTION, lsKey, lbValue
'
'End Sub
'
'Function GetSettingLongParameter(lsKey As String, lvDefault As Long) As Long
'
'    GetSettingLongParameter = GetSetting(App.Title, _
'                                    GC_CONFIGURATION_SECTION, _
'                                    lsKey, lvDefault)
'    SaveSetting App.Title, GC_CONFIGURATION_SECTION, lsKey, GetSettingLongParameter
'
'End Function
'
'
