Attribute VB_Name = "modGeneralFunctions"
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Function Open_File(ByRef CT_CmnDlg As CommonDialog, _
'                    ByRef lPath As String, _
'                    lvFilter As String, _
'                    lTitle As String) As String
'
''Dim sDir        As String
''Dim lFlags      As Long
''Dim lPath       As String
'Dim sFile       As String
'
'    With CT_CmnDlg
'        .Filter = lvFilter  ' "*.csv"
'        .InitDir = lPath
'        .CancelError = False
'        .DialogTitle = lTitle
'        .ShowOpen
'        sFile = .FileName
'    End With
'    'sFile = BrowseForFile(lPath, "*.csv", "Archivo de COmpensación de Salida")
'    Open_File = sFile
'
'End Function

'Function Abrir_Directorio(lPath As String)
'
'Dim sDir        As String
'Dim lFlags      As Long
'Dim lPath       As String
'Dim sFile       As String
'
'    lFlags = BIF_RETURNONLYFSDIRS
'
'    lPath = "" 'GV_Actual_Project.Path_Files_Datos
'
'    sDir = BrowseForFolder(Me.hWnd, "Seleccionar Directorio", lPath, lFlags)
'
'    If Err = 0 Then
'        Me.txtProyecto(1).Text = sDir
'    Else
'        'MsgBox "Se ha cancelado la operación, el error devuelto es:" & vbCrLf & _
'               "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
'        Err = 0
'    End If
'
'End Function

Public Function CommonDialogForFolder(LV_CommDialog As CommonDialog, ByVal sPrompt As String, _
                Optional PathInicial As String = "", _
                Optional ByVal lvFilter As String = "*.*", _
                Optional ByVal lFlags As Long = BIF_RETURNONLYFSDIRS) _
                As String
                
    With LV_CommDialog
        .Flags = cdlOFNPathMustExist
        .Filter = lvFilter
        .InitDir = PathInicial
        .CancelError = True
        .DialogTitle = sPrompt
        On Error Resume Next
        .ShowOpen
        If Err <> cdlCancel Then
            CommonDialogForFolder = .FileName
        End If
        On Error GoTo 0
    End With
                
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

    Obtener_Configuracion = GetSetting(App.Title, lsSection, lsKey, lsDefault)
    SaveSetting App.Title, lsSection, lsKey, Obtener_Configuracion
    
End Function

Sub Guardar_Configuracion(lsSection As String, lsKey As String, lsDefault As String)

    SaveSetting App.Title, lsSection, lsKey, lsDefault

End Sub

Sub Load_Mdi_Params(ByRef MdiFrm As MDIForm)

    With MdiFrm
        .Left = Obtener_Configuracion("Settings", .Name & ".Left", .Left)
        .Top = Obtener_Configuracion("Settings", .Name & ".Top", .Top)
        .Width = Obtener_Configuracion("Settings", .Name & ".Width", .Width)
        .Height = Obtener_Configuracion("Settings", .Name & ".Height", .Height)
        .WindowState = Obtener_Configuracion("Settings", .Name & ".WindowState", .WindowState)
    End With
    
End Sub

Sub Save_Mdi_Params(ByRef MdiFrm As MDIForm)

    With MdiFrm
        Guardar_Configuracion "Settings", .Name & ".Left", .Left
        Guardar_Configuracion "Settings", .Name & ".Top", .Top
        Guardar_Configuracion "Settings", .Name & ".Width", .Width
        Guardar_Configuracion "Settings", .Name & ".Height", .Height
        Guardar_Configuracion "Settings", .Name & ".WindowState", .WindowState
    End With
    
End Sub

Sub Load_Form_Params(ByRef MdiFrm As Form)

    With MdiFrm
        .Left = Obtener_Configuracion("Settings", .Name & ".Left", .Left)
        .Top = Obtener_Configuracion("Settings", .Name & ".Top", .Top)
        .Width = Obtener_Configuracion("Settings", .Name & ".Width", .Width)
        .Height = Obtener_Configuracion("Settings", .Name & ".Height", .Height)
    End With
    
End Sub

Sub Save_Control_Text(lvText As Control)

    SaveSetting App.Title, lvText.Name, ".Text", lvText.Text

End Sub

Sub Save_Form_Params(ByRef MdiFrm As Form)

    With MdiFrm
        Guardar_Configuracion "Settings", .Name & ".Left", .Left
        Guardar_Configuracion "Settings", .Name & ".Top", .Top
        Guardar_Configuracion "Settings", .Name & ".Width", .Width
        Guardar_Configuracion "Settings", .Name & ".Height", .Height
    End With
    
End Sub

Function GetSettingCheckBox(LV_ChkBx As CheckBox) As Integer

    With LV_ChkBx
        GetSettingCheckBox = GetSetting(App.Title, .Container.Name, .Name & ".Value", .Value)
        .Value = GetSettingCheckBox
    End With
    
End Function

Sub SetSettingBooleanParameter(lsKey As String, lbValue As Boolean)

    SaveBooleanSetting GC_CONFIGURATION_SECTION, lsKey, lbValue

End Sub

Function GetSettingBooleanParameter(lsKey As String, lvDefault As Boolean) As Boolean

    GetSettingBooleanParameter = GetBooleanSetting( _
                                    GC_CONFIGURATION_SECTION, _
                                    lsKey, lvDefault)
    SaveBooleanSetting GC_CONFIGURATION_SECTION, lsKey, GetSettingBooleanParameter
    
End Function

Sub SaveSettingCheckBox(LV_ChkBx As CheckBox)

    With LV_ChkBx
        SaveSetting App.Title, .Container.Name, .Name & ".Value", .Value
    End With
    
End Sub

'Sub Get_Sub_Folder(ls_Folder() As String, lv_Root As String)
'
'Dim lv_Path         As String
'Dim i               As Integer
''Dim k               As Integer
'Dim LV_Count        As Integer
'Dim fso
'
'    ReDim ls_Folder(0)
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    ls_Folder(0) = lv_Root
'    'k = 0
'    LV_Count = 1
'
'    Do
'        lv_Path = Dir(ls_Folder(i) & "\", vbDirectory)
'        Do
'            If lv_Path <> "." And lv_Path <> ".." Then
'                If lv_Path = "" Then
'                    Exit Do
'                ElseIf fso.folderexists(ls_Folder(i) & "\" & lv_Path) = True Then
'                    ReDim Preserve ls_Folder(LV_Count)
'                    ls_Folder(LV_Count) = ls_Folder(i) & "\" & lv_Path
'                    LV_Count = LV_Count + 1
'                End If
'            End If
'            lv_Path = Dir()
'        Loop Until lv_Path = ""
'        i = i + 1
'    Loop Until i >= LV_Count
'
'End Sub

Function IfBooleanText(lvFlag As Boolean, lvTextTrue As String, lvTextFalse As String)

    If lvFlag = False Then
        IfBooleanText = lvTextFalse
    Else
        IfBooleanText = lvTextTrue
    End If
    
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

Function Get_Cursor_Pos() As Long

Dim tP As POINTAPI
Dim lHwnd As Long
Dim iPos As Long
Dim sBuf As String
Dim lR As Long
   
   GetCursorPos tP
   lHwnd = WindowFromPoint(tP.X, tP.Y)
   Get_Cursor_Pos = lHwnd
'   txthWnd = Hex$(lhWnd)
'   txtClassName = ""
'   txtStyles = ""
'   txtParenthWnd = ""
'   If lhWnd <> 0 Then
'      txtParenthWnd = GetWindow(lhWnd, GW_OWNER)
'      sBuf = String$(255, 0)
'      lR = GetClassName(lhWnd, sBuf, 255)
'      If lR > 0 Then
'         iPos = InStr(sBuf, vbNullChar)
'         If (iPos > 0) Then
'            txtClassName = Left$(sBuf, iPos - 1)
'         Else
'            txtClassName = sBuf
'         End If
'      End If
'      lR = GetWindowLong(lhWnd, GWL_STYLE)
'      txtStyles = Hex$(lR) & ShowStyles(lR)
'      lR = GetWindowLong(lhWnd, GWL_EXSTYLE)
'      txtExStyles = Hex$(lR) & ShowExStyles(lR)
'   End If

End Function

Sub Set_Control_Size(lvCtl As Control, lWidth, lHeight)

    With lvCtl
        .Width = lWidth - 2 * .Left
        .Height = lHeight - 2 * .Top
    End With
    
End Sub

Sub Set_MousePointer(lvState As MousePointerConstants)

    Screen.MousePointer = lvState
    
End Sub

Sub Verifi_Key_Action(ByVal KeyAscii As Integer, _
                        ByRef lvForm As Form, _
                        lvKeyMatch As KeyCodeConstants _
                        )

    
    If KeyAscii = lvKeyMatch Then
'        Select Case lvAction
'
'            Case CloseForm
'                Unload lvForm
'
'        End Select
    End If

End Sub

Function UBoundArray(lvArray() As Variant) As Integer

    On Error GoTo UBoundArray_Error
    
    UBoundArray = UBound(lvArray)
    
    Exit Function
    
UBoundArray_Error:

    UBoundArray = -1
    
End Function

Function UBoundStringArray(lvArray() As String) As Integer

    On Error GoTo UBoundStringArray_Error
    
    UBoundStringArray = UBound(lvArray)
    
    Exit Function
    
UBoundStringArray_Error:

    UBoundStringArray = -1
    
End Function


Function Obtener_Ultimo_Path_Form(lsSection As String, lsDefault As String) As String

    Obtener_Ultimo_Path_Form = Obtener_Configuracion("Configuracion " & lsSection, "Ultimo Path", lsDefault)
    
End Function

Function Guardar_Ultimo_Path_Form(lsSection As String, lsPath As String) As String

    'Obtener_Ultimo_Path = Obtener_Configuracion("Configuracion Aplicacion", "Ultimo Path", lsDefault)
    Guardar_Configuracion "Configuracion " & lsSection, "Ultimo Path", lsPath

End Function

Function GetBooleanSetting(Section As String, Key As String, _
        Default As Boolean) As Boolean

Dim lvState         As Integer

    GetBooleanSetting = True
    lvState = Default
    lvState = GetSetting(App.Title, Section, Key, lvState)
    If lvState = 0 Then
        GetBooleanSetting = False
    End If
    
End Function

Sub SaveBooleanSetting(Section As String, Key As String, _
        lvValue As Boolean)

Dim lvState         As Integer

    lvState = 0
    If lvValue = True Then
        lvState = 1
    End If
    SaveSetting App.Title, Section, Key, lvState

End Sub

Function GetMissionDateFromMissionPath(ByVal LV_MissionPath As String) As String

Dim lsCampo()       As String
Dim lsLastChar      As String
    
    lsCampo = Split(LV_MissionPath, "\")
    If UBound(lsCampo) >= 1 Then
        GetMissionDateFromMissionPath = lsCampo(UBound(lsCampo))
        GetMissionDateFromMissionPath = Right$(GetMissionDateFromMissionPath, 19)
    End If
    
End Function

Function GetMissionNameFromMissionPath(ByVal LV_MissionPath As String) As String

Dim lsCampo()       As String
Dim lsLastChar      As String
    
    lsCampo = Split(LV_MissionPath, "\")
    If UBound(lsCampo) >= 1 Then
        GetMissionNameFromMissionPath = lsCampo(UBound(lsCampo))
'        If Len(GetMissionNameFromMissionPath) < 19 Then
'            GetMissionNameFromMissionPath = ""
'            Exit Function
'        End If
'        GetMissionNameFromMissionPath = Left$(GetMissionNameFromMissionPath, Len(GetMissionNameFromMissionPath) - 19)
'        lsLastChar = Right$(GetMissionNameFromMissionPath, 1)
'        If lsLastChar = "_" Or lsLastChar = "-" Then
'            GetMissionNameFromMissionPath = Left$(GetMissionNameFromMissionPath, Len(GetMissionNameFromMissionPath) - 1)
'        End If
    End If
    
End Function

Function GetMissionDateFromPath(ByVal LV_PulsePath As String) As String

Dim lsCampo()       As String
Dim lsLastChar      As String
    
    lsCampo = Split(LV_PulsePath, "\")
    If UBound(lsCampo) >= 1 Then
        GetMissionDateFromPath = lsCampo(UBound(lsCampo) - 1)
        GetMissionDateFromPath = Right$(GetMissionDateFromPath, 19)
    End If
    
End Function

Function GetMissionNameFromPath(ByVal LV_PulsePath As String) As String

Dim lsCampo()       As String
Dim lsLastChar      As String
    'GetMissionName = Retroceder_Path(LV_PulsePath)
    lsCampo = Split(LV_PulsePath, "\")
    If UBound(lsCampo) >= 1 Then
        GetMissionNameFromPath = lsCampo(UBound(lsCampo) - 1)
        If Len(GetMissionNameFromPath) < 19 Then
            GetMissionNameFromPath = ""
            Exit Function
        End If
        GetMissionNameFromPath = Left$(GetMissionNameFromPath, Len(GetMissionNameFromPath) - 19)
        lsLastChar = Right$(GetMissionNameFromPath, 1)
        If lsLastChar = "_" Or lsLastChar = "-" Then
            GetMissionNameFromPath = Left$(GetMissionNameFromPath, Len(GetMissionNameFromPath) - 1)
        End If
    End If
    
End Function

Function GetFileName(LV_File As String) As String

Dim lsStr()                 As String

    lsStr = Split(LV_File, "\")
    On Error Resume Next
    
    GetFileName = lsStr(UBound(lsStr))
    
End Function

Function IsNothing(lvObj As Object) As Boolean

Dim LV_Form         As Form

   On Error GoTo IsNothing_Error

    IsNothing = True
    If lvObj.Name <> "" Then
        IsNothing = False
    End If
    
   On Error GoTo 0
   Exit Function

IsNothing_Error:

    IsNothing = True
    
End Function
