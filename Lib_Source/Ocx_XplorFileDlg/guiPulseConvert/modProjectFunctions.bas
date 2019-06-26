Attribute VB_Name = "modProjectFunctions"
'---------------------------------------------------------------------------------------
' Module    : modProjectFunctions
' Author    : Leo Herrera
' Date      : 07/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public mProject        As Project

Private PV_Pool_Memory_Initialized      As Boolean

Sub SetLabelsFieldActualProject(LstVw As ListView)

Dim LstSbItm        As ListSubItem

    With LstVw.ListItems
        .Add , , "Nombre"
        .Add , , "Ubicación"
        .Add , , "Salida"
        .Add , , "Estado Salida"
        .Add , , "Ubicación Pulsos"
        .Add , , "Tiempo Inicio Pulsos"
        .Add , , "Tiempo Fin Pulsos"
        .Add , , "Archivos de Salida"
        .Add , , "Hojas por Archivo"
        .Add , , "Configuración"
    End With
    
End Sub

Function GetValueFieldProject(Index As Integer) As String

    Select Case Index
        Case Is = 1
            GetValueFieldProject = mProject.ProjectName
        Case Is = 2
            GetValueFieldProject = mProject.ProjectPath
        Case Is = 3
            GetValueFieldProject = mProject.OutputPath
        Case Is = 4
            'GetValueFieldProject = mProject
        Case Is = 5
            GetValueFieldProject = mProject.PulsePath
        Case Is = 6
            GetValueFieldProject = mProject.PulseIniTime
        Case Is = 7
            GetValueFieldProject = mProject.PulseEndTime
        'Case Is = 1
        '    GetValueFieldProject = mProject.p
        Case Else
            GetValueFieldProject = ""
    End Select
    
End Function

Sub ShowActualProject(LstVw As ListView)

Dim i               As Integer
Dim lsLbl           As String

    With mProject
        If LstVw.ListItems.Count = 0 Then
            SetLabelsFieldActualProject LstVw
        End If
        For i = 1 To LstVw.ListItems.Count
            lsLbl = GetValueFieldProject(i)
            If LstVw.ListItems(i).ListSubItems.Count = 0 Then
                LstVw.ListItems(i).ListSubItems.Add , , lsLbl
            Else
                LstVw.ListItems(i).ListSubItems(1).Text = lsLbl
            End If
        Next
    End With
    ModListViewFunctions.AutoAjusteColumnWidth LstVw
    
End Sub

Sub SendFieldFormat()

Dim ColumnOrderArray()      As Long
Dim ColumnVisibleArray()    As Long
Dim i                       As Integer

    With mProject.ColumnConfig
        If .Count Then
        ReDim ColumnOrderArray(.Count - 1)
        ReDim ColumnVisibleArray(.Count - 1)
        
        For i = 0 To .Count - 1
            ColumnOrderArray(i) = .Column(i).Order
            If .Column(i).Visible = True Then
                ColumnVisibleArray(i) = 1
            Else
                ColumnVisibleArray(i) = 0
            End If
        Next
        Pulse_SetFieldFormat ColumnOrderArray(0), ColumnVisibleArray(0)
        End If
    End With
    
End Sub

Function Add_Project_Extension(FileName As String) As String

Dim lvFlag          As Boolean

    lvFlag = False
    If Len(FileName) >= Len(GC_PROJECT_EXTENSION) Then
        If Right$(FileName, Len(GC_PROJECT_EXTENSION)) <> GC_PROJECT_EXTENSION Then
            lvFlag = True
        End If
    Else
        lvFlag = True
    End If
    If lvFlag = True Then
        Add_Project_Extension = FileName & GC_PROJECT_EXTENSION
    Else
        Add_Project_Extension = FileName
    End If
    
End Function

Sub AddProjectToList()

Dim lvListProject       As ProjectList
Dim flagSave            As Boolean
Dim Index               As Integer

    flagSave = False
    Index = -1
    ReadListProject GV_ListProjectFile, lvListProject
    
    If lvListProject.Count Then
        If IsProjectInList(lvListProject, mProject) = False Then
            Index = lvListProject.Count
            ReDim Preserve lvListProject.List(lvListProject.Count)
            lvListProject.Count = lvListProject.Count + 1
            flagSave = True
        End If
    Else
        ReDim lvListProject.List(0)
        flagSave = True
        Index = 0
        lvListProject.Count = 1
    End If
    
    If Index >= 0 Then
        With lvListProject.List(Index)
            .ProjectPath = mProject.ProjectPath
            .ProjectName = mProject.ProjectName
        End With
    End If
    If flagSave = True Then
        SaveListProject GV_ListProjectFile, lvListProject
    End If
    
End Sub

Sub ClearProject()

    With mProject
        .MissionDate = ""
        .MissionName = ""
        .MissionPath = ""
        .OutputPath = ""
        .ProjectName = ""
        .ProjectPath = ""
        .ProjectFileName = ""
        .PulsePath = ""
        .PulseEndTime = ""
        .PulseIniTime = ""
        .PulseFileCount = 0
        .WorkSpacePath = ""
        
        .ProjectClosed = False
        .ProjectEmpty = False
        .NewProject = True
        .PulsesAnalized = False
        .Saved = False
        .Changed = False
        .SheetConfigured = False
        .SheetGenerated = False
        .IntermediateDataReady = False
        .ProjectFolderSelected = False
        .MissionSelected = False
        
        .AsociaMissionName = 1
        .CreateFolderProject = 1
        .OutputSubFolder = 1
        .WrkSpcSubFolder = 1
        
        .ColumnConfig.Count = 0
        .SpreadConfig.SpreadConfigName = ""

        .FilesPerWorkSpace = GetSettingFilesPerWorkSpace
        '
    End With
    
End Sub

Function GetSettingFilesPerWorkSpace() As Long

    GetSettingFilesPerWorkSpace = GetSetting(App.Title, _
                                    GC_CONFIGURATION_SECTION, _
                                    ".FilesPerWorkSpace", 32)
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, _
                                    ".FilesPerWorkSpace", GetSettingFilesPerWorkSpace

End Function

Function GetCaptionForm(ByVal lvFormName As String) As String

    If lvFormName = "frmStart" Then
        GetCaptionForm = "Pantalla de Inicio"
    ElseIf mProject.NewProject = True Then
        GetCaptionForm = "Proyecto Nuevo"
    Else
        GetCaptionForm = "Misión " & mProject.MissionName
    End If
    
End Function

Sub CloseProject()

Dim msg             As String

    If mProject.Changed = True Then
        If mProject.NewProject = True Then
            msg = "El Proyecto Nuevo no ha sido guardado. " & _
                "¿Desea guardar los cambios?"
            If MsgBox(msg, vbYesNo, "Cerrando Proyecto") = vbYes Then
                SaveProject
            End If
        Else
            msg = "El proyecto actual se va a cerrar. ¿Desea guardar los cambios?"
            If MsgBox(msg, vbYesNo, "Cerrando Proyecto") = vbYes Then
                SaveProject
            End If
        End If
    End If
    DiscardProject
    
End Sub

Sub CreateFolders()

    Create_Folder mProject.ProjectPath
    Create_Folder mProject.OutputPath
    Create_Folder mProject.WorkSpacePath
    
End Sub

Sub DiscardWorkSpace()

    Pulse_DestroyWorkSpace

End Sub

Sub DiscardProject()

    With mProject
        .MissionPath = ""
        .OutputPath = ""
        .ProjectName = ""
        .ProjectPath = ""
        .PulsePath = ""
        .WorkSpacePath = ""
        
        .ProjectClosed = True
        .ProjectEmpty = True
        .NewProject = False
        .PulsesAnalized = False
        .Saved = False
        .Changed = False
        .SheetConfigured = False
        .SheetGenerated = False
        .IntermediateDataReady = False
    End With
    
End Sub

Function GetFilesCount() As Long

    GetFilesCount = mProject.PulseFileCount
    
End Function

Function GetMissionDate() As String

    GetMissionDate = mProject.MissionDate

End Function

Function GetMissionName() As String

    GetMissionName = mProject.MissionName

End Function

Function GetMissionPath() As String

    GetMissionPath = mProject.MissionPath

End Function

Function GetPulseIniTime() As String

    GetPulseIniTime = mProject.PulseIniTime

End Function

Function GetPulseEndTime() As String

    GetPulseEndTime = mProject.PulseEndTime

End Function

Function GetPulsePath() As String

    GetPulsePath = mProject.PulsePath

End Function


Function GetName() As String

    GetName = mProject.ProjectName

End Function

Function GetOutputPath() As String

    GetOutputPath = mProject.OutputPath
    
End Function

Function GetProjectPath() As String

    GetProjectPath = mProject.ProjectPath
    
End Function

Function GetAsociaMissionName() As Integer

    GetAsociaMissionName = mProject.AsociaMissionName

End Function

Function GetCreateFolderProject() As Integer

    GetCreateFolderProject = mProject.CreateFolderProject

End Function

Function GetOutputSubFolder() As Integer

    GetOutputSubFolder = mProject.OutputSubFolder

End Function

Function GetWrkSpcSubFolder() As Integer

    GetWrkSpcSubFolder = mProject.WrkSpcSubFolder

End Function

Function IsClosedProject() As Boolean

    IsClosedProject = mProject.ProjectClosed

End Function

Function IsEmptyProject() As Boolean

    IsEmptyProject = mProject.ProjectEmpty

End Function

Function IsNewProject() As Boolean

    IsNewProject = mProject.NewProject
    
End Function

Function IsProjectInRemoteList(lvListProject As RemoteProjectList, _
                                lvProject As Project) As Boolean

Dim i           As Integer

    If lvListProject.Count Then
        For i = 0 To lvListProject.Count - 1
            If lvListProject.ProjectList(i).ProjectName = lvProject.ProjectName Then
                If lvListProject.ProjectList(i).ProjectPath = lvProject.ProjectPath Then
                    IsProjectInRemoteList = True
                    Exit Function
                End If
            End If
        Next
    End If
    IsProjectInRemoteList = False
    
End Function

Function IsProjectInList(lvListProject As ProjectList, lvProject As Project) As Boolean

Dim i           As Integer

    If lvListProject.Count Then
        For i = 0 To lvListProject.Count - 1
            If lvListProject.List(i).ProjectName = lvProject.ProjectName Then
                If lvListProject.List(i).ProjectPath = lvProject.ProjectPath Then
                    IsProjectInList = True
                    Exit Function
                End If
            End If
        Next
    End If
    IsProjectInList = False
    
End Function

Sub Project_Constructor()

    ClearProject
    With mProject
        .ProjectClosed = True
        .ProjectEmpty = True
        .MissionSelected = False
        .ProjectFolderSelected = False
        .FilesPerWorkSpace = GetSettingFilesPerWorkSpace
    End With
    
End Sub

Sub Project_Destructor()

    CloseProject
    
End Sub

Sub NewProject()

    ClearProject
    With mProject
        .ProjectClosed = False
        .ProjectEmpty = False
        .NewProject = True
        .Saved = False
        .ExportQueued = False
        .SheetConfigured = False
        .SheetGenerated = False
        .SheetGenerating = False
    End With

End Sub

Sub ReadListProject(lvFile As String, lv_ListProject As ProjectList)

Dim h           As Integer

    'On Error GoTo ReadListProject_Error

    h = FreeFile
    Open GV_ListProjectFile For Binary Access Read As h
    
    Get #h, , lv_ListProject

    Close h
    
    'On Error GoTo 0
    Exit Sub

ReadListProject_Error:

    lv_ListProject.Count = 0
    
End Sub

Sub SaveListProject(lvFile As String, lv_ListProject As ProjectList)

Dim h           As Integer
    
    'On Error GoTo SaveListProject_Error

    h = FreeFile
    Open GV_ListProjectFile For Binary Access Write As h
    
    Put #h, , lv_ListProject

    Close h
    
    'On Error GoTo 0
    Exit Sub

SaveListProject_Error:

    
End Sub

Sub SetPulsesTime(lvTimeIni As String, lvTimeEnd As String)
    
    mProject.PulseIniTime = lvTimeIni
    mProject.PulseEndTime = lvTimeEnd
    
End Sub

Sub SetProjectFolderSelected()

    mProject.ProjectFolderSelected = True
    
End Sub

Sub ClearProjectFolderSelected()

    mProject.ProjectFolderSelected = False
    
End Sub

Function GetProjectFolderSelected() As Boolean

    GetProjectFolderSelected = mProject.ProjectFolderSelected
    
End Function

Sub SetMissionSelected()

    mProject.MissionSelected = True
    
End Sub

Sub ClearMissionSelected()

    mProject.MissionSelected = False
    
End Sub

Function GetMissionSelected() As Boolean

    GetMissionSelected = mProject.MissionSelected
    
End Function

Sub SetFileCount(lvPulseFilesCount As Long)

    mProject.PulseFileCount = lvPulseFilesCount
    
End Sub

Sub SetMissionPath(lvPath As String)

    mProject.MissionPath = lvPath
    mProject.MissionName = GetMissionNameFromMissionPath(lvPath)
    mProject.MissionDate = GetMissionDateFromMissionPath(lvPath)
    
End Sub

Sub SetPulsePath(lvPath As String)

    mProject.PulsePath = lvPath
    
End Sub

Sub SetName(NewName As String)

    mProject.ProjectName = NewName
    'modStructDatos.Pulse_Set_ProjectName NewName

End Sub

Sub SetSheetGenerated()

    mProject.SheetGenerated = True
    mProject.Changed = True
    
End Sub

Sub ClearSheetGenerated()

    mProject.SheetGenerated = False
    mProject.Changed = True
    
End Sub

Function GetSheetGenerated() As Boolean

    GetSheetGenerated = mProject.SheetGenerated
    
End Function

Sub SetSheetGenerating()

    mProject.SheetGenerating = True
    mProject.Changed = True
    
End Sub

Sub ClearSheetGenerating()

    mProject.SheetGenerating = False
    mProject.Changed = True
    
End Sub

Function GetSheetGenerating() As Boolean

    GetSheetGenerating = mProject.SheetGenerating
    mProject.Changed = True
    
End Function

Sub SetExportQueued()

    mProject.ExportQueued = True
    mProject.Changed = True
    
End Sub

Sub ClearExportQueued()

    mProject.ExportQueued = False
    
End Sub

Function GetExportQueued() As Boolean

    GetExportQueued = mProject.ExportQueued
    
End Function

Sub SetIntermediateDataReady()

    mProject.IntermediateDataReady = True
    mProject.Changed = True
    
End Sub

Sub ClearIntermediateDataReady()

    mProject.IntermediateDataReady = False
    mProject.ExportQueued = False
    mProject.SheetGenerating = False
    mProject.SheetGenerated = False
    mProject.Changed = True
    
End Sub

Function GetIntermediateDataReady() As Boolean

    GetIntermediateDataReady = mProject.IntermediateDataReady
    
End Function

Sub SetPulsesAnalized()

    mProject.PulsesAnalized = True
    mProject.Changed = True
    
End Sub

Sub ClearPulsesAnalized()

    mProject.PulsesAnalized = False
    mProject.Changed = True
    
End Sub

Function GetPulsesAnalized() As Boolean

    GetPulsesAnalized = mProject.PulsesAnalized
    
End Function

Sub SetOutputPath(NewPath As String)

    If mProject.OutputPath <> NewPath Then
        Create_Folder NewPath
        mProject.OutputPath = NewPath
        mProject.Changed = True
    End If
    
End Sub

Sub SetProjectPath(NewPath As String)

    Create_Folder NewPath
    mProject.ProjectPath = NewPath
    
End Sub

Sub SetWorkSpacePath(NewPath As String)

    Create_Folder NewPath
    mProject.WorkSpacePath = NewPath
    
End Sub

Function GetWorkSpacePath() As String

    GetWorkSpacePath = mProject.WorkSpacePath

End Function

Sub SetAsociaMissionName(lvValue As Integer)

    mProject.AsociaMissionName = lvValue

End Sub

Sub SetColumnConfig(LV_ColumnConfig As typeConfigSheetColumns)

    Cpy_Column_Config mProject.ColumnConfig, LV_ColumnConfig
    
End Sub

Sub SetSpreadConfig(LV_SpreadConfig As typeConfigSpreadSheet)

    Cpy_Spread_Config mProject.SpreadConfig, LV_SpreadConfig
    
End Sub

Sub SetSheetConfigured()

    mProject.SheetConfigured = True
    
End Sub

Sub ClearSheetConfigured()

    mProject.SheetConfigured = False
    
End Sub

Function GetSheetConfigured() As Boolean

    GetSheetConfigured = mProject.SheetConfigured
    
End Function

Sub SetCreateFolderProject(lvValue As Integer)

    mProject.CreateFolderProject = lvValue

End Sub

Sub SetOutputSubFolder(lvValue As Integer)

    mProject.OutputSubFolder = lvValue

End Sub

Sub SetWrkSpcSubFolder(lvValue As Integer)

    mProject.WrkSpcSubFolder = lvValue

End Sub

' If Locked => Verify = True
Function VerifyProjectLocked() As Boolean

Dim FileName    As String
Dim LV_Pjt      As Project

    FileName = mProject.ProjectPath & "\" & mProject.ProjectName
    FileName = Add_Project_Extension(FileName)
    LV_Pjt = LoadStructProject(FileName)
    VerifyProjectLocked = False
    If IsProjectInRemoteList(BackGroundProjectList, mProject) = True Then
        With LV_Pjt
            If .ExportQueued = True Then
                VerifyProjectLocked = True
            End If
            If .SheetGenerating = True Then
                VerifyProjectLocked = True
            End If
        End With
    End If
End Function

Function VerifyProjectSheetGenerated() As Boolean

Dim FileName    As String
Dim LV_Pjt      As Project

    FileName = mProject.ProjectPath & "\" & mProject.ProjectName
    FileName = Add_Project_Extension(FileName)
    LV_Pjt = LoadStructProject(FileName)
    VerifyProjectSheetGenerated = False
    With LV_Pjt
        If .SheetGenerated = True Then
            VerifyProjectSheetGenerated = True
        End If
    End With
    
End Function

Function VerifyProjectExist() As Boolean

Dim FileName    As String

    FileName = mProject.ProjectPath & "\" & mProject.ProjectName
    FileName = Add_Project_Extension(FileName)

    VerifyProjectExist = File_Exist(FileName)
    
End Function

Sub SetFilesPerWorkSpace(lvVal As Long)

    mProject.FilesPerWorkSpace = lvVal
    
End Sub

Function GetFilesPerWorkSpace() As Long

    GetFilesPerWorkSpace = mProject.FilesPerWorkSpace
    
End Function
Function Get_IntermediaFileCount() As Long
    
    Get_IntermediaFileCount = mProject.IntermediaFileCount

End Function

Sub Set_IntermediaFileCount()
    
    Pulse_SetIntermediaFileCount mProject.IntermediaFileCount
    Init_Pool_Memory

End Sub

Sub Init_Pool_Memory()

Dim Count                           As Long
Dim lFilePwdLstDataSize             As Long
Dim lFilePwdDataSize                As Long
Dim lPwdNFDataSize                  As Long
Dim lPwdDataSize                    As Long

Dim lFilePwdLstCount                As Long
Dim lFilePwdCount                   As Long
Dim lPwdNFCount                     As Long

    If PV_Pool_Memory_Initialized = True Then
        Exit Sub
    End If
    PV_Pool_Memory_Initialized = True
    
    Pulse_GetStructSize lFilePwdLstDataSize, lFilePwdDataSize, lPwdNFDataSize, lPwdDataSize
    
    Count = GetSettingFilesPerWorkSpace
    
    lFilePwdLstCount = 4
    lFilePwdCount = Count * lFilePwdLstCount
    lPwdNFCount = 3414 * lFilePwdCount
    
    SaveSetting App.Title, GC_MEMORY_SECTION, "lFilePwdLstDataSize", lFilePwdLstDataSize
    SaveSetting App.Title, GC_MEMORY_SECTION, "lFilePwdDataSize", lFilePwdDataSize
    SaveSetting App.Title, GC_MEMORY_SECTION, "lPwdNFDataSize", lPwdNFDataSize
    
    SaveSetting App.Title, GC_MEMORY_SECTION, "lFilePwdLstCount", lFilePwdLstCount
    SaveSetting App.Title, GC_MEMORY_SECTION, "lFilePwdCount", lFilePwdCount
    SaveSetting App.Title, GC_MEMORY_SECTION, "lPwdNFCount", lPwdNFCount
    
    SaveSetting App.Title, GC_MEMORY_SECTION, "FilePwdLst_Size[KB]", lFilePwdLstCount * lFilePwdLstDataSize / 1024
    SaveSetting App.Title, GC_MEMORY_SECTION, "FilePwd_Size[MB]", lFilePwdCount * lFilePwdDataSize / 1024 / 1024
    SaveSetting App.Title, GC_MEMORY_SECTION, "PwdNF_Size[MB]", lPwdNFCount * lPwdNFDataSize / 1024 / 1024
    SaveSetting App.Title, GC_MEMORY_SECTION, "Pwd_Size[MB]", lPwdNFCount * lPwdDataSize / 1024 / 1024
    
    SaveSetting App.Title, GC_MEMORY_SECTION, "Total_Memory_Pool[MB]", _
                (lFilePwdLstCount * lFilePwdLstDataSize / 1024 + _
                 lFilePwdCount * lFilePwdDataSize / 1024 + _
                 lPwdNFCount * (lPwdNFDataSize + lPwdDataSize) / 1024) / 1024
    
    Pulse_Init_PoolMemory lFilePwdLstCount, lFilePwdCount, lPwdNFCount

End Sub

Function Get_Import_Parameters(ByRef IndexFilePwdLst As Long, _
            ByRef IndexFile As Long, _
            ByRef FileCount As Long, _
            ByRef ProccessDone As Long) As Long

    Pulse_Import_File_Status IndexFilePwdLst, IndexFile, FileCount, ProccessDone
    mProject.IntermediaFileCount = IndexFilePwdLst
    Get_Import_Parameters = IndexFilePwdLst
    
End Function

Sub Set_Parameters_After_Pulse_Analize()

Dim IndexFilePwdLst As Long
Dim IndexFile As Long
Dim FileCount As Long
Dim ProccessDone As Long
    
    WriteLogFile "Sub Set_Parameters_After_Pulse_Analize"
    
    Pulse_Import_File_Status IndexFilePwdLst, IndexFile, FileCount, ProccessDone
    mProject.IntermediaFileCount = IndexFilePwdLst
    
        WriteLogFile "IndexFilePwdLst = " & IndexFilePwdLst
        WriteLogFile "IndexFile = " & IndexFile
        WriteLogFile "FileCount = " & FileCount
        WriteLogFile "ProccessDone = " & ProccessDone
        
    If IndexFilePwdLst > 0 Then
        mProject.PulsesAnalized = True
        mProject.NewProject = False
        WriteLogFile "PulsesAnalized = True"
    Else
        mProject.PulsesAnalized = False
        WriteLogFile "PulsesAnalized = False"
    End If

End Sub

Sub Run_Pulse_Analize_BG()

Dim lFiles          As Long
Dim lvResult        As Long
Dim lvMissionPath   As String
Dim lvPulsePath     As String
Dim lvTimeIni       As String
Dim lvTimeEnd       As String
Dim lvFilesCount    As Long

    If mProject.MissionSelected = True And _
                mProject.ProjectFolderSelected = True Then
        Pulse_FilesPerWorkSpace mProject.FilesPerWorkSpace
        Pulse_SetWorkSpacePath mProject.WorkSpacePath
        Pulse_OutputPath mProject.OutputPath
        'Run_Pulse_Analize = Pulse_Import_File(mProject.PulsePath)
        
        WriteLogFile "mProject.FilesPerWorkSpace = " & mProject.FilesPerWorkSpace
        WriteLogFile "mProject.WorkSpacePath = " & mProject.WorkSpacePath
        WriteLogFile "mProject.OutputPath = " & mProject.OutputPath
        'WriteLogFile
        
        Pulse_Import_File_BG mProject.PulsePath
        
'        mProject.IntermediaFileCount = Run_Pulse_Analize
'        If Run_Pulse_Analize > 0 Then
'            mProject.PulsesAnalized = True
'            mProject.NewProject = False
'        Else
'            mProject.PulsesAnalized = False
'        End If
'        mProject.Changed = True
        
    Else
        mProject.PulsesAnalized = False
    End If
    
End Sub

Function Run_Pulse_Analize() As Long

Dim lFiles          As Long
Dim lvResult        As Long
Dim lvMissionPath   As String
Dim lvPulsePath     As String
Dim lvTimeIni       As String
Dim lvTimeEnd       As String
Dim lvFilesCount    As Long

    Run_Pulse_Analize = -1
    If mProject.MissionSelected = True And _
                mProject.ProjectFolderSelected = True Then
        Pulse_FilesPerWorkSpace mProject.FilesPerWorkSpace
        Pulse_SetWorkSpacePath mProject.WorkSpacePath
        Pulse_OutputPath mProject.OutputPath
        Run_Pulse_Analize = Pulse_Import_File(mProject.PulsePath)
        mProject.IntermediaFileCount = Run_Pulse_Analize
        If Run_Pulse_Analize > 0 Then
            mProject.PulsesAnalized = True
            mProject.NewProject = False
        Else
            mProject.PulsesAnalized = False
        End If
        mProject.Changed = True
    Else
        mProject.PulsesAnalized = False
    End If
    
End Function

Sub RevertParameters()

    mProject.Changed = True
    mProject.ExportQueued = False
    mProject.SheetGenerating = False
    
End Sub

Function SaveProject(Optional LV_Force As Boolean = False) As Boolean

Dim h           As Integer
Dim FileName    As String

    SaveProject = False
    If mProject.ProjectName <> "" Then
        If mProject.NewProject = True Then
            If LV_Force = False And GV_PrevInstance = False And VerifyProjectExist = True Then
                If VerifyProjectLocked = True And _
                    IsProjectInRemoteList(BackGroundProjectList, mProject) = True Then
                    MsgBox _
                        "Está creando un Proyecto el cual está exportando archivos. " & _
                        "Este proyecto está bloqueado. " & _
                        "Por favor, escoja otros pulsos y otra ubicación", _
                        vbOKOnly, "Archivo Existente Bloqueado"
                    Exit Function
                Else
                    If VerifyProjectSheetGenerated = True Then
                        If MsgBox( _
                            "Está creando un Proyecto el cual tiene todos sus Archivos de Exportación generados. " & _
                            "Si continúa, podría peerder los archivos ya generados. " & _
                            "¿Desea continuar creando este Proyecto?", _
                            vbYesNo, "Archivos Excel Generados") _
                            <> vbYes Then
                            Exit Function
                        End If
                    Else
                        If MsgBox( _
                            "El archivo del proyecto que está creando ya existe. ¿Desea sobreescribirlo?", vbYesNo, "Sobreescribir Archivo Existente") _
                            <> vbYes Then
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        Create_Folder mProject.ProjectPath
        mProject.Changed = False
        mProject.Saved = True
        FileName = mProject.ProjectPath & "\" & mProject.ProjectName
        FileName = Add_Project_Extension(FileName)
        mProject.ProjectFileName = FileName
        
        SaveProjectStruct FileName, mProject
        
        VerificarGrabacionProyecto mProject
        SaveProject = True
    End If
    
End Function

Sub SaveProjectStruct(FileName As String, lPjt As Project)

Dim h           As Integer

    h = FreeFile
    On Error Resume Next
    Kill FileName
    'On Error GoTo 0
    Open FileName For Binary Access Write As h
    
    If True Then
        Put #h, , lPjt
    Else
        Put #h, , lPjt.PulseFileCount
        Put #h, , lPjt.IntermediaFileCount
        Put #h, , lPjt.FilesPerWorkSpace
        Put #h, , lPjt.Spare_Long
    
        Put #h, , lPjt.AsociaMissionName
        Put #h, , lPjt.CreateFolderProject
        Put #h, , lPjt.OutputSubFolder
        Put #h, , lPjt.WrkSpcSubFolder
    
        Put #h, , lPjt.ProjectEmpty
        Put #h, , lPjt.ProjectClosed
        Put #h, , lPjt.NewProject
        Put #h, , lPjt.Changed
    
        Put #h, , lPjt.Saved
        Put #h, , lPjt.SheetConfigured
        Put #h, , lPjt.PulsesAnalized
        Put #h, , lPjt.IntermediateDataReady
    
        Put #h, , lPjt.SheetGenerated
        Put #h, , lPjt.SheetGenerating
        Put #h, , lPjt.ExportQueued
        Put #h, , lPjt.ProjectFolderSelected
    
        Put #h, , lPjt.MissionSelected
        Put #h, , lPjt.Spare_Boolean1
        Put #h, , lPjt.Spare_Boolean2
        Put #h, , lPjt.Spare_Boolean3
    
    
        SaveStringFromFile h, lPjt.ProjectName
        SaveStringFromFile h, lPjt.ProjectPath
        SaveStringFromFile h, lPjt.ProjectFileName
        SaveStringFromFile h, lPjt.OutputPath
        SaveStringFromFile h, lPjt.WorkSpacePath
        SaveStringFromFile h, lPjt.PulsePath
        SaveStringFromFile h, lPjt.MissionPath
        SaveStringFromFile h, lPjt.MissionDate
        SaveStringFromFile h, lPjt.MissionName
        SaveStringFromFile h, lPjt.PulseIniTime
        SaveStringFromFile h, lPjt.PulseEndTime
        
        SaveInFileColumnConfig h, lPjt.ColumnConfig
        SaveInFileSpreadSheetConfig h, lPjt.SpreadConfig
    End If
    
    Close h

End Sub

Sub SaveInFileColumnConfig(h As Integer, lvColumnConfig As typeConfigSheetColumns)

Dim i           As Integer

    With lvColumnConfig
        Put #h, , .Count
        SaveStringFromFile h, .ColumnConfigName
        For i = 0 To .Count - 1
            SaveStringFromFile h, .Column(i).ColumnName
            Put #h, , .Column(i).Order
            Put #h, , .Column(i).Visible
        Next
    End With
    
End Sub

Sub SaveInFileSpreadSheetConfig(h As Integer, lvSpreadConfig As typeConfigSpreadSheet)

    With lvSpreadConfig
        Put #h, , .byFiles
        Put #h, , .byInterval
        Put #h, , .byPulses
        Put #h, , .IntervalPerSheet
        Put #h, , .PulsesPerSheet
        Put #h, , .SheetsPerSpreadSheet
        SaveStringFromFile h, .SpreadConfigName
    End With
    
End Sub

Function CompararProyectos(LV_Pjt_1 As Project, LV_Pjt_2 As Project) As Boolean

End Function

Sub VerificarGrabacionProyecto(LV_Pjt_Saved As Project)

Dim LV_Pjt_Loaded       As Project
Dim h                   As Integer
'Dim FilePath            As String
Dim FileName            As String

    FileName = LV_Pjt_Saved.ProjectFileName
'    h = FreeFile
'    Open FileName For Binary Access Read As h
'    Get #h, , LV_Pjt_Loaded
'    Close h
    LV_Pjt_Loaded = LoadStructProject(FileName)
    CompararProyectos LV_Pjt_Saved, LV_Pjt_Loaded
    
End Sub

Sub LoadInfoProject(LV_LstVw As ListView)

Dim i               As Integer

    For i = 1 To LV_LstVw.ListItems.Count
        With LV_LstVw.ListItems(i).ListSubItems
            .Clear
            Select Case i
                Case Is = 1
                    .Add , , mProject.ProjectName
                Case Is = 2
                    .Add , , mProject.ProjectName
                Case Is = 3
                    .Add , , mProject.ProjectPath
                Case Is = 4
                    .Add , , mProject.MissionPath
                Case Is = 5
                    .Add , , mProject.OutputPath
                Case Is = 6
                    .Add , , IfBooleanText(mProject.PulsesAnalized, "Completado", "No realizado")
                Case Is = 7
                    .Add , , IfBooleanText(mProject.IntermediateDataReady, "Completado", "No realizado")
                Case Is = 8
                    .Add , , IfBooleanText(mProject.SheetGenerated, "Completado", "No realizado")
                Case Is = 9
            End Select
        End With
    Next
    
End Sub

Function LoadProject(ByVal FileName As String, Optional FilePath As String) As Project

Dim h           As Integer
Dim lPjt        As Project

    If FilePath <> "" Then
        FileName = FilePath & "\" & FileName
    End If
    
    FileName = Add_Project_Extension(FileName)
    
    
    mProject = LoadStructProject(FileName)
    
    Sync_Project
    
    LoadProject = mProject
    
End Function

Function LoadStructProject(FileName As String) As Project

Dim h       As Integer
Dim lPjt    As Project

    h = FreeFile
    Open FileName For Binary Access Read As h
       
    If True Then
        
        Get #h, , lPjt
        LoadStructProject = lPjt
    Else
        Get #h, , LoadStructProject.PulseFileCount
        Get #h, , LoadStructProject.IntermediaFileCount
        Get #h, , LoadStructProject.FilesPerWorkSpace
        Get #h, , LoadStructProject.Spare_Long
    
        Get #h, , LoadStructProject.AsociaMissionName
        Get #h, , LoadStructProject.CreateFolderProject
        Get #h, , LoadStructProject.OutputSubFolder
        Get #h, , LoadStructProject.WrkSpcSubFolder
    
        Get #h, , LoadStructProject.ProjectEmpty
        Get #h, , LoadStructProject.ProjectClosed
        Get #h, , LoadStructProject.NewProject
        Get #h, , LoadStructProject.Changed
    
        Get #h, , LoadStructProject.Saved
        Get #h, , LoadStructProject.SheetConfigured
        Get #h, , LoadStructProject.PulsesAnalized
        Get #h, , LoadStructProject.IntermediateDataReady
    
        Get #h, , LoadStructProject.SheetGenerated
        Get #h, , LoadStructProject.SheetGenerating
        Get #h, , LoadStructProject.ExportQueued
        Get #h, , LoadStructProject.ProjectFolderSelected
    
        Get #h, , LoadStructProject.MissionSelected
        Get #h, , LoadStructProject.Spare_Boolean1
        Get #h, , LoadStructProject.Spare_Boolean2
        Get #h, , LoadStructProject.Spare_Boolean3
    
    
        LoadStructProject.ProjectName = LoadStringFromFile(h)
        LoadStructProject.ProjectPath = LoadStringFromFile(h)
        LoadStructProject.ProjectFileName = LoadStringFromFile(h)
        LoadStructProject.OutputPath = LoadStringFromFile(h)
        LoadStructProject.WorkSpacePath = LoadStringFromFile(h)
        LoadStructProject.PulsePath = LoadStringFromFile(h)
        LoadStructProject.MissionPath = LoadStringFromFile(h)
        LoadStructProject.MissionDate = LoadStringFromFile(h)
        LoadStructProject.MissionName = LoadStringFromFile(h)
        LoadStructProject.PulseIniTime = LoadStringFromFile(h)
        LoadStructProject.PulseEndTime = LoadStringFromFile(h)
        
        LoadFromFileColumnConfig h, LoadStructProject.ColumnConfig
        LoadFromFileSpreadSheetConfig h, LoadStructProject.SpreadConfig
    End If
    
    Close h

End Function

Function LoadFromFileColumnConfig(h As Integer, lvColumnConfig As typeConfigSheetColumns) As Boolean

Dim i           As Integer

    With lvColumnConfig
        Get #h, , .Count
        .ColumnConfigName = LoadStringFromFile(h)
        If .Count Then
            ReDim .Column(.Count - 1)
            For i = 0 To .Count - 1
                .Column(i).ColumnName = LoadStringFromFile(h)
                Get #h, , .Column(i).Order
                Get #h, , .Column(i).Visible
            Next
        End If
    End With
    
End Function

Function LoadFromFileSpreadSheetConfig(h As Integer, lvSpreadConfig As typeConfigSpreadSheet) As Boolean

    With lvSpreadConfig
        Get #h, , .byFiles
        Get #h, , .byInterval
        Get #h, , .byPulses
        Get #h, , .IntervalPerSheet
        Get #h, , .PulsesPerSheet
        Get #h, , .SheetsPerSpreadSheet
        .SpreadConfigName = LoadStringFromFile(h)
    End With
    LoadFromFileSpreadSheetConfig = True
    
End Function

Function LoadStringFromFile(h As Integer) As String

Dim lvLargo         As Long
    
    Get #h, , lvLargo
    If lvLargo > 290 Then
        LoadStringFromFile = ""
        Exit Function
    End If
    If lvLargo Then
        LoadStringFromFile = Space(lvLargo)
        Get #h, , LoadStringFromFile
    Else
        LoadStringFromFile = ""
    End If
    
End Function

Sub SaveStringFromFile(h As Integer, lvStr As String)

Dim lvLargo         As Long
    
    lvLargo = Len(lvStr)
    Put #h, , lvLargo
    If lvLargo Then
        Put #h, , lvStr
    End If
    
End Sub

Sub LoadWorkSpace()

    Pulse_LoadWorkSpace
    
End Sub

Sub Set_Spread_Config()

End Sub

Sub Set_Column_Config()

End Sub

Sub Generate_Intermedia_Data()

End Sub

Sub Sync_Project()

    If mProject.ProjectEmpty = False Then
        Pulse_FilesPerWorkSpace mProject.FilesPerWorkSpace
        Pulse_SetWorkSpacePath mProject.WorkSpacePath
        Pulse_OutputPath mProject.OutputPath
        Set_IntermediaFileCount
        If mProject.IntermediateDataReady = True Then
            Set_Spread_Config
            Set_Column_Config
            'Generate_Intermedia_Data
        End If
    End If
    
End Sub

Function IsThereProjectList() As Boolean

    IsThereProjectList = False
    'On Error GoTo IsThereProjectList_Error

    IsThereProjectList = File_Exist(GV_ListProjectFile)

    'On Error GoTo 0
    Exit Function

IsThereProjectList_Error:

    
End Function

Sub ShowListProjects(LV_LstVw As ListView)

Dim lvLstSubItms        As ListSubItems
Dim lvListProject       As ProjectList
Dim Index               As Integer

    LV_LstVw.ListItems.Clear
    ReadListProject GV_ListProjectFile, lvListProject
    If lvListProject.Count = 0 Then
        Exit Sub
    End If
    For Index = 0 To lvListProject.Count - 1
        Set lvLstSubItms = LV_LstVw.ListItems.Add(, , Index + 1)
        lvLstSubItms.Add , , lvListProject.List(Index).ProjectName
        lvLstSubItms.Add , , lvListProject.List(Index).ProjectPath
    Next
    
End Sub

Sub SetXlsDll()

Dim lsXlsDll            As String
Dim lsXlsDef            As String
Dim StrBookConstructor  As String
Dim StrBookDestructor   As String
Dim StrBookSave         As String
Dim StrBookSetHeader    As String
Dim StrBookSetSheet     As String
Dim StrBookSetOrder     As String
Dim StrBookCvtBin       As String

Dim lsProcedureList()   As String

    If GV_PrevInstance = False Then
        Exit Sub
    End If
    lsXlsDll = GV_XlsDll_Path & "\test.dll"
    lsXlsDef = GV_XlsDll_Path & "\libtest.dll.def"
    
    LoadProcedureList lsXlsDef, lsProcedureList
    
    StrBookConstructor = FindItemContain(lsProcedureList, "StrBookConstructor")
    StrBookDestructor = FindItemContain(lsProcedureList, "StrBookDestructor")
    StrBookSave = FindItemContain(lsProcedureList, "StrBookSave")
    StrBookSetHeader = FindItemContain(lsProcedureList, "StrBookSetHeader")
    StrBookSetSheet = FindItemContain(lsProcedureList, "StrBookSetSheet")
    StrBookSetOrder = FindItemContain(lsProcedureList, "StrBookSetOrder")
    StrBookCvtBin = FindItemContain(lsProcedureList, "StrBookCvtBin")
    'StrBookCvtBin = FindItemContain(lsProcedureList, "test1")
    
    Pulse_Set_Xls_Dll lsXlsDll, _
            StrBookConstructor, _
            StrBookDestructor, _
            StrBookSave, _
            StrBookSetHeader, _
            StrBookSetSheet, _
            StrBookSetOrder, _
            StrBookCvtBin

End Sub

Sub LoadProcedureList(lsFile As String, lsList() As String)

Dim Count           As Integer
Dim h               As Integer
Dim lsLine          As String

    ReDim lsList(99)
    Count = 0
    h = FreeFile
    
    Open lsFile For Input Access Read As h
    Do
        If EOF(h) = False Then
            Input #h, lsLine
            lsList(Count) = FirstWord(lsLine)
            Count = Count + 1
            If Count > UBound(lsList) Then
                ReDim Preserve lsList(Count + 99)
            End If
        Else
            Exit Do
        End If
    Loop Until EOF(h) = True
    Close #h
    
End Sub

Function FirstWord(lsStr As String)

Dim lsArray()           As String

    If lsStr <> "" Then
        lsArray = Split(lsStr, " ")
        FirstWord = lsArray(0)
    Else
        FirstWord = lsStr
    End If
    
End Function

Function FindItemContain(ByRef lsLst() As String, ByVal lsCriteria As String) As String

Dim i           As Integer

    i = InStr(lsCriteria, "Book")
    If i Then
        lsCriteria = Right$(lsCriteria, Len(lsCriteria) - (Len("Book") + i))
    End If
    FindItemContain = ""
    For i = LBound(lsLst) To UBound(lsLst)
        If InStr(lsLst(i), lsCriteria) Then
            FindItemContain = lsLst(i)
            Exit Function
        End If
    Next
    
End Function

