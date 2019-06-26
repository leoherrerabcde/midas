VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub SendFieldFormat()

    modProjectFunctions.SendFieldFormat
    
End Sub

Sub SetSheetConfigured()

    modProjectFunctions.SetSheetConfigured
    
End Sub

Sub ClearSheetConfigured()

    modProjectFunctions.ClearSheetConfigured
    
End Sub

Function GetSheetConfigured() As Boolean

    GetSheetConfigured = modProjectFunctions.GetSheetConfigured
    
End Function

Sub SetName(NewName As String)

    modProjectFunctions.SetName NewName

End Sub

Sub NewProject()

    modProjectFunctions.NewProject
    
End Sub

Sub ClearProject()

    modProjectFunctions.ClearProject
    
End Sub

Sub CloseProject()

    modProjectFunctions.CloseProject
    
End Sub

Sub CreateFolders()

    modProjectFunctions.CreateFolders
    
End Sub


Sub DiscardWorkSpace()

    modProjectFunctions.DiscardWorkSpace

End Sub


Sub DiscardProject()

    modProjectFunctions.DiscardProject
    
End Sub
Function GetName() As String

    GetName = modProjectFunctions.GetName

End Function

Function GetFilesCount() As Long

    GetFilesCount = modProjectFunctions.GetFilesCount
    
End Function

Function GetMissionDate() As String

    GetMissionDate = modProjectFunctions.GetMissionDate

End Function

Function GetMissionName() As String

    GetMissionName = modProjectFunctions.GetMissionName

End Function

Function GetMissionPath() As String

    GetMissionPath = modProjectFunctions.GetMissionPath

End Function

Function GetPulseIniTime() As String

    GetPulseIniTime = modProjectFunctions.GetPulseIniTime

End Function

Function GetPulseEndTime() As String

    GetPulseEndTime = modProjectFunctions.GetPulseEndTime

End Function

Function GetPulsePath() As String

    GetPulsePath = modProjectFunctions.GetPulsePath

End Function

Function GetAsociaMissionName() As Integer

    GetAsociaMissionName = modProjectFunctions.GetAsociaMissionName

End Function

Function GetCreateFolderProject() As Integer

    GetCreateFolderProject = modProjectFunctions.GetCreateFolderProject

End Function

Function GetOutputSubFolder() As Integer

    GetOutputSubFolder = modProjectFunctions.GetOutputSubFolder

End Function

Function GetWrkSpcSubFolder() As Integer

    GetWrkSpcSubFolder = modProjectFunctions.GetWrkSpcSubFolder

End Function

Sub SetProjectFolderSelected()

    modProjectFunctions.SetProjectFolderSelected
    
End Sub

Sub ClearProjectFolderSelected()

    modProjectFunctions.ClearProjectFolderSelected
    
End Sub

Function GetProjectFolderSelected() As Boolean

    GetProjectFolderSelected = modProjectFunctions.GetProjectFolderSelected
    
End Function

Sub SetMissionSelected()

    modProjectFunctions.SetMissionSelected
    
End Sub

Sub ClearMissionSelected()

    modProjectFunctions.ClearMissionSelected
    
End Sub

Function GetMissionSelected() As Boolean

    GetMissionSelected = modProjectFunctions.GetMissionSelected
    
End Function

Sub SetFileCount(lvPulseFilesCount As Long)

    modProjectFunctions.SetFileCount lvPulseFilesCount
    
End Sub

Sub SetMissionPath(lvPath As String)

    modProjectFunctions.SetMissionPath lvPath
    
End Sub

Sub SetPulsePath(lvPath As String)

    modProjectFunctions.SetPulsePath lvPath
    
End Sub

Sub SetPulsesTime(lvTimeIni As String, lvTimeEnd As String)

    modProjectFunctions.SetPulsesTime lvTimeIni, lvTimeEnd
    
End Sub

Sub SetOutputPath(NewPath As String)

    modProjectFunctions.SetOutputPath NewPath
    
End Sub

Function GetOutputPath() As String

    GetOutputPath = modProjectFunctions.GetOutputPath
    
End Function

Sub SetProjectPath(NewPath As String)

    modProjectFunctions.SetProjectPath NewPath
    
End Sub

Function GetProjectPath() As String

    GetProjectPath = modProjectFunctions.GetProjectPath
    
End Function

Sub SetWorkSpacePath(NewPath As String)

    modProjectFunctions.SetWorkSpacePath NewPath
    
End Sub

Function GetWorkSpacePath() As String

    GetWorkSpacePath = modProjectFunctions.GetWorkSpacePath

End Function

Sub SetFilesPerWorkSpace(lvVal As Long)

    modProjectFunctions.SetFilesPerWorkSpace lvVal
    
End Sub

Function GetFilesPerWorkSpace() As Long

    GetFilesPerWorkSpace = modProjectFunctions.GetFilesPerWorkSpace
    
End Function

Sub SetAsociaMissionName(lvValue As Integer)

    modProjectFunctions.SetAsociaMissionName lvValue

End Sub

Sub SetCreateFolderProject(lvValue As Integer)

    modProjectFunctions.SetCreateFolderProject lvValue

End Sub

Sub SetOutputSubFolder(lvValue As Integer)

    modProjectFunctions.SetOutputSubFolder lvValue

End Sub

Sub SetWrkSpcSubFolder(lvValue As Integer)

    modProjectFunctions.SetWrkSpcSubFolder lvValue

End Sub

Function Get_IntermediaFileCount() As Long
    
    Get_IntermediaFileCount = modProjectFunctions.Get_IntermediaFileCount

End Function

Sub Set_IntermediaFileCount()
    
    modProjectFunctions.Set_IntermediaFileCount

End Sub

Sub Init_Pool_Memory()

    modProjectFunctions.Init_Pool_Memory
    
End Sub

Sub Init_Xls_Dll()
    
    modProjectFunctions.SetXlsDll
    
End Sub
Function Run_Pulse_Analize() As Long

    Run_Pulse_Analize = modProjectFunctions.Run_Pulse_Analize
    
End Function

Function Get_Import_Parameters(ByRef IndexFilePwdLst As Long, _
            ByRef IndexFile As Long, _
            ByRef FileCount As Long, _
            ByRef ProccessDone As Long) As Long

    Get_Import_Parameters = modProjectFunctions.Get_Import_Parameters(IndexFilePwdLst, IndexFile, FileCount, ProccessDone)
    
End Function

Sub Set_Parameters_After_Pulse_Analize()

    modProjectFunctions.Set_Parameters_After_Pulse_Analize
    
End Sub

Sub Run_Pulse_Analize_BG()

    modProjectFunctions.Run_Pulse_Analize_BG
    
End Sub

Sub RevertParameters()

    modProjectFunctions.RevertParameters
    
End Sub

Function SaveProject(Optional LV_Force As Boolean = False) As Boolean

    SaveProject = modProjectFunctions.SaveProject(LV_Force)
    
End Function

Function IsClosedProject() As Boolean

    IsClosedProject = modProjectFunctions.IsClosedProject
    
End Function

Function IsEmptyProject() As Boolean

    IsEmptyProject = modProjectFunctions.IsEmptyProject
    
End Function

Function IsNewProject() As Boolean

    IsNewProject = modProjectFunctions.IsNewProject
    
End Function

Sub LoadInfoProject(LV_LstVw As ListView)

    modProjectFunctions.LoadInfoProject LV_LstVw
    
End Sub

Sub LoadProject(ByVal FileName As String, Optional FilePath As String)

Dim h           As Integer

    modProjectFunctions.LoadProject FileName, FilePath

End Sub

Sub LoadWorkSpace()

    modProjectFunctions.LoadWorkSpace
    
End Sub

Sub SetSheetGenerated()

    modProjectFunctions.SetSheetGenerated
    
End Sub

Sub ClearSheetGenerated()

    modProjectFunctions.ClearSheetGenerated
    
End Sub

Function GetSheetGenerated() As Boolean

    GetSheetGenerated = modProjectFunctions.GetSheetGenerated
    
End Function

Sub SetSheetGenerating()

    modProjectFunctions.SetSheetGenerating
    
End Sub

Sub ClearSheetGenerating()

    modProjectFunctions.ClearSheetGenerating
    
End Sub

Function GetSheetGenerating() As Boolean

    GetSheetGenerating = modProjectFunctions.GetSheetGenerating
    
End Function

Sub SetExportQueued()

    modProjectFunctions.SetExportQueued
    
End Sub

Sub ClearExportQueued()

    modProjectFunctions.ClearExportQueued
    
End Sub

Function GetExportQueued() As Boolean

    GetExportQueued = modProjectFunctions.GetExportQueued
    
End Function

Sub SetIntermediateDataReady()

    modProjectFunctions.SetIntermediateDataReady
    
End Sub

Sub ClearIntermediateDataReady()

    modProjectFunctions.ClearIntermediateDataReady
    
End Sub

Function GetIntermediateDataReady() As Boolean

    GetIntermediateDataReady = modProjectFunctions.GetIntermediateDataReady
    
End Function

Sub SetPulsesAnalized()

    modProjectFunctions.SetPulsesAnalized
    
End Sub

Sub ClearPulsesAnalized()

    modProjectFunctions.ClearPulsesAnalized
    
End Sub

Function GetPulsesAnalized() As Boolean

    GetPulsesAnalized = modProjectFunctions.GetPulsesAnalized
    
End Function

Function IsThereProjectList() As Boolean

    IsThereProjectList = modProjectFunctions.IsThereProjectList

End Function

Sub ShowListProjects(LV_LstVw As ListView)

    modProjectFunctions.ShowListProjects LV_LstVw
    
End Sub

Private Sub Class_Initialize()

    modProjectFunctions.Project_Constructor
    
End Sub

Private Sub Class_Terminate()

    modProjectFunctions.Project_Destructor
    
End Sub

