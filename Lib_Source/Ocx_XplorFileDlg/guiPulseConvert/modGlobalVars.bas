Attribute VB_Name = "modGlobalVars"
'---------------------------------------------------------------------------------------
' Module    : modGlobalVars
' Author    : lherrera
' Date      : 19/01/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public GV_Mdi                   As MDIMain
Public GV_Pulse_Path            As String
Public GV_MissionName           As String
Public GV_Config_Path           As String
Public GV_XlsDll_Path           As String
Public GV_Bin2Xls_App           As String
Public GV_ListProjectFile       As String
Public GV_Project_Opened        As Boolean
Public GV_Project_Closed        As Boolean
Public GV_Restore_Project_Mnu   As Boolean
Public GV_Send_Progress         As Boolean
Public GV_PrevInstance          As Boolean
Public GV_Index_Pjt             As Integer
Public GV_Finalizar_App         As Boolean
Public GV_XlsProcessCanceled    As Boolean

Public GV_clsErrorView                  As clsInstanceControl
Public GV_clsTemplateConfigSpreadSheet  As clsInstanceControl
Public GV_clsSecondInstance             As clsInstanceControl
Public GV_clsOpenPulseDialog            As clsInstanceControl
Public GV_clsSpreadSheetView            As clsInstanceControl
Public GV_clsExportSpreadSheet          As clsInstanceControl
Public GV_clsConfigSpreadSheet          As clsInstanceControl
Public GV_clsStart                      As clsInstanceControl
Public GV_clsStartUp                    As clsInstanceControl
Public GV_clsProjectSelLocation         As clsInstanceControl
Public GV_clsPjtFromFile                As clsInstanceControl
Public GV_clsPjtFromList                As clsInstanceControl

Public GV_WorkSpace                     As String
Public GV_Output                        As String

Public GV_Msg_Len()                     As Long
Public GV_Msg_Header()                  As String

Public GV_TimeToKillSecond              As Long
Public GV_TimeToKillMySelf              As Long
Public GV_TimeOutComm                   As Long
Public GV_TimeOutSendAlive              As Long

Public GV_ActualColumnConfig            As typeConfigSheetColumns

Enum ProjectState
    NewProject = 1
    
End Enum

Public Enum mnuProjectConstant
    MnuEmpty = 0
    SelPathProject = 1
    SelPathMission
    ConfigOutput
    PreviewOutput
    GenOutput
    VerifyErrors
End Enum

Type typeDebug
    Template        As Boolean
End Type


Type ProjectInfo
    ProjectName     As String
    ProjectPath     As String
End Type

Type ProjectList
    Count                   As Integer
    List()                  As ProjectInfo
End Type

Type typeConfigColumn
    ColumnName              As String
    Visible                 As Boolean
    Order                   As Integer
End Type

Type typeConfigSheetColumns
    ColumnConfigName        As String
    Count                   As Integer
    Column()                As typeConfigColumn
End Type

Type typeColumnsConfigList
    Count                   As Integer
    Config()                As typeConfigSheetColumns
End Type

Type typeConfigSpreadSheet
    SpreadConfigName        As String
    byFiles                 As Boolean
    byPulses                As Boolean
    byInterval              As Boolean
    SheetsPerSpreadSheet    As Long
    PulsesPerSheet          As Long
    IntervalPerSheet        As Long
End Type

Type typeConfigSpreadSheetList
    Count                   As Integer
    ConfigList()            As typeConfigSpreadSheet
End Type

Type typeSpreadSheet
    PulseConfig             As typeConfigSpreadSheet
    ColumnConfig            As typeConfigSheetColumns
End Type
    
Public Type Project_Old_Struct
    ProjectName     As String
    ProjectPath     As String
    ProjectFileName As String
    OutputPath      As String
    WorkSpacePath   As String
    PulsePath       As String
    MissionPath     As String
    MissionDate     As String
    MissionName     As String
    PulseIniTime    As String
    PulseEndTime    As String
    PulseFileCount  As Long
    IntermediaFileCount     As Long
    
    ProjectEmpty    As Boolean
    ProjectClosed   As Boolean
    NewProject      As Boolean
    Changed         As Boolean
    Saved           As Boolean
    SheetConfigured As Boolean
    PulsesAnalized  As Boolean
    IntermediateDataReady   As Boolean
    SheetGenerated  As Boolean
    
    ProjectFolderSelected   As Boolean
    MissionSelected         As Boolean
    
    AsociaMissionName   As Integer
    CreateFolderProject As Integer
    OutputSubFolder     As Integer
    WrkSpcSubFolder     As Integer
    
    ColumnConfig        As typeConfigSheetColumns
    SpreadConfig        As typeConfigSpreadSheet
    
    'Processing Parameters
    FilesPerWorkSpace   As Long
End Type

Public Type Project
    PulseFileCount          As Long
    IntermediaFileCount     As Long
    FilesPerWorkSpace       As Long
    Spare_Long              As Long
    
    AsociaMissionName       As Integer
    CreateFolderProject     As Integer
    OutputSubFolder         As Integer
    WrkSpcSubFolder         As Integer
    
    ProjectEmpty            As Boolean
    ProjectClosed           As Boolean
    NewProject              As Boolean
    Changed                 As Boolean
    
    Saved                   As Boolean
    SheetConfigured         As Boolean
    PulsesAnalized          As Boolean
    IntermediateDataReady   As Boolean
    
    SheetGenerated          As Boolean
    SheetGenerating         As Boolean
    ExportQueued            As Boolean
    ProjectFolderSelected   As Boolean
    
    MissionSelected         As Boolean
    Spare_Boolean1          As Boolean
    Spare_Boolean2          As Boolean
    Spare_Boolean3          As Boolean
    
    
    ProjectName             As String
    ProjectPath             As String
    ProjectFileName         As String
    OutputPath              As String
    WorkSpacePath           As String
    PulsePath               As String
    MissionPath             As String
    MissionDate             As String
    MissionName             As String
    PulseIniTime            As String
    PulseEndTime            As String
    
    ColumnConfig            As typeConfigSheetColumns
    SpreadConfig            As typeConfigSpreadSheet
    
End Type


Public Type RemoteProject
    WorkSpacePath       As String
    OutputPath          As String
    ProjectName         As String
    ProjectPath         As String
    FileName            As String
    IndexPulse          As Long
    IndexSheet          As Long
    IndexSpread         As Long
    IndexPulseMax       As Long
    IndexSheetMax       As Long
    IndexSpreadMax      As Long
    TimeIni             As Date
    TimeEnd             As Date
    TickIni             As Long
    TickEnd             As Long
    OutFiles()          As String
    OutFilesSheetCount() As Long
    OutFilesCount       As Long
    GenerationDone      As Boolean
    ProjectStarted      As Boolean
    Processing          As Boolean
    Status              As Boolean
    ProjectState        As Message_Header_Const
    IndexTrVw           As Long
    hId                 As Double
End Type

Public Type RemoteProjectList
    ProjectList()       As RemoteProject
    Count               As Long
    ListIndex           As Long
End Type

Public Type typeXlsConvertList
    XlsFile             As String
    XlsMark             As String
    Converted           As Boolean
End Type

Public XlsConvertList()     As typeXlsConvertList

Public BackGroundProjectList            As RemoteProjectList

Public m_Project                        As ClassProject
Public m_MakeRound                      As clsMakeRound

'Public GV_OpenPulseDialog_Loaded        As Boolean
'Public GV_SpreadSheetView_Loaded        As Boolean
'Public GV_ExportSpreadSheet_Loaded      As Boolean
'Public GV_ConfigSpreadSheet_Loaded      As Boolean
'Public GV_Start_Loaded                  As Boolean
'Public GV_ProjectSelLocation_Loaded     As Boolean
'
'Public GV_Load_OpenPulseDialog          As Boolean
'Public GV_Load_SpreadSheetView          As Boolean
'Public GV_Load_ExportSpreadSheet        As Boolean
'Public GV_Load_ConfigSpreadSheet        As Boolean
'Public GV_Load_Start                    As Boolean
'Public GV_Load_ProjectSelLocation       As Boolean

Public GV_Debug     As typeDebug


