Attribute VB_Name = "modGlobalConstants"
'---------------------------------------------------------------------------------------
' Module    : modGlobalConstants
' Author    : Leo Herrera
' Date      : 16/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Const GC_RELEASE_DATE = "11 de Octubre del 2013"
Public Const GC_NORMALES = "\NORMALES"
Public Const GC_PROJECT_EXTENSION = ".pcp"

Public Const GC_CONFIGURATION_SECTION = "Configuration"
Public Const GC_MEMORY_SECTION = "Memory Configuration"

Public Const GC_START_UP_DISABLE = "StartUp Disable"
Public Const GC_CONFIG_PATH = "Config"
Public Const GC_DLL_PATH = "Lib_Source"
Public Const GC_MAX_COUNT_LIST_PROJECT = 32
Public Const GC_DEBUG = "Debug"
Public Const GC_ERASE_DEBUG_SETTING = "Erase_Debug_Setting"
Public Const GC_ENABLE_RUN_CVTXLS = "Enable_Cvt2Xls"
Public Const GC_ENABLE_HIDE_CVTXLS = "Enable_Hide_Cvt2Xls"
Public Const GC_VISIBLE_SND_INSTANCE = "Visible_Snd_Instance"
Public Const GC_ENABLE_EXPORTFILE_LOG = "Enable_ExportFile_Log"
Public Const GC_ENABLE_DBG_FILE_LOG = "Enable_Dbg_File_Log"
Public Const GC_ENABLE_DBG_DLL_FILE_LOG = "Enable_Dbg_Dll_File_Log"
Public Const GC_INVERSE_PREV_INSTANCE = "Enable_Inverse_Prev_Instance_Value"
Public Const GC_BIN_2_XLS_APP = "Bin2Xls_App_Name"
Public Const GC_ENABLE_XLS_OP = "Enable_Xls_Optimized"

Public Const GC_MSG_RUNXLS = "RUNXLS"
Public Const GC_MSG_ALIVE = "ALIVE"
Public Const GC_MSG_START_PROJECT = "START_PROJECT"
Public Const GC_MSG_FILE_START = "FILE_START"
Public Const GC_MSG_STATUS = "STATUS"
Public Const GC_MSG_SAVING_FILE = "SAVING_FILE"
Public Const GC_MSG_XLS_FILE_READY = "XLS_FILE_READY"
Public Const GC_MSG_END_PROJECT = "END_PROJECT"

Public Enum Message_Header_Const
    MSG_RUNXLS = 0
    MSG_ALIVE
    MSG_START_PROJECT
    MSG_FILE_START
    MSG_STATUS
    MSG_SAVING_FILE
    MSG_XLS_FILE_READY
    MSG_END_PROJECT
    MSG_ERROR = 99
End Enum



'---
