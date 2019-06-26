Attribute VB_Name = "modStructDatos"
'---------------------------------------------------------------------------------------
' Module    : modStructDatos
' Author    : Leo Herrera
' Date      : 06/04/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Declare Function Conv_Byte_2_Str Lib "..\Lib\GralFunctions.dll" _
                    (ByRef hBuffer As Byte, _
                    ByVal lCount As Long, _
                    ByRef sStr As String)
                    
Declare Function Conv_Str_2_IdEvent Lib "..\Lib\GralFunctions.dll" _
                    (ByVal lsEvent As String, ByVal liDigits As Byte) As Long
''
Declare Function Pulse_Import_File Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal StrPath As String) As Long
            
Declare Sub Pulse_Export_File Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal StrPath As String)
            
Declare Sub Pulse_Get_File Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexSpread As Integer, _
            ByVal IndexSheet As Integer, _
            ByRef StrFileName As String)
            
Declare Sub Pulse_Field_Header Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexFile As Integer, _
            ByRef StrFileName As String)
            
            
Declare Sub Pulse_Get_Pwd Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexSheet As Integer, _
            ByVal IndexPulse As Long, _
            ByRef Pwd As Double)
            
Declare Sub Pulse_GetPwd Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexFile As Long, _
            ByVal IndexSheet As Long, _
            ByVal IndexPulse As Long, _
            ByRef Pwd As Double)
            
Declare Function Pulse_Files_Count Lib _
            "..\Lib_Source\StructDatos.dll" _
            () As Integer
            
Declare Function Pulse_Field_Count Lib _
            "..\Lib_Source\StructDatos.dll" _
            () As Integer
            
Declare Function Pulse_Count Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexSheet As Integer) As Long
            
Declare Function Pulse_GetSheetPulseCount Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexFile As Long, _
            ByVal IndexSheet As Long) As Long
            
Declare Function Pulse_SpreadSheet_Saved Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef lvDone As Boolean) As Long

Declare Function Pulse_SpreadSheetDone Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef lvDone As Boolean) As Long

Declare Sub Pulse_SaveAsStart Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal FileName As String)

Declare Sub Pulse_SpreadSheet_SaveStatus Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef IndexFile As Integer, _
            ByRef PulseQty As Long)

Declare Sub Pulse_SpreadSheetStatus Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef IndexFile As Long, _
            ByRef PulseQty As Long)

Declare Sub Pulse_Debug Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef IndexFile As Integer, _
            ByRef PulseQty As Long, _
            ByVal PathName As String)

Declare Sub Pulse_Sheets_Per_File Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal Sheets As Long)
            
Declare Sub Pulse_Sheets_Per_Pulses Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal Pulses As Long)

Declare Sub Pulse_Sheets_Per_Interval Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal Interval As Double)

Declare Sub Pulse_CreateSheet Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal SpreadIndex As Long, ByVal SheetIndex As Long)

Declare Sub Pulse_SaveSpreadSheet Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal SpreadIndex As Long)

Declare Sub Pulse_SetWorkSpacePath Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal PathName As String)

Declare Sub Pulse_OutputPath Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal PathName As String)

Declare Sub Pulse_Create_Xls_File_Op Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal FileName As String, _
            ByVal SpreadIndex As Long)

Declare Sub Pulse_Finish_Xls Lib _
            "..\Lib_Source\StructDatos.dll" ()
            
Declare Sub Pulse_Create_Xls_File Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal FileName As String, _
            ByVal SpreadIndex As Long, _
            ByVal GenBinEnable As Boolean)
            
Declare Function Pulse_GetSpreadFileCount Lib _
            "..\Lib_Source\StructDatos.dll" _
            () As Long
            
Declare Function Pulse_GetSheetCount Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal SpreadIndex As Long) As Long
            
Declare Sub Pulse_CreateWorkSpace Lib _
            "..\Lib_Source\StructDatos.dll" ()

Declare Sub Pulse_FilesPerWorkSpace Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal FileCount As Long)

Declare Sub Pulse_Destroy_All Lib _
            "..\Lib_Source\StructDatos.dll" _
            ()

'/-------------------------------------------
'/-         Get Info
'/-------------------------------------------
Declare Sub Pulse_GetProjectInfo Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef PulseQty As Long, _
            ByRef TimeIni As Double, _
            ByRef TimeEnd As Double)

Declare Sub Pulse_GetSpreadFileInfo Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexFile As Long, _
            ByRef PulseQty As Long, _
            ByRef TimeIni As Double, _
            ByRef TimeEnd As Double)
            
Declare Sub Pulse_GetSheetInfo Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal IndexFile As Long, _
            ByVal SheetIndex As Long, _
            ByRef PulseQty As Long, _
            ByRef TimeIni As Double, _
            ByRef TimeEnd As Double)
            
Declare Sub Pulse_GetMissionInfo Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal lvPath As String, _
            ByRef FileCount As Long, _
            ByRef TimeIni As String, _
            ByRef TimeEnd As String)

Declare Sub Pulse_SetFieldFormat Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef ColumnOrder As Long, _
            ByRef ColumnView As Long)


Declare Sub Pulse_SetIntermediaFileCount Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal lvCount As Long)



Declare Sub Pulse_Import_File_BG Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal StrPath As String)

Declare Sub Pulse_CreateWorkSpace_BG Lib _
            "..\Lib_Source\StructDatos.dll" ()

Declare Sub Pulse_Import_File_Status Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef IndexFilePwdLst As Long, _
            ByRef IndexFile As Long, _
            ByRef FileCount As Long, _
            ByRef ProccessDone As Long)

Declare Sub Pulse_LoadWorkSpace Lib _
            "..\Lib_Source\StructDatos.dll" ()

Declare Sub Pulse_DestroyWorkSpace Lib _
            "..\Lib_Source\StructDatos.dll" ()

Declare Function Pulse_CancelXlsProcess Lib _
            "..\Lib_Source\StructDatos.dll" () As Boolean

Declare Function Pulse_GetErrorListCount Lib _
            "..\Lib_Source\StructDatos.dll" () As Long
            
Declare Sub Pulse_GetErrorFieldCount Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef CountLong As Long, _
            ByRef CountDouble As Long)
            
Declare Sub Pulse_GetErrorFieldHeader Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal Index As Long, ByRef lsHeader As String)
            
Declare Sub Pulse_GetErrorPointer Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal Index As Long, _
            ByRef lArray As Long, _
            ByRef dArray As Double)

' Pool Memory
Declare Sub Pulse_Init_PoolMemory Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal PoolFilePwdListSize As Long, _
            ByVal PoolFilePwdSize As Long, _
            ByVal PoolPwdNFSize As Long)
            
Declare Sub Pulse_GetStructSize Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByRef FilePwdListSize As Long, _
            ByRef FilePwdSize As Long, _
            ByRef PwdNFSize As Long, _
            ByRef PwdSize As Long)
            
Declare Sub Pulse_Destroy_PoolMemory Lib _
            "..\Lib_Source\StructDatos.dll" _
            ()
            
Declare Sub Pulse_Set_Xls_Dll Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal StrDll As String, ByVal StrBookConstructor As String, _
            ByVal StrBookDestructor As String, _
            ByVal StrBookSave As String, ByVal StrBookSetHeader As String, _
            ByVal StrBookSetSheet As String, ByVal StrBookSetOrder As String, _
            ByVal StrBookCvtBin As String)
            

Declare Function CvtBinXls Lib "..\Lib_Source\cvt2xls.dll" _
            Alias "_Z11Book_CvtBinPc" (ByVal SheetCount As Long, _
            ByVal BinFileName As String) As String


Declare Sub Pulse_CvtBinXls Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal SheetCount As Long, ByVal BinFileName As String)

Declare Sub Pulse_Log_Path Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal StrPath As String)

Declare Sub Pulse_Close_Log Lib _
            "..\Lib_Source\StructDatos.dll" _
            ()



Declare Sub Test_Proccess_Map_File Lib _
            "..\Lib_Source\StructDatos.dll" _
            (ByVal StrPath As String)



