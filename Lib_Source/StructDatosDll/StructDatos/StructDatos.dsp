# Microsoft Developer Studio Project File - Name="StructDatos" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=StructDatos - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "StructDatos.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "StructDatos.mak" CFG="StructDatos - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "StructDatos - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "StructDatos - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "StructDatos - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "STRUCTDATOS_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "STRUCTDATOS_EXPORTS" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x340a /d "NDEBUG"
# ADD RSC /l 0x340a /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386

!ELSEIF  "$(CFG)" == "StructDatos - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "..\..\"
# PROP Intermediate_Dir "Debug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "STRUCTDATOS_EXPORTS" /YX /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "STRUCTDATOS_EXPORTS" /FR /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x340a /d "_DEBUG"
# ADD RSC /l 0x340a /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept

!ENDIF 

# Begin Target

# Name "StructDatos - Win32 Release"
# Name "StructDatos - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\BasicExcel.cpp
# End Source File
# Begin Source File

SOURCE=.\BookMark.cpp
# End Source File
# Begin Source File

SOURCE=.\cFileName.cpp
# End Source File
# Begin Source File

SOURCE=.\cFilesList.cpp
# End Source File
# Begin Source File

SOURCE=.\ControlTable.cpp
# End Source File
# Begin Source File

SOURCE=.\cProjectManager.cpp
# End Source File
# Begin Source File

SOURCE=.\dirent.cpp
# End Source File
# Begin Source File

SOURCE=.\ErrorHandler.cpp
# End Source File
# Begin Source File

SOURCE=.\ExcelFormat.cpp
# End Source File
# Begin Source File

SOURCE=.\filelist.cpp
# End Source File
# Begin Source File

SOURCE=.\filelist2.cpp
# End Source File
# Begin Source File

SOURCE=.\filelist_op.cpp
# End Source File
# Begin Source File

SOURCE=.\gralFuntions.cpp
# End Source File
# Begin Source File

SOURCE=.\MakePulseConvert.bat
# End Source File
# Begin Source File

SOURCE=.\PoolTable.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_excel_export.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_export_dll.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_export_file.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_export_op.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_import_optimize.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_import_optimize2.cpp
# End Source File
# Begin Source File

SOURCE=.\pulse_import_Verify_Errors.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseexport.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseexport_filtered.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseformat.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseimport.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseimport2.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseimport3.cpp
# End Source File
# Begin Source File

SOURCE=.\pulseimport_Destroy.cpp
# End Source File
# Begin Source File

SOURCE=.\Pulseproject.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProject2.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProject_byInterval.cpp
# End Source File
# Begin Source File

SOURCE=.\Pulseproject_ErrorList.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProject_Get.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProjectDestroy.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProjectNew.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProjectSaveWrkSp.cpp
# End Source File
# Begin Source File

SOURCE=.\PulseProjectSpreadSheet.cpp
# End Source File
# Begin Source File

SOURCE=.\ReadFile.cpp
# End Source File
# Begin Source File

SOURCE=.\StructDatos.cpp
# End Source File
# Begin Source File

SOURCE=.\TpoReal.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\BasicExcel.hpp
# End Source File
# Begin Source File

SOURCE=.\BookMark.h
# End Source File
# Begin Source File

SOURCE=.\cFileName.h
# End Source File
# Begin Source File

SOURCE=.\cFilesList.h
# End Source File
# Begin Source File

SOURCE=.\ControlTable.h
# End Source File
# Begin Source File

SOURCE=.\cProjectManager.h
# End Source File
# Begin Source File

SOURCE=.\dirent.h
# End Source File
# Begin Source File

SOURCE=.\ErrorHandler.h
# End Source File
# Begin Source File

SOURCE=.\ExcelFormat.h
# End Source File
# Begin Source File

SOURCE=.\filelist.h
# End Source File
# Begin Source File

SOURCE=.\gralFunctions.h
# End Source File
# Begin Source File

SOURCE=.\ListToArray.h
# End Source File
# Begin Source File

SOURCE=.\PoolMemory.h
# End Source File
# Begin Source File

SOURCE=.\PoolTable.h
# End Source File
# Begin Source File

SOURCE=.\pulse_conv_struct_define.h
# End Source File
# Begin Source File

SOURCE=.\pulse_excel_export.h
# End Source File
# Begin Source File

SOURCE=.\pulseexport.h
# End Source File
# Begin Source File

SOURCE=.\pulseformat.h
# End Source File
# Begin Source File

SOURCE=.\pulseimport.h
# End Source File
# Begin Source File

SOURCE=.\Pulseproject.h
# End Source File
# Begin Source File

SOURCE=.\ReadFile.h
# End Source File
# Begin Source File

SOURCE=.\StructDatos.def
# End Source File
# Begin Source File

SOURCE=.\StructDatos.odl
# End Source File
# Begin Source File

SOURCE=.\TpoReal.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\StructDatos.rc
# End Source File
# End Group
# End Target
# End Project
