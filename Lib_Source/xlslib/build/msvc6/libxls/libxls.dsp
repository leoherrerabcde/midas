# Microsoft Developer Studio Project File - Name="libxls" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=libxls - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "libxls.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "libxls.mak" CFG="libxls - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "libxls - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "libxls - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "libxls - Win32 Release"

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
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "LIBXLS_EXPORTS" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "LIBXLS_EXPORTS" /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0xc0a /d "NDEBUG"
# ADD RSC /l 0xc0a /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386

!ELSEIF  "$(CFG)" == "libxls - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "LIBXLS_EXPORTS" /Yu"stdafx.h" /FD /GZ  /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "LIBXLS_EXPORTS" /Yu"stdafx.h" /FD /GZ  /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0xc0a /d "_DEBUG"
# ADD RSC /l 0xc0a /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept

!ENDIF 

# Begin Target

# Name "libxls - Win32 Release"
# Name "libxls - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=..\..\..\src\xlslib\assert_assist.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\binfile.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\blank.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\boolean.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\cbridge.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\cell.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\colinfo.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\colors.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\continue.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\datast.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\docsumminfo.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\err.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\extformat.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\font.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\format.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\formula.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\globalrec.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\HPSF.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\index.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\label.cpp
# End Source File
# Begin Source File

SOURCE=.\libxls.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\merged.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\note.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\number.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\oledoc.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\olefs.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\oleprop.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\overnew.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\range.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\recdef.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\record.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\row.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\sheetrec.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\summinfo.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\unit.cpp
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\workbook.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Group "build"

# PROP Default_Filter ""
# End Group
# Begin Group "src"

# PROP Default_Filter "*.cpp"
# Begin Group "common"

# PROP Default_Filter "*.cpp"
# Begin Source File

SOURCE=..\..\..\src\common\overnew.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\stringtok.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\systype.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\timespan.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\xls_poppack.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\xls_pshpack1.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\xls_pshpack2.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\xls_pshpack4.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\xlstypes.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\common\xlsys.h
# End Source File
# End Group
# Begin Group "oledoc"

# PROP Default_Filter "*.cpp"
# Begin Source File

SOURCE=..\..\..\src\oledoc\binfile.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\oledoc.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\olefs.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\oledoc\oleprop.h
# End Source File
# End Group
# Begin Group "xlslib"

# PROP Default_Filter "*.cpp"
# Begin Source File

SOURCE=..\..\..\src\xlslib\biffsection.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\blank.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\boolean.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\cbridge.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\cell.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\colinfo.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\colors.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\common.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\continue.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\datast.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\docsumminfo.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\err.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\extformat.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\font.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\format.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\formtags.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\formula.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\globalrec.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\HPSF.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\index.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\label.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\merged.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\note.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\number.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\range.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\recdef.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\record.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\rectypes.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\row.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\sheetrec.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\summinfo.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\tostr.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\unit.h
# End Source File
# Begin Source File

SOURCE=..\..\..\src\xlslib\workbook.h
# End Source File
# End Group
# Begin Source File

SOURCE=..\..\..\src\xlslib.h
# End Source File
# End Group
# Begin Source File

SOURCE=.\libxls.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
