Attribute VB_Name = "modSelDirectorio"
'
'//////////////////////////////////////////////////////////////////////////////
'/////               ESTE C�DIGO INSERTALO EN UN M�DULO BAS               /////
'//////////////////////////////////////////////////////////////////////////////
'
'------------------------------------------------------------------------------
' M�dulo con las declaraciones y funciones para BrowseForFolder     (12/May/99)
'
' �Guillermo 'guille' Som, 1999
'------------------------------------------------------------------------------
Option Explicit

'//////////////////////////////////////////////////////////////////////////////
' Variables, constantes y funciones para usar con BrowseForFolder   (12/May/99)
'//////////////////////////////////////////////////////////////////////////////
'
Private sFolderIni As String
'
Private Const WM_USER = &H400&
Public Const MAX_PATH = 260&
'
' Tipo para usar con SHBrowseForFolder
Private Type BrowseInfo
    hWndOwner               As Long             ' hWnd del formulario
    pidlRoot                As Long             ' Especifica el pID de la carpeta inicial
    pszDisplayName          As String           ' Nombre del item seleccionado
    lpszTitle               As String           ' T�tulo a mostrar encima del �rbol
    ulFlags                 As Long             '
    lpfnCallback            As Long             ' Funci�n CallBack
    lParam                  As Long             ' Informaci�n extra a pasar a la funci�n Callback
    iImage                  As Long             '
End Type

Public Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'
'// Browsing for directory.
Public Const BIF_RETURNONLYFSDIRS = &H1&       '// For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2&      '// For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4&
Public Const BIF_RETURNFSANCESTORS = &H8&
Public Const BIF_EDITBOX = &H10&
Public Const BIF_VALIDATE = &H20&              '// insist on valid result (or CANCEL)
'
Public Const BIF_BROWSEFORCOMPUTER = &H1000&   '// Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000&    '// Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000&  '// Browsing for Everything
'
'// message from browser
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_VALIDATEFAILED = 3          '// lParam:szPath ret:1(cont),0(EndDialog)
'Public Const BFFM_VALIDATEFAILEDW = 4&         '// lParam:wzPath ret:1(cont),0(EndDialog)
'
'// messages to browser
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BFFM_ENABLEOK = (WM_USER + 101)
Public Const BFFM_SETSELECTION = (WM_USER + 102)
'Public Const BFFM_SETSELECTIONW = (WM_USER + 103&)
'Public Const BFFM_SETSTATUSTEXTW = (WM_USER + 104&)
'
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

Public Const LPTR = (&H0 Or &H40)


'Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
        (lpbi As BrowseInfo) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
    "SHBrowseForFolderA" ( _
    lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
'
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
        (ByVal hMem As Long)
'
'Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
'        (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Public Declare Function LocalAlloc Lib "kernel32" ( _
    ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long

Public Declare Function LocalFree Lib "kernel32" ( _
    ByVal hMem As Long) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
    "SHGetPathFromIDListA" ( _
    ByVal pidl As Long, _
    ByVal pszPath As String) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" ( _
    pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)


Public Function FunctionPointer(FunctionAddress As Long) As Long
    FunctionPointer = FunctionAddress
End Function

'Public Function BrowseFolderCallbackProc(ByVal hWndOwner As Long, _
'                                        ByVal uMsg As Long, _
'                                        ByVal lParam As Long, _
'                                        ByVal pData As Long) As Long
'    ' Llamada CallBack para usar con la funci�n BrowseForFolder     (12/May/99)
'    Dim szDir As String
'
'    On Local Error Resume Next
'
'    Select Case uMsg
'    '--------------------------------------------------------------------------
'    ' Este mensaje se enviar� cuando se inicia el di�logo,
'    ' entonces es cuando hay que indicar el directorio de inicio.
'    Case BFFM_INITIALIZED
'        ' El path de inicio ser� el directorio indicado,
'        ' si no se ha asignado, usar el directorio actual
'        If Len(sFolderIni) Then
'            szDir = sFolderIni & Chr$(0)
'        Else
'            szDir = CurDir$ & Chr$(0)
'        End If
'        ' WParam  ser� TRUE  si se especifica un path.
'        '         ser� FALSE si se especifica un pIDL.
'        Call SendMessage(hWndOwner, BFFM_SETSELECTION, 1&, ByVal szDir)
'    '--------------------------------------------------------------------------
'    ' Este mensaje se produce cuando se cambia el directorio
'    ' Si nuestro form est� subclasificado para recibir mensajes,
'    ' puede interceptar el mensaje BFFM_SETSTATUSTEXT
'    ' para mostrar el directorio que se est� seleccionando.
'    Case BFFM_SELCHANGED
'        szDir = String$(MAX_PATH, 0)
'        ' Notifica a la ventana del directorio actualmente seleccionado,
'        ' (al menos en teor�a, ya que no lo hace...)
'        If SHGetPathFromIDList(lParam, szDir) Then
'            Call SendMessage(hWndOwner, BFFM_SETSTATUSTEXT, 0&, ByVal szDir)
'        End If
'        Call CoTaskMemFree(lParam)
'    End Select
'
'    Err = 0
'    BrowseFolderCallbackProc = 0
'
''------------------------------------------------------------------------------
'' Este es el c�digo de C en el que est� basada esta funci�n Callback
'' C�digo obtenido de la MSDN Library de Microsoft:
'' HOWTO: Browse for Folders from the Current Directory
'' Article ID: Q179378
''
''         TCHAR szDir[MAX_PATH];
''
''         switch(uMsg) {
''            case BFFM_INITIALIZED: {
''               if GetCurrentDirectory(sizeof(szDir)/sizeof(TCHAR),
''                                      szDir)) {
''                  // WParam is TRUE since you are passing a path.
''                  // It would be FALSE if you were passing a pidl.
''                  SendMessage(hwnd,BFFM_SETSELECTION,TRUE,(LPARAM)szDir);
''               }
''               break;
''            }
''            case BFFM_SELCHANGED: {
''               // Set the status window to the currently selected path.
''               if (SHGetPathFromIDList((LPITEMIDLIST) lp ,szDir)) {
''                  SendMessage(hwnd,BFFM_SETSTATUSTEXT,0,(LPARAM)szDir);
''               }
''               break;
''            }
''           default:
''               break;
''         }
''         return 0;
''------------------------------------------------------------------------------
'End Function
'
Public Function BrowseCallbackProcStr(ByVal hwnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
    If uMsg = 1 Then
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If
End Function

Public Function rtnAddressOf(lngProc As Long) As Long
    ' Devuelve la direcci�n pasada como par�metro
    ' Esto se usar� para asignar a una variable la direcci�n de una funci�n
    ' o procedimiento.
    ' Por ejemplo, si en un tipo definido se asigna a una variable la direcci�n
    ' de una funci�n o procedimiento
    rtnAddressOf = lngProc
End Function


'Public Function BrowseForFolder_1(ByVal hWndOwner As Long, ByVal sPrompt As String, _
'                Optional sInitDir As String = "", _
'                Optional ByVal lFlags As Long = BIF_RETURNONLYFSDIRS) As String
'    ' Muestra el di�logo de selecci�n de directorios de Windows
'    ' Si todo va bien, devuelve el directorio seleccionado
'    ' Si se cancela, se devuelve una cadena vac�a y se produce el error 32755
'    '
'    ' Los par�metros de entrada:
'    '   El hWnd de la ventana
'    '   El t�tulo a mostrar
'    '   Opcionalmente el directorio de inicio
'    '   En lFlags se puede especificar lo que se podr� seleccionar:
'    '       BIF_BROWSEINCLUDEFILES, etc.
'    '       por defecto es: BIF_RETURNONLYFSDIRS
'    '
'    Dim iNull As Integer
'    Dim lpIDList As Long
'    Dim lResult As Long
'    Dim sPath As String
'    Dim udtBI As BrowseInfo
'
'    On Local Error Resume Next
'
'    With udtBI
'        .hWndOwner = hWndOwner
'        ' T�tulo a mostrar encima del �rbol de selecci�n
'        .lpszTitle = sPrompt & vbNullChar
'        ' Que es lo que debe devolver esta funci�n
'        .ulFlags = lFlags
'        '.ulFlags = lFlags Or BIF_RETURNONLYFSDIRS
'        '
'        ' Si se especifica el directorio por el que se empezar�...
'        If Len(sInitDir) Then
'            ' Asignar la variable que contendr� el directorio de inicio
'            sFolderIni = sInitDir
'            ' Indicar la funci�n Callback a usar.
'            ' Como hay que asignar esa direcci�n a una variable,
'            ' se usa una funci�n "intermedia" que devuelve el valor
'            ' del par�metro pasado... es decir: �la direcci�n de la funci�n!
'            .lpfnCallback = rtnAddressOf(AddressOf BrowseFolderCallbackProc)
'        End If
'    End With
'    Err = 0
'    On Local Error GoTo 0
'
'    ' Mostramos el cuadro de di�logo
'    lpIDList = SHBrowseForFolder_1(udtBI)
'    '
'    If lpIDList Then
'        ' Si se ha seleccionado un directorio...
'        '
'        ' Obtener el path
'        sPath = String$(MAX_PATH, 0)
'        lResult = SHGetPathFromIDList(lpIDList, sPath)
'        Call CoTaskMemFree(lpIDList)
'        ' Quitar los caracteres nulos del final
'        iNull = InStr(sPath, vbNullChar)
'        If iNull Then
'            sPath = Left$(sPath, iNull - 1)
'        End If
'    Else
'        ' Si se ha pulsado en cancelar...
'        '
'        ' Devolver una cadena vac�a y asignar un error
'        sPath = ""
'        With Err
'            .Source = "MBrowseFolder::BrowseForFolder"
'            .Number = 32755
'            .Description = "Cancelada la operaci�n de BrowseForFolder"
'        End With
'    End If
'
'    BrowseForFolder_1 = sPath
'End Function

Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal sPrompt As String, _
                Optional ByRef PathInicial As String = "", _
                Optional ByVal lFlags As Long = BIF_RETURNONLYFSDIRS) _
                As String

Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
    
    Dim CheckPath As String
    CheckPath = PathInicial
    If Len(CheckPath) > 0 Then
        If Not Right$(CheckPath, 1) <> "\" Then
            CheckPath = Left$(CheckPath, Len(CheckPath) - 1)
        End If
    End If
    PathInicial = CheckPath
    
    With Browse_for_folder
        .hOwner = hWndOwner
        ' t�tulo del di�logo
        .lpszTitle = sPrompt & vbNullChar
        ' Que es lo que debe devolver esta funci�n
        '.ulFlags = lFlags
         ' CallBack Function
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr)
        selectedPathPointer = LocalAlloc(LPTR, Len(PathInicial) + 1)
        CopyMemory ByVal selectedPathPointer, ByVal PathInicial, Len(PathInicial) + 1
        .lParam = selectedPathPointer
    End With
    
    itemID = SHBrowseForFolder(Browse_for_folder)
    If itemID Then
        If SHGetPathFromIDList(itemID, tmpPath) Then
            BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(itemID)
    End If
    Call LocalFree(selectedPathPointer)
End Function







