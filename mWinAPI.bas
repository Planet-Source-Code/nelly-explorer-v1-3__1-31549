Attribute VB_Name = "mWinAPI"
Public lSendMsgRet(1) As Long
Public lListItemCounter(1) As Long

'***********************************************************************************
'API used to retrieve the ComputerName of the current local system.
Public Declare Function GetComputerName Lib "kernel32" Alias _
"GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'***********************************************************************************

'***********************************************************************************
'API`s used to retrieve Folders, Files and Information.
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
(ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260
Private Const AMAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA                  'Public Type
    dwFileAttributes As Long                 'Specifies the file attributes of the file found.
    ftCreationTime As FILETIME               'Specifies a FILETIME structure containing the time the file was created.
    ftLastAccessTime As FILETIME             'Specifies a FILETIME structure containing the time that the file was last accessed.
    ftLastWriteTime As FILETIME              'Specifies a FILETIME structure containing the time that the file was last written to.
    nFileSizeHigh As Long                    'Specifies the high-order DWORD value of the file size, in bytes.
    nFileSizeLow As Long                     'Specifies the low-order DWORD value of the file size, in bytes.
    dwReserved0 As Long                      'If the dwFileAttributes member includes the FILE_ATTRIBUTE_REPARSE_POINT attribute, this member specifies the reparse tag. Otherwise, this value is undefined and should not be used.
    dwReserved1 As Long                      'Reserved for future use.
    cFileName As String * MAX_PATH           'A null-terminated string that is the name of the file.
    cAlternateFileName As String * AMAX_PATH 'A null-terminated string that is an alternative name for the file.
End Type

Public Enum FILE_ATTRIBUTES           'Self explanitary
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum
'***********************************************************************************

'***********************************************************************************
'The FileTimeToSystemTime function converts a 64-bit file time to system time format.
Public Declare Function FileTimeToSystemTime Lib "kernel32" _
(lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
'***********************************************************************************

'***********************************************************************************
'The GetLogicalDriveStrings function fills a buffer with strings that specify valid
'drives in the system.
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
"GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
ByVal lpBuffer As String) As Long
'***********************************************************************************


'***********************************************************************************
'The GetDriveType function determines whether a disk drive is a removable, fixed,
'CD-ROM, RAM disk, or network drive.
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
(ByVal nDrive As String) As Long
'***********************************************************************************

'***********************************************************************************
'The GetVolumeInformation function returns information about a file system and
'volume whose root directory is specified.
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, _
ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'***********************************************************************************


'***********************************************************************************
'ShellExecute (Opens or prints a specified File.)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWDEFAULT = 10
'***********************************************************************************

'***********************************************************************************
'The CopyMemory function copies a block of memory from one location to another.
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, _
pSrc As Any, ByVal ByteLen As Long)
'***********************************************************************************

'***********************************************************************************
'Tells us how long a string is.
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'***********************************************************************************

'***********************************************************************************
'Finds the extension of a File. eg. "c:\temp.txt" = ".txt"
Public Declare Function PathFindExtension Lib "Shlwapi" Alias "PathFindExtensionW" _
(ByVal pPath As Long) As Long
'***********************************************************************************

'***********************************************************************************
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
(ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const ILD_TRANSPARENT = &H1
Public Const ILD_NORMAL = &H0

Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000

Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, _
ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal _
fStyle As Long) As Long

Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
'***********************************************************************************

'***********************************************************************************
'The SendMessage function sends the specified message to a window or windows.
'The function calls the window procedure for the specified window and does not
'return until the window procedure has processed the message.
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) _
As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Public Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
'***********************************************************************************
                                                                                    
'***********************************************************************************
'The SetTimer function creates a timer with the specified time-out value.
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, _
ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'***********************************************************************************

'***********************************************************************************
'The KillTimer function destroys the specified timer.
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) _
As Long
'***********************************************************************************

'***********************************************************************************
'The TimerProc function is an application-defined callback function that processes
'WM_TIMER messages. The TIMERPROC type defines a pointer to this callback function.
'TimerProc is a placeholder for the application-defined function name.
Public Sub TIMERPROC(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwtime As Long)
    
    lSendMsgRet&(0) = SendMessage(frmMain.FileList.hwnd, LVM_GETTOPINDEX, 0&, 0&)

    If lSendMsgRet&(0) <> lSendMsgRet&(1) Then
        
        lSendMsgRet&(1) = lSendMsgRet&(0)
        
        'Only add icons if we have not scrolled all the Files.----------------------
        If lListItemCounter(0) < lListItemCounter&(1) Then
            Call frmMain.subAddIcons(mVariables.sExplorerTreePath)
        End If
        '---------------------------------------------------------------------------
    
    End If
    
End Sub
'***********************************************************************************

