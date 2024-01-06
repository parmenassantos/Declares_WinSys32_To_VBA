'
'   version: 2024-01-06
'   Created by Parmenas Santos
'   parmenassantos@gmail.com
'
'   GitHub repository - check for updates
'   https://github.com/parmenassantos/Declares_WinSys32_To_VBA.git
'
Option Explicit

' // Constants Windows to title end bar
Public Const GWL_STYLE As Long = (-16)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_EX_DLGMODALFRAME As Long = &H1

'// Constants WIndows to transparency
Public Const WS_EX_LAYERED As Long = &H80000
Public Const LWA_COLORKEY As Long = &H1
Public Const LWA_ALPHA As Long = &H2

Public Const BM_CLICK As Long = &HF5&
Public Const BM_SETCHECK As Long = &HF1&
Public Const BST_CHECKED As Long = &H1&
Public Const EM_REPLACESEL As Long = &HC2&
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOACTIVATE As Long = &H10&
Public Const SWP_NOMOVE As Long = &H2&
Public Const SWP_NOSIZE As Long = &H1&
Public Const SWP_SHOWWINDOW As Long = &H40&
Public Const TCM_SETCURFOCUS As Long = &H1330&
Public Const WM_GETICON As Long = &H7F
Public Const SW_NORMAL As Long = 1
Public Const WM_SYSCOMMAND = &H112
Public Const WM_CLOSE = &H10
Public Const SC_CLOSE = &HF060
'Public Const CWM_SETPATH As Long = (WM_USER + 2)
Public Const WM_SETCURSOR As Long = &H20
Public Const WM_SETFOCUS As Long = &H7
Public Const WM_SETFONT As Long = &H30
Public Const WM_SETHOTKEY As Long = &H32
Public Const WM_SETICON As Long = &H80
Public Const WM_SETREDRAW As Long = &HB
Public Const WM_SETTEXT As Long = &HC
'Public Const WM_SETTINGCHANGE As Long = WM_WININICHANGE
'// Commands To manipule Window
Public Const SW_SHOW As Long = 5
Public Const SW_HIDE As Long = 0
Public Const SW_MAXIMIZE As Long = 3
Public Const SW_MINIMIZE As Long = 6
Public Const SW_OTHERUNZOOM As Long = 4
Public Const SW_OTHERZOOM As Long = 2
Public Const SW_RESTORE As Long = 9
Public Const SW_SCROLLCHILDREN As Long = &H1

#If VBA7 Then  '// Win64
       Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (MyDest As Any, MySource As Any, ByVal MySize As Long)
       Public Declare PtrSafe Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
       Public Declare PtrSafe Sub SetCursorPos Lib "user32" (ByVal X As Integer, ByVal Y As Integer)
       Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
       Public Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
       Public Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal lngHWnd As Long) As Long
       Public Declare PtrSafe Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Boolean
       Public Declare PtrSafe Function EndDialog Lib "user32" (ByVal hWnd As Long, ByVal result As Long) As Boolean
       Public Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal pEnumProc As Long, ByVal lParam As Long) As Long
       Public Declare PtrSafe Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
       Public Declare PtrSafe Function FindExecutable Lib "Shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVallpDirectory As String, ByVal lpResult As String) As Long
       Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
       Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
       Public Declare PtrSafe Function GetClassNameA Lib "user32" (ByVal hWnd As Long, ByVal szClassName As String, ByVal lLength As Long) As Long
       Public Declare PtrSafe Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
       Public Declare PtrSafe Function GetCommandLineParams Lib "kernel32" Alias "GetCommandLineA" () As Long
       Public Declare PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
       Public Declare PtrSafe Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
       Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
       Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
       Public Declare PtrSafe Function GetFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
       Public Declare PtrSafe Function GetLastError Lib "kernel32" () As Integer
       Public Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
       Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
       Public Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
       Public Declare PtrSafe Function GetTickCountMs Lib "kernel32" Alias "GetTickCount" () As Long
       Public Declare PtrSafe Function GetUserName Lib "AdvApi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
       Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
       Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As winRect) As Long
       Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal szWindowText As String, ByVal lLength As Long) As Long
       Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
       Public Declare PtrSafe Function IsCharAlphaNumericA Lib "user32" (ByVal byChar As Byte) As Long
       Public Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function LStrCpynA Lib "kernel32" Alias "lstrcpynA" (ByVal pDestination As String, ByVal pSource As Long, ByVal iMaxLength As Integer) As Long
       Public Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
       Public Declare PtrSafe Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
       Public Declare PtrSafe Function PathAddBackslashByPointer Lib "ShlwApi" Alias "PathAddBackslashW" (ByVal lpszPath As Long) As Long
       Public Declare PtrSafe Function PathAddBackslashByString Lib "ShlwApi" Alias "PathAddBackslashW" (ByVal lpszPath As String) As Long 'http://msdn.microsoft.com/en-us/library/aa155716%28office.10%29.aspx
       Public Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
       Public Declare PtrSafe Function RegQueryValue Lib "AdvApi32" (ByVal hKey As Long, ByVal sValueName As String, ByVal dwReserved As Long, ByRef lValueType As Long, ByVal sValue As String, ByRef lResultLen As Long) As Long
       Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
       Public Declare PtrSafe Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
       Public Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare PtrSafe Function SetLocalTime Lib "kernel32" (lpSystem As SystemTime) As Long
       Public Declare PtrSafe Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As winPlacement) As Long
       Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
       Public Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
       Public Declare PtrSafe Function ShellExecute Lib "Shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
       Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
       Public Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
       Public Declare PtrSafe Function StrCpy Lib "kernel32" Alias "lstrcpynA" (ByVal pDestination As String, ByVal pSource As String, ByVal iMaxLength As Integer) As Long
       Public Declare PtrSafe Function StringLen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
       Public Declare PtrSafe Function StrTrimW Lib "ShlwApi" () As Boolean
       Public Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hWnd As Long, ByVal uExitCode As Long) As Long
       Public Declare PtrSafe Function TimeGetTime Lib "Winmm" Alias "timeGetTime" () As Long
       Public Declare PtrSafe Function VarPtrArray Lib "MsVbVm50" Alias "VarPtr" (Var() As Any) As Long
       Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
       Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
       Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
       Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
       Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
       Private Type browseInfo     'used by uBrowseForFolder
              hOwner As Long
              pidlRoot As Long
              pszDisplayName As String
              lpszTitle As String
              ulFlags As Long
              lpfn As Long
              lParam As Long
              iImage As Long
       End Type
       Public Declare PtrSafe Function BrowseForFolder Lib "Shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As browseInfo) As Long
       Private Type ChooseColor    'used by uChooseColor; http://support.microsoft.com/kb/153929 and http://www.cpearson.com/Excel/Colors.aspx
              lStructSize As Long
              hWndOwner As Long
              hInstance As Long
              rgbResult As Long
              lpCustColors As String
              flags As Long
              lCustData As Long
              lpfnHook As Long
              lpTemplateName As String
       End Type
       Public Declare PtrSafe Function ChooseColor Lib "ComDlg32" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
       Private Type FindWindowParameters   'Custom structure for passing in the parameters in/out of the hook enumeration function; could use global variables instead, but this is nicer
              strTitle As String  'INPUT
              hWnd As Long        'OUTPUT
       End Type                            'Find a specific window with dynamic caption from a list of all open windows: http://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground
       Public Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
       Private Type lastInputInfo  'used by uGetLastInputInfo, getLastInputTime
              cbSize As Long
              dwTime As Long
       End Type
       Public Declare PtrSafe Function GetLastInputInfo Lib "user32" (ByRef plii As lastInputInfo) As Long
       'http://www.pgacon.com/visualbasic.htm#Take%20Advantage%20of%20Conditional%20Compilation
       'Logical and Bitwise Operators in Visual Basic: http://msdn.microsoft.com/en-us/library/wz3k228a(v=vs.80).aspx and http://stackoverflow.com/questions/1070863/hidden-features-of-vba
       Private Type SystemTime
              wYear          As Integer
              wMonth         As Integer
              wDayOfWeek     As Integer
              wDay           As Integer
              wHour          As Integer
              wMinute        As Integer
              wSecond        As Integer
              wMilliseconds  As Integer
       End Type
       Public Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystem As SystemTime)
       Private Type pointAPI       'used by uSetWindowPlacement
              X As Long
              Y As Long
       End Type
       Private Type rectAPI       'used by uSetWindowPlacement
              Left_Renamed As Long
              Top_Renamed As Long
              Right_Renamed As Long
              Bottom_Renamed As Long
       End Type
       Public Type winPlacement   'used by uSetWindowPlacement
              length As Long
              flags As Long
              showCmd As Long
              ptMinPosition As pointAPI
              ptMaxPosition As pointAPI
              rcNormalPosition As rectAPI
       End Type
       Public Declare PtrSafe Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As winPlacement) As Long
       Private Type winRect     'used by uMoveWindow
              Left As Long
              Top As Long
              Right As Long
              Bottom As Long
       End Type
       Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As Long, xLeft As Long, ByVal yTop As Long, wWidth As Long, ByVal hHeight As Long, ByVal repaint As Long) As Long
       Public Declare PtrSafe Function InternetOpen Lib "WiniNet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long    'Open the Internet object    'ex: lngINet = InternetOpen(“MyFTP Control”, 1, vbNullString, vbNullString, 0)
       Public Declare PtrSafe Function InternetConnect Lib "WiniNet" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long  'Connect to the network  'ex: lngINetConn = InternetConnect(lngINet, "ftp.microsoft.com", 0, "anonymous", "wally@wallyworld.com", 1, 0, 0)
       Public Declare PtrSafe Function FtpGetFile Lib "WiniNet" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean    'Get a file 'ex: blnRC = FtpGetFile(lngINetConn, "dirmap.txt", "c:\dirmap.txt", 0, 0, 1, 0)
       Public Declare PtrSafe Function FtpPutFile Lib "WiniNet" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean  'Send a file  'ex: blnRC = FtpPutFile(lngINetConn, “c:\dirmap.txt”, “dirmap.txt”, 1, 0)
       Public Declare PtrSafe Function FtpDeleteFile Lib "WiniNet" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean 'Delete a file 'ex: blnRC = FtpDeleteFile(lngINetConn, “test.txt”)
       Public Declare PtrSafe Function InternetCloseHandle Lib "WiniNet" (ByVal hInet As Long) As Integer  'Close the Internet object  'ex: InternetCloseHandle lngINetConn    'ex: InternetCloseHandle lngINet
       Public Declare PtrSafe Function FtpFindFirstFile Lib "WiniNet" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
       Private Type FILETIME
              dwLowDateTime As Long
              dwHighDateTime As Long
       End Type
       Private Type WIN32_FIND_DATA
              dwFileAttributes As Long
              ftCreationTime As FILETIME
              ftLastAccessTime As FILETIME
              ftLastWriteTime As FILETIME
              nFileSizeHigh As Long
              nFileSizeLow As Long
              dwReserved0 As Long
              dwReserved1 As Long
              cFileName As String * 1 'MAX_FTP_PATH
              cAlternate As String * 14
       End Type    'ex: lngHINet = FtpFindFirstFile(lngINetConn, "*.*", pData, 0, 0)
       Public Declare PtrSafe Function InternetFindNextFile Lib "WiniNet" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long  'ex: blnRC = InternetFindNextFile(lngHINet, pData)
#Else  '// Win32, Win16
       Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (MyDest As Any, MySource As Any, ByVal MySize As Long)
       Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
       Public Declare Sub SetCursorPos Lib "user32" (ByVal X As Integer, ByVal Y As Integer)                       'Logical and Bitwise Operators in Visual Basic: http://msdn.microsoft.com/en-us/library/wz3k228a(v=vs.80).aspx and http://stackoverflow.com/questions/1070863/hidden-features-of-vba    'http://www.pgacon.com/visualbasic.htm#Take%20Advantage%20of%20Conditional%20Compilation
       Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
       Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
       Public Declare Function BringWindowToTop Lib "user32" (ByVal lngHWnd As Long) As Long
       Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
       Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Boolean
       Public Declare Function EndDialog Lib "user32" (ByVal hWnd As Long, ByVal result As Long) As Boolean
       Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal pEnumProc As Long, ByVal lParam As Long) As Long
       Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
       Public Declare Function FindExecutable Lib "Shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVallpDirectory As String, ByVal lpResult As String) As Long
       Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
       Public Declare Function GetActiveWindow Lib "user32" () As Long
       Public Declare Function GetClassNameA Lib "user32" (ByVal hWnd As Long, ByVal szClassName As String, ByVal lLength As Long) As Long
       Public Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
       Public Declare Function GetCommandLineParams Lib "kernel32" Alias "GetCommandLineA" () As Long
       Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
       Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
       Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
       Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
       Public Declare Function GetForegroundWindow Lib "user32" () As Long
       Public Declare Function GetFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
       Public Declare Function GetLastError Lib "kernel32" () As Integer
       Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
       Public Declare Function GetTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
       Public Declare Function GetTickCountMs Lib "kernel32" Alias "GetTickCount" () As Long
       Public Declare Function GetUserName Lib "AdvApi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
       Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
       Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As winRect) As Long
       Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal szWindowText As String, ByVal lLength As Long) As Long
       Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
       Public Declare Function IsCharAlphaNumericA Lib "user32" (ByVal byChar As Byte) As Long
       Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function LStrCpynA Lib "kernel32" Alias "lstrcpynA" (ByVal pDestination As String, ByVal pSource As Long, ByVal iMaxLength As Integer) As Long
       Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
       Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
       Public Declare Function PathAddBackslashByPointer Lib "ShlwApi" Alias "PathAddBackslashW" (ByVal lpszPath As Long) As Long
       Public Declare Function PathAddBackslashByString Lib "ShlwApi" Alias "PathAddBackslashW" (ByVal lpszPath As String) As Long 'http://msdn.microsoft.com/en-us/library/aa155716%28office.10%29.aspx
       Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
       Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
       Public Declare Function RegQueryValue Lib "AdvApi32" (ByVal hKey As Long, ByVal sValueName As String, ByVal dwReserved As Long, ByRef lValueType As Long, ByVal sValue As String, ByRef lResultLen As Long) As Long
       Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
       Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
       Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
       Public Declare Function SetLocalTime Lib "kernel32" (lpSystem As SystemTime) As Long
       Public Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As winPlacement) As Long
       Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
       Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
       Public Declare Function ShellExecute Lib "Shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
       Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
       Public Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
       Public Declare Function StrCpy Lib "kernel32" Alias "lstrcpynA" (ByVal pDestination As String, ByVal pSource As String, ByVal iMaxLength As Integer) As Long
       Public Declare Function StringLen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
       Public Declare Function StrTrimW Lib "ShlwApi" () As Boolean
       Public Declare Function TerminateProcess Lib "kernel32" (ByVal hWnd As Long, ByVal uExitCode As Long) As Long
       Public Declare Function TimeGetTime Lib "Winmm" Alias "timeGetTime" () As Long
       Public Declare Function VarPtrArray Lib "MsVbVm50" Alias "VarPtr" (Var() As Any) As Long
       Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
       Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
       Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
       Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
       Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
       Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
       Public Type browseInfo     'used by uBrowseForFolder
              hOwner As Long
              pidlRoot As Long
              pszDisplayName As String
              lpszTitle As String
              ulFlags As LongB
              lpfn As Long
              lParam As Long
              iImage As Long
       End Type
       Public Declare Function BrowseForFolder Lib "Shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As browseInfo) As Long
       Private Type ChooseColor    'used by uChooseColor; http://support.microsoft.com/kb/153929 and http://www.cpearson.com/Excel/Colors.aspx
              lStructSize As Long
              hWndOwner As Long
              hInstance As Long
              rgbResult As Long
              lpCustColors As String
              flags As Long
              lCustData As Long
              lpfnHook As Long
              lpTemplateName As String
       End Type
       Public Declare Function ChooseColor Lib "ComDlg32" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
       Private Type FindWindowParameters   'Custom structure for passing in the parameters in/out of the hook enumeration function; could use global variables instead, but this is nicer
              strTitle As String  'INPUT
              hWnd As Long        'OUTPUT
       End Type                            'Find a specific window with dynamic caption from a list of all open windows: http://www.everythingaccess.com/tutorials.asp?ID=Bring-an-external-application-window-to-the-foreground
       Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
       Private Type lastInputInfo  'used by uGetLastInputInfo, getLastInputTime
              cbSize As Long
              dwTime As Long
       End Type
       Public Declare Function GetLastInputInfo Lib "user32" (ByRef plii As lastInputInfo) As Long
       Private Type SystemTime
              wYear          As Integer
              wMonth         As Integer
              wDayOfWeek     As Integer
              wDay           As Integer
              wHour          As Integer
              wMinute        As Integer
              wSecond        As Integer
              wMilliseconds  As Integer
       End Type
       Public Declare Sub GetLocalTime Lib "kernel32" (lpSystem As SystemTime)
       Private Type pointAPI
              X As Long
              Y As Long
       End Type
       Private Type rectAPI
              Left_Renamed As Long
              Top_Renamed As Long
              Right_Renamed As Long
              Bottom_Renamed As Long
       End Type
       Private Type winPlacement
              length As Long
              flags As Long
              showCmd As Long
              ptMinPosition As pointAPI
              ptMaxPosition As pointAPI
              rcNormalPosition As rectAPI
       End Type
       Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, ByRef lpwndpl As winPlacement) As Long
       Private Type winRect
              Left As Long
              Top As Long
              Right As Long
              Bottom As Long
       End Type
       Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, xLeft As Long, ByVal yTop As Long, wWidth As Long, ByVal hHeight As Long, ByVal repaint As Long) As Long
#End If
