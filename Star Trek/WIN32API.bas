Attribute VB_Name = "WIN32API"
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type MCI_SYSINFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
        dwNumber As Long
        wDeviceType As Long
End Type

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        wProcessorLevel As Integer
        wProcessorRevision As Integer
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Type FREQ_INFO
    in_cycles As Long    ' Internal clock cycles during
    ex_ticks As Long     ' Microseconds elapsed during
    raw_freq As Long     ' Raw frequency of CPU in MHz
    norm_freq As Long    ' Normalized frequency of CPU
End Type

' multimedia functions
Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

' functions contained in the cpuinf32.dll file
Public Declare Function CPUIDSupport Lib "cpuinf32" Alias "wincpuidsupport" () As Long
Public Declare Function CPUID Lib "cpuinf32" Alias "wincpuid" () As Long
Public Declare Function CPUIDExt Lib "cpuinf32" Alias "wincpuidext" () As Long
Public Declare Function CPUFeatures Lib "cpuinf32" Alias "wincpufeatures" () As Long
Public Declare Function CPUCount Lib "cpuinf32" Alias "ProcessorCount" () As Long
Public Declare Function CPUSpeed Lib "cpuinf32" Alias "cpurawspeed" () As Long
Public Declare Function CPUNormSpeed Lib "cpuinf32" Alias "cpuspeed" () As FREQ_INFO

' system information functions
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Public Declare Function IsProcessorFeaturePresent Lib "wnaspi32" (ByVal ProcessorFeature As Long) As Boolean
Public Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As String
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

' ini file functions
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' drive information functions
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

' windows tasklist functions
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

' mouse cursor functions
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Global myVer As OSVERSIONINFO
Global Point As POINTAPI

' private ini file routines
Public Sub SaveProfile(ByVal Section$, ByVal Key$, ByVal Value$, ByVal fileINI$)
    Dim ret&
    ret = WritePrivateProfileString(Section, Key, Value, fileINI)
End Sub

Public Function LoadProfile(ByVal Section$, ByVal Key$, ByVal fileINI$) As String
    Dim ret&, temp$
    temp = Space$(80)
    ret = GetPrivateProfileString(Section, Key, "", temp$, 80, fileINI)
    If ret <> 0 Then
        LoadProfile = temp$
    Else
        LoadProfile = ""
    End If
End Function


