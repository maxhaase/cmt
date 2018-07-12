Attribute VB_Name = "Global"
Option Explicit

Public blFilesIni As Boolean, blPhrasesIni As Boolean
Public strConnection As String

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_READ = &H20019
Private Const REG_SZ As Long = 1
Private Const ERROR_SUCCESS = 0
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const strRootKey = "SOFTWARE\MAX\DB\"


' Returns the registry key containing the connection string
Public Function FetchConnectionString(ByVal strPortalId As String) As String
   Dim hKey As Long  ' handle to the SOFTWARE\MAX\DB\
   Dim retval As Long  ' function's return value
   Dim slength As Long  ' receives length of returned data
   Dim stringbuffer As String

   stringbuffer = Space(255)
   slength = 255

   retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRootKey & strPortalId, 0, KEY_READ, hKey)
   retval = RegQueryValueEx(hKey, "CONNECTIONSTRING", 0, 0, ByVal stringbuffer, slength)

   FetchConnectionString = Trim(stringbuffer)

   retval = RegCloseKey(hKey)
 
End Function

Public Function SetConnectionString(ByVal objEnv As Environment) As String
On Error GoTo EH
    Dim hKey As Long
    Dim strCon As String
    Dim retval As Long
    Dim lResult As Long
    
    
    strCon = "Provider=" & objEnv.strProvider & ";" & _
    "Persist Security Info=" & "False" & ";" & _
    "User ID=" & objEnv.strLoginId & ";" & _
    "Password=" & objEnv.strPassword & ";" & _
    "Initial Catalog=" & objEnv.strDatabase & ";" & _
    "Data Source=" & objEnv.strServer

    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRootKey, 0, KEY_ALL_ACCESS, hKey)
    If ERROR_SUCCESS <> retval Then
        Err.Raise vbObjectError + 513
    End If
    
    'RegCreateKeyEx hKey, objEnv.strEnvironmentName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVa
    RegCreateKeyEx HKEY_LOCAL_MACHINE, strRootKey & objEnv.strEnvironmentName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lResult
    retval = RegOpenKeyEx(HKEY_LOCAL_MACHINE, strRootKey & objEnv.strEnvironmentName, 0, KEY_ALL_ACCESS, hKey)
    retval = RegSetValueExString(hKey, "CONNECTIONSTRING", 0&, REG_SZ, strCon, Len(strCon))
    
    RegCloseKey hKey

    SetConnectionString = strCon
      
Exit Function
EH:
SetConnectionString = ""
End Function


