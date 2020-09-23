Attribute VB_Name = "Module1"
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String)

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hkey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long

Private Declare Function RegConnectRegistry Lib "advapi32.dll" _
        Alias "RegConnectRegistryA" _
        (ByVal lpMachineName As String, _
        ByVal hkey As Long, _
        phkResult As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" _
        Alias "RegEnumValueA" _
        (ByVal hkey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Byte, _
        lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hkey As Long) As Long

Public Enum HKEYs
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Const READ_CONTROL = &H20000





Const REG_SZ = 1                         ' Unicode nul terminated string



Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.





         ' Note that if you declare the lpData parameter as String, you must pass it By Value.


Public Function RGGetKeyValue(hkey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
    
    Dim lngRet As Long
    Dim lngResult As Long
    Dim sData As String
    
    lngRet = RegOpenKeyEx(hkey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet = ERROR_SUCCESS Then
        sData = String(128, vbNullChar)
        
        lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
        
        If Not lngRet = ERROR_SUCCESS Then RGGetKeyValue = Default: Exit Function
        RGGetKeyValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
        RegCloseKey lngResult
    Else
        RGGetKeyValue = Default
    End If
End Function

Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)

       Dim lRetVal As Long
       Dim hkey As Long

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hkey)
       lRetVal = RegDeleteValue(hkey, sValueName)
       RegCloseKey (hkey)
End Function


