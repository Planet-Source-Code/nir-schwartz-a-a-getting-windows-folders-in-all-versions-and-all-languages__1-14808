Attribute VB_Name = "GetReg"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
(ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal HKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal HKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
(ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName _
As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, _
lpftLastWriteTime As Any) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
(ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As _
Byte, lpcbData As Long) As Long
Public Function GetSettingString(ByVal HKey As Long, _
ByVal strPath As String, ByVal strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(HKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If

Else
' there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Function
