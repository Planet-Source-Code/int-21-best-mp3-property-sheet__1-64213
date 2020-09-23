Attribute VB_Name = "RegistryStuff"
'Unknown Author

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
Public Enum HKEYS_Constants
  HKEY_CLS_ROOT = &H80000000
  HKEY_CUR_USER = &H80000001
  HKEY_LOC_MAC = &H80000002
  HKEY_USERZ = &H80000003
  HKEY_PERF_DATU = &H80000004
End Enum

Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4

Public GetDwordError As Boolean

Public Sub SaveKey(hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub
Public Function GetString(hkey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    'text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))

            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function
Public Sub SaveString(hkey As Long, strPath As String, strValue As String, strData As String)
    'EXAMPLE:
    'Call savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
    Dim keyhand As Long
    Dim r As Long
    If strData <> "" Then
      r = RegCreateKey(hkey, strPath, keyhand)
      r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
      r = RegCloseKey(keyhand)
    Else
      DeleteValue hkey, strPath, strValue
    End If
End Sub
Function GetDWord(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    'EXAMPLE:
    'Text1.text = getdword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        GetDwordError = False
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
        'Else
        'Call errlog("GetDWORD-" & strPath, Fals
        '     e)
    Else
      GetDwordError = True
    End If
    r = RegCloseKey(keyhand)
End Function
Function SaveDWord(ByVal hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    'Call SaveDword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword", text1.text)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then
    '     Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function
Public Function DeleteKey(ByVal hkey As Long, ByVal strKey As String)
    'EXAMPLE:
    'Call DeleteKey(HKEY_CURRENT_USER, "Software\VBW")
    Dim r As Long
   'r = RegOpenKey(hKey, strKey, keyhand)
    r = RegDeleteKey(hkey, strKey)
    'r = RegCloseKey(keyhand)
End Function
Public Function DeleteValue(ByVal hkey As Long, ByVal strPath As String, ByVal strValue As String)
    On Error GoTo Urd:
    'EXAMPLE:
    'Call DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    Dim keyhand As Long
    r = RegOpenKey(hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
Urd:
End Function

