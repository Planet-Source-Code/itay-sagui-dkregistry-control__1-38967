VERSION 5.00
Begin VB.UserControl dkRegistry 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   Picture         =   "dkRegistry.ctx":0000
   ScaleHeight     =   435
   ScaleWidth      =   450
   ToolboxBitmap   =   "dkRegistry.ctx":0972
End
Attribute VB_Name = "dkRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Privilege declarations. Used for Save/Restore keys
Private Const TOKEN_QUERY As Long = &H8&
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_RESTORE_NAME = "SeRestorePrivilege" 'Important for what we're trying to accomplish
Private Const SE_BACKUP_NAME = "SeBackupPrivilege"
Private Type LUID
   lowpart As Long
   highpart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
   pLuid As LUID
   Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   Privileges As LUID_AND_ATTRIBUTES
End Type
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long                'Used to adjust your program's security privileges, can't restore without it!
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long          'Returns a valid LUID which is important when making security changes in NT.
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

' Registry data type
Private Const REG_SZ As Long = 1
Private Const REG_EXPAND_SZ As Long = 2
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4
Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4
Private Const REG_DWORD_BIG_ENDIAN As Long = 5
Private Const REG_LINK As Long = 6
Private Const REG_MULTI_SZ As Long = 7
Private Const REG_QWORD As Long = (11)
Private Const REG_QWORD_LITTLE_ENDIAN As Long = (11)

' Buffer size
Private Const BUFFER_SIZE As Long = 255

Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

Private Const REG_FORCE_RESTORE As Long = 8& ' Almost as import, will allow you to restore over a key while it's open!

' Notification declarations
Private Const REG_NOTIFY_CHANGE_NAME = &H1            ' Create or delete (child)
Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4            ' time stamp
Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
Private Const REG_NOTIFY_ALL = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)

Private Const STANDARD_RIGHTS_ALL = &H1F0000

Private Const SYNCHRONIZE = &H100000

Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegNotifyChangeKeyValue Lib "advapi32" (ByVal hKey As Long, ByVal bWatchSubtree As Boolean, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronous As Boolean) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long ' Main function
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Sub UserControl_Resize()
    If Width <> 420 Then Width = 420
    If Height <> 420 Then Height = 420
End Sub

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String$(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim vData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, vData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = vData
            End If
        End If
    End If
End Function

Public Sub GetString(ByRef Item As clsItem)
Dim Ret As Long
    'Open the key
    RegOpenKey Item.hKey, Item.strPath, Ret
    'Get the key's content
    Item.vData = RegQueryStringValue(Ret, Item.strValue)
    'Close the key
    RegCloseKey Ret
End Sub

Public Sub SaveData(ByVal Item As clsItem)
Dim Ret As Long
    'Create a new key
    RegCreateKey Item.hKey, Item.strPath, Ret
    'Save a string to the key
    RegSetValueEx Ret, Item.strValue, 0, TypeLng(Item.strType), ByVal Item.vData, Len(Item.vData)
    'close the key
    RegCloseKey Ret
End Sub

Public Sub DelValue(Item As clsItem)
Dim Ret As Long
    'Create a new key
    RegCreateKey Item.hKey, Item.strPath, Ret
    'Delete the key's value
    RegDeleteValue Ret, Item.strValue
    'close the key
    RegCloseKey Ret
End Sub

Public Sub DeleteKey(Item As clsItem)
    RegDeleteKey Item.hKey, Item.strPath
End Sub

Public Function CreateKey(Item As clsItem) As Long
Dim Ret As Long
Dim Result As Long
    RegCreateKeyEx Item.hKey, Item.strPath, ByVal 0&, _
        Item.strType, REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, ByVal 0&, Result, Ret
    CreateKey = Result
End Function

Private Function TypeLng(sType As String) As Long
    Select Case sType
    Case "REG_SZ":     TypeLng = REG_SZ
    Case "REG_DWORD":  TypeLng = REG_DWORD
    Case "REG_BINARY": TypeLng = REG_BINARY
    End Select
End Function

Private Function TypeStr(lType As Long) As String
    Select Case lType
    Case REG_SZ:               TypeStr = "REG_SZ"
    Case REG_DWORD:            TypeStr = "REG_DWORD"
    Case REG_BINARY:           TypeStr = "REG_BINARY"
    Case REG_EXPAND_SZ:        TypeStr = "REG_EXPAND_SZ"
    Case REG_DWORD_BIG_ENDIAN: TypeStr = "REG_DWORD_BIG_ENDIAN"
    Case REG_LINK:             TypeStr = "REG_LINK"
    Case REG_MULTI_SZ:         TypeStr = "REG_MULTI_SZ"
    Case REG_QWORD:            TypeStr = "REG_QWORD"
    End Select
End Function

Public Function EnumKeys(Item As clsItem) As colItems
Dim hKey As Long, Cnt As Long, Ret As Long
Dim sName As String, sType As String
Dim TempCol As New colItems
Dim TempItem As New clsItem
    Ret = BUFFER_SIZE
    If RegOpenKey(Item.hKey, Item.strPath, hKey) = 0 Then
        sType = Space$(BUFFER_SIZE)
        Cnt = 0
        sName = "Starting"
        While sName <> sType
            sName = Space$(BUFFER_SIZE)
            RegEnumKeyEx hKey, Cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&
            If sName <> sType Then
                TempCol.Add Item.hKey, Item.strPath, StripNulls(sName), vbNullString, vbNullString
                Cnt = Cnt + 1
            End If
        Wend
        'close the registry key
        RegCloseKey hKey
        Set EnumKeys = TempCol
    End If
End Function

Public Function EnumValues(Item As clsItem) As colItems
Dim hKey As Long, Cnt As Long, Ret As Long
Dim sName As String, lName As Long
Dim sType As String, lType As Long
Dim vData As Variant, lData As Long
Dim TempStr As String
Dim TempCol As New colItems
Dim TempItem As New clsItem
    If RegOpenKey(Item.hKey, Item.strPath, hKey) = 0 Then
        TempStr = Space$(BUFFER_SIZE)
        Cnt = 0
        sName = "Starting"
        While sName <> sType
            sName = Space$(BUFFER_SIZE) ': lName = BUFFER_SIZE
            'sType = Space$(BUFFER_SIZE): lType = BUFFER_SIZE
            Ret = BUFFER_SIZE
            lData = BUFFER_SIZE
            RegEnumValue hKey, Cnt, sName, Ret, 0, lType, vData, lData
            'RegEnumValue hKey, Cnt, sName, Ret, 0, ByVal 0&, ByVal sData, RetData
            If sName <> sType Then
                TempCol.Add Item.hKey, Item.strPath, StripNulls(sName), vbNullString, vbNullString
                Cnt = Cnt + 1
            End If
        Wend
        'close the registry key
        RegCloseKey hKey
        Set EnumValues = TempCol
    End If
End Function

Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, vbNullChar) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, vbNullChar) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Sub NotifyKeyChange(Item As clsItem)
    RegNotifyChangeKeyValue Item.hKey, True, REG_NOTIFY_ALL, ByVal 0&, False
End Sub

Public Function RestoreKey(Item As clsItem, ByVal sFileName As String) As Boolean
    If EnablePrivilege(SE_RESTORE_NAME) = False Then Exit Function
Dim hKey As Long, lRetVal As Long
    Call RegOpenKeyEx(Item.hKey, Item.strPath, 0&, KEY_ALL_ACCESS, hKey)   ' Must open key to restore it
    'The file it's restoring from was created using the RegSaveKey function
    Call RegRestoreKey(hKey, sFileName, REG_FORCE_RESTORE)
    RegCloseKey hKey ' Don't want to keep the key ope. It causes problems.
End Function

Public Function SaveKey(Item As clsItem, ByVal sFileName As String) As Boolean
    If EnablePrivilege(SE_BACKUP_NAME) = False Then Exit Function
Dim hKey As Long, lRetVal As Long
    Call RegOpenKeyEx(Item.hKey, Item.strPath, 0&, KEY_ALL_ACCESS, hKey)     ' Must open key to save it
    'Don't forget to "KILL" any existing files before trying to save the registry key!
    If Dir$(sFileName) <> "" Then Kill sFileName
    Call RegSaveKey(hKey, sFileName, ByVal 0&)
    RegCloseKey hKey ' Don't want to keep the key ope. It causes problems.
End Function

Private Function EnablePrivilege(seName As String) As Boolean
    Dim p_lngRtn As Long
    Dim p_lngToken As Long
    Dim p_lngBufferLen As Long
    Dim p_typLUID As LUID
    Dim p_typTokenPriv As TOKEN_PRIVILEGES
    Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
    p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)
    If p_lngRtn = 0 Then
        Exit Function ' Failed
    ElseIf Err.LastDllError <> 0 Then
        Exit Function ' Failed
    End If
    p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)  'Used to look up privileges LUID.
    If p_lngRtn = 0 Then
        Exit Function ' Failed
    End If
    ' Set it up to adjust the program's security privilege.
    p_typTokenPriv.PrivilegeCount = 1
    p_typTokenPriv.Privileges.Attributes = SE_PRIVILEGE_ENABLED
    p_typTokenPriv.Privileges.pLuid = p_typLUID
    EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function

