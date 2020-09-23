Attribute VB_Name = "modNotifyAreaBehave"
Option Explicit

'-------------------------------------------------------------------------------------------------------------
'Registry Functions ------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------
'Enumeration - Registry Keys
Public Enum RegistryKeys
   HKEY_CLASSES_ROOT = &H80000000
   HKEY_CURRENT_CONFIG = &H80000005
   HKEY_CURRENT_USER = &H80000001
   HKEY_LOCAL_MACHINE = &H80000002
   HKEY_USERS = &H80000003
   HKEY_DYN_DATA = &H80000006         'Windows 95/98
   HKEY_PERFORMANCE_DATA = &H80000004 'Windows NT/2000
End Enum

'Constants - RegSetValueEx(dwType)
Private Const REG_NONE = 0                 'No defined value type.
Private Const REG_SZ = 1                   'A null-terminated string. It will be a Unicode or ANSI string, depending on whether you use the Unicode or ANSI functions.
Private Const REG_EXPAND_SZ = 2            'A null-terminated string that contains unexpanded references to environment variables (for example, "%PATH%"). It will be a Unicode or ANSI string depending on whether you use the Unicode or ANSI functions. To expand the environment variable references, use the ExpandEnvironmentStrings function.
Private Const REG_BINARY = 3               'Binary data in any form.
Private Const REG_DWORD = 4                'A 32-bit number.
Private Const REG_DWORD_LITTLE_ENDIAN = 4  'A 32-bit number in little-endian format. This is equivalent to REG_DWORD. In little-endian format, a multi-byte value is stored in memory from the lowest byte (the "little end") to the highest byte. For example, the value 0x12345678 is stored as (0x78 0x56 0x34 0x12) in little-endian format. Windows NT/Windows 2000, Windows 95, and Windows 98 are designed to run on little-endian computer architectures. A user may connect to computers that have big-endian architectures, such as some UNIX systems.
Private Const REG_DWORD_BIG_ENDIAN = 5     'A 32-bit number in big-endian format. In big-endian format, a multi-byte value is stored in memory from the highest byte (the "big end") to the lowest byte. For example, the value 0x12345678 is stored as (0x12 0x34 0x56 0x78) in big-endian format.
Private Const REG_LINK = 6                 'A Unicode symbolic link. Used internally; applications should not use this type.
Private Const REG_MULTI_SZ = 7             'An array of null-terminated strings, terminated by two null characters.
Private Const REG_RESOURCE_LIST = 8        'A device-driver resource list.
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9 'Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10

'Constants - RegCreateKeyEx(dwOptions)
Private Const REG_OPTION_NON_VOLATILE = 0   'This key is not volatile; this is the default. The information is stored in a file and is preserved when the system is restarted. The RegSaveKey function saves keys that are not volatile.
Private Const REG_OPTION_VOLATILE = 1       'Windows NT/2000 : All keys created by the function are volatile. The information is stored in memory and is not preserved when the corresponding registry hive is unloaded. For HKEY_LOCAL_MACHINE, this occurs when the system is shut down. For registry keys loaded by the RegLoadKey function, this occurs when the corresponding RegUnloadKey is performed. The RegSaveKey function does not save volatile keys. This flag is ignored for keys that already exist.
                                            'Windows 95      : This value is ignored. If REG_OPTION_VOLATILE is specified, the RegCreateKeyEx function creates nonvolatile keys and returns ERROR_SUCCESS.
Private Const REG_OPTION_BACKUP_RESTORE = 4 'Windows NT/2000 : If this flag is set, the function ignores the samDesired parameter and attempts to open the key with the access required to backup or restore the key. If the calling thread has the SE_BACKUP_NAME privilege enabled, the key is opened with ACCESS_SYSTEM_SECURITY and KEY_READ access. If the calling thread has the SE_RESTORE_NAME privilege enabled, the key is opened with ACCESS_SYSTEM_SECURITY and KEY_WRITE access. If both privileges are enabled, the key has the combined accesses for both privileges.

'Constants - RegCreateKeyEx(lpdwDisposition)
Private Const REG_CREATED_NEW_KEY = &H1     'The key did not exist and was created.
Private Const REG_OPENED_EXISTING_KEY = &H2 'The key existed and was simply opened without being changed.

'Constants - RegOpenKeyEx(samDesired)
Private Const KEY_CREATE_LINK = 32          'Permission to create a symbolic link.
Private Const KEY_CREATE_SUB_KEY = 4        'Permission to create subkeys.
Private Const KEY_ENUMERATE_SUB_KEYS = 8    'Permission to enumerate subkeys.
Private Const KEY_EXECUTE = 131097          'Permission for read access.
Private Const KEY_NOTIFY = 16               'Permission for change notification.
Private Const KEY_QUERY_VALUE = 1           'Permission to query subkey data.
Private Const KEY_SET_VALUE = 2             'Permission to set subkey data.
Private Const KEY_ALL_ACCESS = 983103       'Combines the KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, KEY_CREATE_SUB_KEY, KEY_CREATE_LINK, and KEY_SET_VALUE access rights, plus all the standard access rights except SYNCHRONIZE.
Private Const KEY_READ = 131097             'Combines the STANDARD_RIGHTS_READ, KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and KEY_NOTIFY access rights.
Private Const KEY_WRITE = 131078            'Combines the STANDARD_RIGHTS_WRITE, KEY_SET_VALUE, and KEY_CREATE_SUB_KEY access rights.

'Constants - GetLastErr_Msg
Private Const MAX_PATH = 260
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

'Variables - REG_GetLastError
Private ErrLast_Num  As Long
Private ErrLast_Desc As String

'General Windows API Declarations
Private Declare Sub SetLastError Lib "KERNEL32.DLL" (ByVal dwErrCode As Long)
Private Declare Function FormatMessage Lib "KERNEL32.DLL" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetLastError Lib "KERNEL32.DLL" () As Long

'Registry Related Windows API Declarations
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
'-------------------------------------------------------------------------------------------------------------
'End Registry Functions ------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'Notification Area Behavior Set
'HMcS Computers
'by Howard L. McHenry
'avhlm@comcast.net
'www.hmcscomputers.com
'02/05/2009
'Sets Notification Area (Systray) icon to:
'  Always Show (Default)
'  Always Hide
'  Hide when inactive
'Stops and restarts explorer to set behavior

'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify\IconStreams (REG_BINARY)
'Start           Length   Type            Data_Type
'--------------------------------------------------
'&H0000 /   0        20   Data            Byte
'&H0014 /  20       522   FileSpec        Unicode
'&H021C / 540        16   Data            Byte
'&H022C / 556       526   Title (ToolTip) Unicode
'--------------------------------------------------
'Record Length     1084
'&H0220: Hide when Inactive = 0, Always Hide = 1, Always Show = 0
'&H0224: Hide when Inactive = 0, Always Hide = 1, Always Show = 2

Public Const BHV_ALWSHOWS = &H1  'Always Show (Default)
Public Const BHV_ALWHIDES = &H2  'Always Hide
Public Const BHV_HIDINACT = &H3  'Hide when Inactive

Private aByte() As Byte

'Usage:
'If BehaviorSet("C:\Program Files\TempTray\TempTray.exe") Then                'Alway Show
'   MsgBox "Notification Area Behavior Successfully Set", vbInformation
'Else
'   MsgBox "Problem Setting Notification Area Behavior", vbCritical
'End If
'
'If BehaviorSet("C:\Program Files\TempTray\TempTray.exe", BHV_ALWHIDES") Then 'Alway Hide
'If BehaviorSet("C:\Program Files\TempTray\TempTray.exe", BHV_HIDINACT") Then 'Hide when inactive
'
'Set Notification Area behavior
Public Function BehaviorSet(ByVal cFileSpec As String, Optional nBehave As Byte = BHV_ALWSHOWS) As Boolean
   Dim lRet  As Boolean
   Dim x     As Long
   Dim lFnd  As Boolean
   Dim nBhv1 As Byte
   Dim nBhv2 As Byte
   '
   cFileSpec = Trim(cFileSpec)
   
   'Get byte array from registry
   lRet = REG_GetBinary_BYTE(HKEY_CURRENT_USER, _
                             "Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify", _
                             "IconStreams", _
                             aByte())
   
   If lRet Then
      'Get a record in registry byte array
      For x = 0 To UBound(aByte) Step 1084
         'Find file spec in registry byte record
         If LCase(GetText(x, 20, 522)) = LCase(cFileSpec) Then
            'Found file spec
            lFnd = True
            Exit For
         End If
      Next
      If lFnd Then
         Select Case nBehave
            Case BHV_ALWSHOWS
               nBhv1 = 0
               nBhv2 = 2
            
            Case BHV_ALWHIDES
               nBhv1 = 1
               nBhv2 = 1
               
            Case BHV_HIDINACT
               nBhv1 = 0
               nBhv2 = 0
         
         End Select
         'Change notification area behavior in registry
         BehaviorSet = SetRegistry(x / 1084, nBhv1, nBhv2)
         
         'Is process running
         If Not IsProcessRun(GetText(x, 20, 522)) Then
            MsgBox "Process is Not Running but Changes were Set:" & vbCrLf & cFileSpec, vbInformation
         End If
      Else
         MsgBox "Can't find File Spec in Registry List, Reboot and Try Again:" & vbCrLf & cFileSpec, vbCritical
      End If
   Else
      MsgBox "Error Retrieving Registry Value:" & vbCrLf & _
             "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify\IconStreams", vbCritical
   End If

End Function

'Usage:
'MsgBox BehaviorGet("C:\Program Files\TempTray\TempTray.exe")
'
'Get current Notification Area behavior of file spec
Public Function BehaviorGet(ByVal cFileSpec As String) As String
   Dim lRet  As Boolean
   Dim x     As Long
   Dim lFnd  As Boolean
   Dim cDat1 As String
   Dim cDat2 As String
   Dim cTxt  As String
   '
   'Get byte array from registry
   lRet = REG_GetBinary_BYTE(HKEY_CURRENT_USER, _
                             "Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify", _
                             "IconStreams", _
                             aByte())
   
   If lRet Then
      'Get a record in registry byte array
      For x = 0 To UBound(aByte) Step 1084
         'Find filespec in registry byte record
         If LCase(GetText(x, 20, 522)) = LCase(cFileSpec) Then
            'Found
            lFnd = True
            Exit For
         End If
      Next
      If lFnd Then
         cTxt = ""
         'Data &H0224 HI = 0, AH = 1, AS = 0
         cDat1 = GetData(x, &H220, &H220)
         
         'Data &H0224 HI = 0, AH = 1, AS = 2
         cDat2 = GetData(x, 548, 548)
         
         If cDat1 & cDat2 <> "" Then
            Select Case cDat1 & " " & cDat2
               Case "00 00"
                  'Hide when Inactive
                  cTxt = "Hide when Inactive"
                  
               Case "01 01"
                  'Always Hide
                  cTxt = "Always Hide"
                  
               Case "00 02"
                  'Always Show
                  cTxt = "Always Show"
                  
               Case Else
                  'UnKnown
                  cTxt = "UnKnown"
            
            End Select
            
            'Test for Past Item
            If Not IsProcessRun(GetText(x, 20, 522)) Then
               cTxt = "Process is Not Running"
            End If
            
            'ToolTip &H022C to &H043C
            cTxt = cTxt & vbCrLf & GetText(x, 556, 1084)
            If Trim(cTxt) <> "" Then
               BehaviorGet = Trim(cTxt)
            End If
         End If
      Else
         BehaviorGet = "Can't find File Spec in Registry List"
      End If
   Else
      BehaviorGet = "Error Retrieving Registry Value:" & vbCrLf & _
                    "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify\IconStreams"
   End If
   
End Function

'Set byte in registry
Private Function SetRegistry(nRec As Long, nData1 As Byte, nData2 As Byte) As Boolean
   Dim x    As Long
   Dim lRet As Boolean
   '
   '&H220 = 544
   '&H224 = 548
   For x = 0 To UBound(aByte) Step 1084
      If x / 1084 = nRec Then
         If x + 548 < UBound(aByte) Then
            aByte(x + 544) = nData1
            aByte(x + 548) = nData2
            Exit For
         Else
            Exit Function
         End If
      End If
   Next
   'Save byte array to registry
   lRet = REG_SaveBinary_BYTE(HKEY_CURRENT_USER, _
                              "Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify", _
                              "IconStreams", _
                              aByte())
   If lRet Then
      SetRegistry = True
      ExplorerTerminate
   Else
      MsgBox "Error Saving Changes to Registry", vbCritical
   End If
   
End Function

'Get unicode text from byte array
Private Function GetText(ByVal nRec As Long, ByVal nStrt As Long, ByVal nStop As Long) As String
   Dim cTxt As String
   Dim i    As Long
   Dim Y    As Long
   '
   For Y = nStrt To nStop
      i = nRec + Y
      If i > UBound(aByte) Then
         Exit For
      End If
      cTxt = cTxt & Chr(aByte(i))
   Next
   cTxt = StrConv(cTxt, vbFromUnicode)
   cTxt = TrimWithoutPrejudice(cTxt)
   
   GetText = cTxt
   
End Function

'Get hex string from byte array
Private Function GetData(ByVal nRec As Long, ByVal nStrt As Long, ByVal nStop As Long) As String
   Dim cTxt As String
   Dim i    As Long
   Dim Y    As Long
   '
   For Y = nStrt To nStop
      i = nRec + Y
      If i > UBound(aByte) Then
         Exit For
      End If
      If cTxt = "" Then
         cTxt = cTxt & Right("00" & Hex(aByte(i)), 2)
      Else
         cTxt = cTxt & " " & Right("00" & Hex(aByte(i)), 2)
      End If
   Next
   
   GetData = cTxt
   
End Function

'Eliminate non-printable characters
Private Function TrimWithoutPrejudice(ByVal InputString As String) As String
   Dim sAns  As String
   Dim sWkg  As String
   Dim sChar As String
   Dim lLen  As Long
   Dim lCtr  As Long
   '
   sAns = InputString
   lLen = Len(InputString)
   
   If lLen > 0 Then
      'Ltrim
      For lCtr = 1 To lLen
         sChar = Mid(sAns, lCtr, 1)
         If Asc(sChar) > 32 Then
            Exit For
         End If
      Next
   
      sAns = Mid(sAns, lCtr)
      lLen = Len(sAns)
   
      'Rtrim
      If lLen > 0 Then
         For lCtr = lLen To 1 Step -1
            sChar = Mid(sAns, lCtr, 1)
            If Asc(sChar) > 32 Then
               Exit For
            End If
         Next
      End If
      sAns = Left$(sAns, lCtr)
   End If
   
   TrimWithoutPrejudice = sAns

End Function

'Explorer terminate and restart
Private Sub ExplorerTerminate()
   Dim Process As Object
   '
   For Each Process In GetObject("winmgmts:"). _
      ExecQuery("select * from Win32_Process where name='explorer.exe'")
      Process.Terminate (0)
   Next
      
End Sub

'Is process running
Private Function IsProcessRun(ByVal cFileSpec As String) As Boolean
   Dim Process As Object
   '
   cFileSpec = Right(cFileSpec, Len(cFileSpec) - InStrRev(cFileSpec, "\"))
   For Each Process In GetObject("winmgmts:"). _
      ExecQuery("select * from Win32_Process where name='" & cFileSpec & "'")
      IsProcessRun = True
   Next
      
End Function

'-------------------------------------------------------------------------------------------------------------
'Registry Functions ------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------

'=============================================================================================================
'REG_GetBinary_BYTE
'
'Purpose :
'¯¯¯¯¯¯¯¯¯
'Function that retrieves a Binary value from the specified registry entry.
'* IMPORTANT - This function will return a BYTE ARRAY in the Variant parameter passed.
'
'Param :          Use :
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'hKey             Specifies the main HKEY_?? to look for the specified key
'strKey           Path to the key to look in for the specified value
'strValue         Name of the binary registry value to get the BYTE data from
'Return_ByteArray Variable that returns the byte array retrieved
'
'Return:
'¯¯¯¯¯¯¯
'Returns FALSE if failed.  Call the REG_GetLastError to get the error number and
'description.
'
'Sample Use:
'¯¯¯¯¯¯¯¯¯¯¯
'=============================================================================================================
Public Function REG_GetBinary_BYTE(ByVal hKey As RegistryKeys, _
                                   ByVal strKey As String, _
                                   ByVal strValue As String, _
                                   ByRef Return_ByteArray As Variant) As Boolean
   On Error GoTo ErrorTrap
   Dim ReturnValue As Long
   Dim TheKey      As Long
   Dim TheType     As Long
   Dim TheSize     As Long
   Dim ByteArray() As Byte
   '
   'Clear return variable
   ReDim Return_ByteArray(0) As Byte
   
   'Get the handle to the registry key specified by the user
   ReturnValue = RegOpenKeyEx(hKey, strKey, 0, KEY_ALL_ACCESS, TheKey)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegOpenKeyEx", ErrLast_Num, ErrLast_Desc, False
      Exit Function
   End If
   
   'Get the size and type of the data
   If REG_GetDataType(hKey, strKey, strValue, TheType, , TheSize) = False Then
      GoTo FreeMemory
   End If
   
   'Make sure that the specified value holds a string value
   If TheType <> REG_BINARY Then
      ErrLast_Num = -1
      ErrLast_Desc = "Specified Key\Value combination is not a 'Binary'value."
      GoTo FreeMemory
   End If
   
   'Resize the buffer to receive the data
   ReDim ByteArray(TheSize - 1) As Byte
   
   'Get the Byte array
   ReturnValue = RegQueryValueEx(TheKey, strValue, 0, TheType, ByteArray(0), TheSize)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegQueryValueEx", ErrLast_Num, ErrLast_Desc, False
      GoTo FreeMemory
   End If
   
   'Assign the return value
   Return_ByteArray = ByteArray
   
   REG_GetBinary_BYTE = True
   
FreeMemory:
   
   'Close the opened key
   ReturnValue = RegCloseKey(TheKey)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegCloseKey", ErrLast_Num, ErrLast_Desc, False
   End If
   
   Exit Function
   
ErrorTrap:
   
   If Err.Number = 0 Then      'No Error
      Resume Next
   ElseIf Err.Number = 20 Then 'Resume Without Error
      Resume Next
   Else                        'Other Error
      ErrLast_Num = Err.Number
      ErrLast_Desc = Err.Description
      Err.Clear
      Err.Number = 0
      REG_GetBinary_BYTE = False
      Exit Function
   End If
  
End Function

'=============================================================================================================
'REG_GetDataType
'
'Purpose :
'¯¯¯¯¯¯¯¯¯
'Function that inspects a specified registry Key\Value to see what type of entry it is.
'
'Param :          Use :
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'hKey             Specifies the main HKEY_?? to look for the specified key
'strKey           Path to the key to look in for the specified value
'strValue         Name of the value to get the data type of
'Return_TypeLNG   Optional. Returns the data type as a LONG variable
'Return_TypeSTR   Optional. Returns the data type as a STRING variable
'Return_DataSize  Optional. Returns the size in BYTEs of the data
'
'Return:
'¯¯¯¯¯¯¯
'Returns FALSE if failed.  Call the REG_GetLastError to get the error number and
'description.
'
'Sample Use:
'¯¯¯¯¯¯¯¯¯¯¯
'=============================================================================================================
Private Function REG_GetDataType(ByVal hKey As RegistryKeys, _
                                ByVal strKey As String, _
                                ByVal strValue As String, _
                                Optional ByRef Return_TypeLNG As Long, _
                                Optional ByRef Return_TypeSTR As String, _
                                Optional ByRef Return_DataSize As Long) As Boolean
   On Error GoTo ErrorTrap
   Dim ReturnValue As Long
   Dim TheKey      As Long
   '
   'Get the handle to the registry key specified by the user
   ReturnValue = RegOpenKeyEx(hKey, strKey, 0, KEY_ALL_ACCESS, TheKey)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegOpenKeyEx", ErrLast_Num, ErrLast_Desc, False
      Exit Function
   End If
   
   'Get the size and type of the data
   ReturnValue = RegQueryValueEx(TheKey, strValue, 0, Return_TypeLNG, ByVal 0&, Return_DataSize)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegQueryValueEx", ErrLast_Num, ErrLast_Desc, False
      GoTo FreeMemory
   End If
   
   'Find what type the return was and return a string equivelent for it
   Select Case Return_TypeLNG
      Case REG_SZ                  '1 = A null-terminated string. It will be a Unicode or ANSI string, depending on whether you use the Unicode or ANSI functions.
         Return_TypeSTR = "String"
      Case REG_BINARY              '3 = Binary data in any form.
         Return_TypeSTR = "Binary"
      Case REG_DWORD               '4 = A 32-bit number.
         Return_TypeSTR = "DWORD"
      Case REG_DWORD_LITTLE_ENDIAN '4 = A 32-bit number in little-endian format. This is equivalent to REG_DWORD. In little-endian format, a multi-byte value is stored in memory from the lowest byte (the "little end") to the highest byte. For example, the value 0x12345678 is stored as (0x78 0x56 0x34 0x12) in little-endian format.
         'Windows NT/Windows 2000, Windows 95, and Windows 98 are designed to run on little-endian computer architectures. A user may connect to computers that have big-endian architectures, such as some UNIX systems.
         Return_TypeSTR = "DWORD - Little Endian"
      Case REG_DWORD_BIG_ENDIAN    '5 = A 32-bit number in big-endian format. In big-endian format, a multi-byte value is stored in memory from the highest byte (the "big end") to the lowest byte. For example, the value 0x12345678 is stored as (0x12 0x34 0x56 0x78) in big-endian format.
         Return_TypeSTR = "DWORD - Big Endian"
      Case REG_EXPAND_SZ           '2 = A null-terminated string that contains unexpanded references to environment variables (for example, "%PATH%"). It will be a Unicode or ANSI string depending on whether you use the Unicode or ANSI functions. To expand the environment variable references, use the ExpandEnvironmentStrings function.
         Return_TypeSTR = "Unexpanded references to an environment variable"
      Case REG_LINK                '6 = A Unicode symbolic link. Used internally; applications should not use this type.
         Return_TypeSTR = "Unicode Symbolic Link"
      Case REG_MULTI_SZ            '7 = An array of null-terminated strings, terminated by two null characters.
         Return_TypeSTR = "String Array"
      Case REG_RESOURCE_LIST       '8 = A device-driver resource list.
         Return_TypeSTR = "Device Driver Resource List"
      Case REG_NONE                '0 = No defined value type.
         Return_TypeSTR = "Undefined Type"
      Case Else
         Return_TypeSTR = "Unknown Type"
   End Select
   
   REG_GetDataType = True
   
FreeMemory:
   
   'Close the opened key
   ReturnValue = RegCloseKey(TheKey)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegCloseKey", ErrLast_Num, ErrLast_Desc, False
   End If
   
   Exit Function
   
ErrorTrap:
   
   If Err.Number = 0 Then      'No Error
      Resume Next
   ElseIf Err.Number = 20 Then 'Resume Without Error
      Resume Next
   Else                        'Other Error
      ErrLast_Num = Err.Number
      ErrLast_Desc = Err.Description
      Err.Clear
      Err.Number = 0
      REG_GetDataType = False
      Exit Function
   End If
   
End Function

'=============================================================================================================
'REG_GetLastError
'
'Purpose :
'¯¯¯¯¯¯¯¯¯
'Function that returns the number and description of the last error that occured.
'
'Param :          Use :
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Return_ErrNum    Optional. If an error occured, returns the error number
'Return_ErrDesc   Optional. If an error occured, returns the error description
'
'Return:
'¯¯¯¯¯¯¯
'Returns FALSE if failed.  Call the REG_GetLastError to get the error number and
'description.
'
'Sample Use:
'¯¯¯¯¯¯¯¯¯¯¯
'=============================================================================================================
Private Function REG_GetLastError(Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String) As Boolean
   On Error Resume Next
   '
   If ErrLast_Num = 0 And ErrLast_Desc = "" Then
      REG_GetLastError = False
      Exit Function
   End If
   
   Return_ErrNum = ErrLast_Num
   Return_ErrDesc = ErrLast_Desc
   REG_GetLastError = True
   
End Function

'=============================================================================================================
'REG_SaveBinary_BYTE
'
'Purpose :
'¯¯¯¯¯¯¯¯¯
'Function that saves a BYTE array to the binary registry Key\Value specified.
'
'Param :          Use :
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'hKey             Specifies the main HKEY_?? to look for the specified key
'strKey           Path to the key to look in for the specified value
'strValue         Name of the binary registry value to save the byte array to
'DynamicByteArray The dynamic byte array to save - passed in the form of a Variant
'
'Return:
'¯¯¯¯¯¯¯
'Returns FALSE if failed.  Call the REG_GetLastError to get the error number and
'description.
'
'Sample Use:
'¯¯¯¯¯¯¯¯¯¯¯
'=============================================================================================================
Public Function REG_SaveBinary_BYTE(ByVal hKey As RegistryKeys, ByVal strKey As String, ByVal strValue As String, ByVal DynamicByteArray As Variant) As Boolean
   On Error GoTo ErrorTrap
   Dim ReturnValue    As Long
   Dim TheKey         As Long
   Dim TheSize        As Long
   Dim ByteArray()    As Byte
   Dim TheDisposition As Long
   '
   'Make sure the value passed is a valid byte array
   If VarType(DynamicByteArray) <> vbByte + vbArray Then
      ErrLast_Num = -1
      ErrLast_Desc = "Invalid byte array passed to the 'REG_SaveBinary_BYTE'function"
      Exit Function
   End If
   
   'Assign the DYNAMIC byte array to a STANDARD byte array to work with
   ByteArray = DynamicByteArray
   
   'If the specified key did not exist before, create it, otherwise open it.
   ReturnValue = RegCreateKeyEx(hKey, strKey, 0, 0, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, TheKey, TheDisposition)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegCreateKeyEx", ErrLast_Num, ErrLast_Desc, False
      Exit Function
   End If
   
   'Test if it was created or opened
   If TheDisposition = REG_CREATED_NEW_KEY Then
      Debug.Print "Created new key"
   ElseIf TheDisposition = REG_OPENED_EXISTING_KEY Then
      Debug.Print "Key already existed"
   End If
   
   'Get the size of the byte array
   TheSize = CLng(UBound(ByteArray)) + 1
   
   'Set the binary value
   ReturnValue = RegSetValueEx(TheKey, strValue, 0, REG_BINARY, ByteArray(LBound(ByteArray)), TheSize)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegSetValueEx", ErrLast_Num, ErrLast_Desc, False
   End If
   
   REG_SaveBinary_BYTE = True
   
FreeMemory:
   
   'Close the opened key
   ReturnValue = RegCloseKey(TheKey)
   If ReturnValue <> 0 Then
      GetLastErr_Msg ReturnValue, "RegCloseKey", ErrLast_Num, ErrLast_Desc, False
   End If
   
   Exit Function
   
ErrorTrap:
   
   If Err.Number = 0 Then      'No Error
      Resume Next
   ElseIf Err.Number = 20 Then 'Resume Without Error
      Resume Next
   Else                        'Other Error
      ErrLast_Num = Err.Number
      ErrLast_Desc = Err.Description
      Err.Clear
      Err.Number = 0
      REG_SaveBinary_BYTE = False
      Exit Function
   End If
   
End Function

Private Function GetLastErr_Msg(Optional ByVal ErrorNumber As Long, Optional ByVal LastAPICalled As String = "last", Optional ByRef Return_ErrNum As Long, Optional ByRef Return_ErrDesc As String, Optional ByVal DisplayError As Boolean = False) As Boolean
   On Error GoTo ErrorTrap
   Dim ErrMsg As String
   '
   'Clear the return variables
   Return_ErrNum = 0
   Return_ErrDesc = ""
   
   'If no error message is specified then check for one
   If ErrorNumber = 0 Then
      ErrorNumber = GetLastError
      If ErrorNumber = 0 Then
         GetLastErr_Msg = False
         Exit Function
      End If
   End If
   
   'Allocate a buffer for the error description
   ErrMsg = String(MAX_PATH + 1, 0)
   
   'Get the error description
   FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrorNumber, 0, ErrMsg, MAX_PATH + 1, 0
   ErrMsg = Left(ErrMsg, InStr(ErrMsg, Chr(0)) - 1)
   
   'Display the error message
   If DisplayError = True Then
      MsgBox "An error occured while calling the " & LastAPICalled & " Windows API function." & Chr(13) & "Below is the error information:" & Chr(13) & Chr(13) & "Error Number = " & CStr(ErrorNumber) & Chr(13) & "Error Description = " & ErrMsg, vbOKOnly + vbExclamation, "  Windows API Error"
   End If
   
   'Return the information
   Return_ErrNum = ErrorNumber
   Return_ErrDesc = ErrMsg
   GetLastErr_Msg = True
   
   'Set the last error to 0 (no error) so next time through it doesn't report the same error twice
   SetLastError 0
   
   Exit Function
   
ErrorTrap:
   
   If Err.Number = 0 Then      'No Error
      Resume Next
   ElseIf Err.Number = 20 Then 'Resume Without Error
      Resume Next
   Else                        'Other Error
      ErrLast_Num = Err.Number
      ErrLast_Desc = Err.Description
      Err.Clear
      Err.Number = 0
      GetLastErr_Msg = True
      Exit Function
   End If
   
End Function


'Only the registry functions needed for this application were extracted from modRegistry_Adv.bas:
'http://www.thevbzone.com/s_modules.htm
'=============================================================================================================
'
' modRegistry_Adv Module
' ----------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Created On  : August 02, 2000
' Last Update : August 21, 2003
'
' VB Versions : 5.0 / 6.0
'
' Requires    : Nothing
'
' Description : This module is meant to make it easy to access advanced registry functionality via the Win32
'               API.  See each function for more details on how each works.
'
' NOTE        : If the user is on Windows NT 4.0 or Windows 2000, the REG_DeleteKey function executes more
'               efficiently if Microsoft Internet Explorer 4.x or greater is installed on their computer.
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================
















