Attribute VB_Name = "ApiSupport"
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Const ERROR_BAD_USERNAME = 2202&

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const PROGRESS_CANCEL = 1
Private Const PROGRESS_CONTINUE = 0
Private Const PROGRESS_QUIET = 3
Private Const PROGRESS_STOP = 2
Private Const COPY_FILE_FAIL_IF_EXISTS = &H1
Private Const COPY_FILE_RESTARTABLE = &H2

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long


Public Function InIDE(Optional IDEExeName As String = "vb6.exe") As Boolean

    On Error GoTo InIDE_EH

    Dim sModuleName As String, lIDEExeNameLen As Long
    Dim bRet As Boolean

100     sModuleName = GetEXEPathName
    
102     If sModuleName <> "" Then

104         lIDEExeNameLen = Len(IDEExeName)
106         bRet = (LCase$(Right$(sModuleName, lIDEExeNameLen)) = IDEExeName)

        Else
108         bRet = False
        
        End If
    
110     InIDE = bRet
    
        Exit Function
    

InIDE_EH:

112     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in InIDE()" & IIf(Erl <> 0, " at line " & Erl, "")
    
114     InIDE = False

End Function


Public Function GetEXEPathName() As String

    On Error GoTo GetEXEPathName_EH

    Dim sModuleName As String, lRet As Long

110     sModuleName = String(512, 0)
115     lRet = GetModuleFileName(App.hInstance, sModuleName, 512)
120     sModuleName = Left$(sModuleName, lRet)

125     GetEXEPathName = sModuleName

        Exit Function


GetEXEPathName_EH:
    
130     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in GetEXEPathName()" & IIf(Erl <> 0, " at line " & Erl, "")

135     GetEXEPathName = ""
    
End Function


Public Function GetLastApiError(ErrorNumber As Long) As String

On Error Resume Next

Dim sBuffer As String
    
    'Create a string buffer
    sBuffer = Space(200)
    
    'Format the message string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrorNumber, LANG_NEUTRAL, sBuffer, 200, ByVal 0&

    'Show the message
    GetLastApiError = sBuffer

End Function


Public Function GetFileName(ByVal FileName As String, Optional ByVal IncludeExt As Boolean = True) As String

On Error Resume Next

Dim iStart As Integer, iEnd As Integer

    iStart = InStrRev(FileName, "\", , vbTextCompare) + 1
    If IncludeExt = True Then
        iEnd = (Len(FileName) - iStart) + 1
    Else
        iEnd = InStrRev(FileName, ".", , vbTextCompare)
        iEnd = iEnd - iStart
    End If
    
    GetFileName = Mid$(FileName, iStart, iEnd)
    
End Function


Public Function IsFormLoaded(ByVal FormName As String) As Boolean

On Error Resume Next

Dim oFrm As Form, bRet As Boolean

    bRet = False
    
    For Each oFrm In Forms
        If LCase$(oFrm.Name) = LCase$(FormName) Then
            bRet = True
            Exit For
        
        End If
    
    Next

    IsFormLoaded = bRet

End Function


Public Sub SleepEx(Optional Milliseconds As Long = 1)

On Error Resume Next

Dim lTicks As Long

    lTicks = GetTickCount + Milliseconds

    Do
        Sleep IIf(Milliseconds = 1 Or Milliseconds = 0, 0, 1)
        DoEvents
    Loop While GetTickCount < lTicks

End Sub


Public Function GetINISection(ByVal Section As String, ByRef KeyPairs() As String, ByVal INIFile As String) As Long

    On Error GoTo GetINISection_EH

    Dim lRet As Long, sBuffer As String
    Dim sErrMsg As String

100     lRet = 4096
102     sBuffer = String$(lRet, 0)
104     lRet = GetPrivateProfileSection(Section, sBuffer, lRet, INIFile)
106     If lRet > 0 Then
108         sBuffer = Left$(sBuffer, lRet - 1)
110         If sBuffer <> "" Then
112             If InStr(1, sBuffer, vbNullChar) > 0 Then
114                 KeyPairs = Split(sBuffer, vbNullChar)
    
                Else
116                 ReDim KeyPairs(0)
118                 KeyPairs(0) = sBuffer
    
                End If
120             lRet = UBound(KeyPairs()) + 1
    
            Else
122             lRet = 0
    
            End If
    
        Else
124         lRet = 0

        End If

GetINISection_End:

        On Error Resume Next

126     GetINISection = lRet
    
        Exit Function


GetINISection_EH:

128     sErrMsg = "Error [" & Err.Number & "] " & Err.Description & " occurred in SyncDocAgent.ApiSupport.GetINISection()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
130     Debug.Print sErrMsg

132     lRet = 0
    
134     Resume GetINISection_End
    
End Function


Public Function GetINISetting(ByVal SectionName As String, ByVal KeyName As String, ByVal DefaultValue As String, ByVal INIFile As String) As String

    On Error GoTo GetINISetting_EH

    Dim lRet As Long, sRet As String
    Dim sErrMsg As String

100     lRet = 1024
102     sRet = String(lRet, 0)
104     lRet = GetPrivateProfileString(SectionName, KeyName, DefaultValue, sRet, lRet, INIFile)
106     If lRet >= 0 Then
108         sRet = Left$(sRet, lRet)
        End If

GetINISetting_End:

        On Error Resume Next

110     GetINISetting = sRet
    
        Exit Function


GetINISetting_EH:

112     sErrMsg = "Error [" & Err.Number & "] " & Err.Description & " occurred in SyncDocAgent.ApiSupport.GetINISetting()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
114     Debug.Print sErrMsg

116     sRet = ""
    
118     Resume GetINISetting_End
    
End Function


Public Function SaveINISetting(ByVal SectionName As String, ByVal KeyName As String, ByVal Value As String, ByVal INIFile As String) As Boolean

    On Error GoTo SaveINISetting_EH

    Dim lRet As Long
    Dim sErrMsg As String

100     lRet = WritePrivateProfileString(SectionName, KeyName, Value, INIFile)
    
SaveINISetting_End:

        On Error Resume Next

102     SaveINISetting = (lRet > 0)

        Exit Function


SaveINISetting_EH:
    
104     sErrMsg = "Error [" & Err.Number & "] " & Err.Description & " occurred in SyncDocAgent.ApiSupport.SaveINISetting()" & IIf(Erl <> 0, " at line " & Erl, "")

        On Error Resume Next

106     Debug.Print sErrMsg

108     lRet = 0
    
110     Resume SaveINISetting_End
    
End Function


Public Function AddSlash(ByVal Value As String, Optional Forward As Boolean = False) As String
    
On Error Resume Next

Dim sSlash As String

    If Forward Then
        sSlash = "/"
    Else
        sSlash = "\"
    End If
    
    Value = Trim$(Value)
    
    If Value <> "" Then
        AddSlash = Value & IIf(Right$(Value, 1) = sSlash, "", sSlash)
    End If

End Function


Public Function CBoolEx(Value As String, Optional Default As Boolean = False) As Boolean

On Error GoTo CBoolEx_EH

Dim bRet As Boolean

    bRet = CBool(Value)

    CBoolEx = bRet
    
    Exit Function
    

CBoolEx_EH:

    CBoolEx = Default

End Function


Public Function NowEx(Optional ByVal DateFormat As String = "yyyy-mm-dd", _
                      Optional ByVal DateTimeSeparator As String = " ", _
                      Optional ByVal TimeFormat As String = "hh:nn:ss", _
                      Optional ByVal TimeMillisecondSeparator As String = ".", _
                      Optional ByVal MillisecondFormat As String = "000") As String

On Error Resume Next

Dim sDate As String, sTime As String, sMils As String
Dim sRet As String

    If DateFormat <> "" Then sDate = Format$(Now(), DateFormat)
    If TimeFormat <> "" Then sTime = Format$(Now(), TimeFormat)
    If MillisecondFormat <> "" Then sMils = Format$(GetMillisecond(), MillisecondFormat)

    sRet = sDate & _
           IIf(sDate <> "" And sTime <> "", DateTimeSeparator, "") & sTime & _
           IIf(sTime <> "" And sMils <> "", TimeMillisecondSeparator & sMils, "")

    NowEx = sRet

End Function


Private Function GetMillisecond() As Integer

On Error Resume Next

Dim typTime As SYSTEMTIME
    
    GetSystemTime typTime
    
    GetMillisecond = typTime.wMilliseconds

End Function


Public Sub AppendString(ByVal FileName As String, ByVal TextString As String, Optional ByVal PrefixTimeStamp As Boolean = True)

On Error GoTo AppendString_End

Dim hFileNumber As Integer
    
    hFileNumber = 0
    
    If FileName = "" Then
        FileName = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & Format$(Now(), "yyyymmdd") & ".log"
        If PrefixTimeStamp Then TextString = Format$(Now(), "hh:mm:ss") & " " & Format$(GetMillisecond, "000") & vbTab & TextString
    Else
        If PrefixTimeStamp Then TextString = Format$(Now(), "mm/dd/yy hh:mm:ss") & " " & Format$(GetMillisecond, "000") & vbTab & TextString
    End If
    
    hFileNumber = FreeFile                          ' Get the next free file number
    Open FileName For Append As hFileNumber         ' Open the file for output.
    Print #hFileNumber, TextString                  ' Append the string

AppendString_End:

    On Error Resume Next
    
    If hFileNumber <> 0 Then Close hFileNumber      ' Close the file

End Sub


Public Function BuildDbConnectionString(Optional ByVal Provider As String = "{SQL SERVER}", _
                                          Optional ByVal DataSource As String = "", _
                                          Optional InitialCatalog As String = "", _
                                          Optional UseNTSecurity As Boolean = True, _
                                          Optional UserID As String = "", _
                                          Optional Password As String = "", _
                                          Optional ConnectTimeout As Long = 30) As String
    
On Error Resume Next

Dim sRet As String

    Select Case UCase$(Provider)
    
    Case "{SQL SERVER}"
    
        If UseNTSecurity Then
        
            sRet = "DRIVER=" & Provider & ";" & _
                   "SERVER=" & DataSource & ";" & _
                   "DATABASE=" & InitialCatalog & ";" & _
                   "Trusted_Connection=Yes;" & _
                   "Connect Timeout=" & CStr(ConnectTimeout)

        Else
        
            sRet = "DRIVER=" & Provider & ";" & _
                   "SERVER=" & DataSource & ";" & _
                   "DATABASE=" & InitialCatalog & ";" & _
                   "dsn='';" & _
                   "uid=" & UserID & ";" & _
                   "pwd=" & Password & ";" & _
                   "Connect Timeout=" & CStr(ConnectTimeout)
        End If

    Case "SQLOLEDB", "SQLOLEDB.1"
    
        If UseNTSecurity = True Then
    
            sRet = "Provider=" & Provider & ";" & _
                   "Data Source=" & DataSource & ";" & _
                   "Initial Catalog=" & InitialCatalog & ";" & _
                   "Integrated Security=SSPI;" & _
                   "Persist Security Info=False;" & _
                   "Connect Timeout=" & CStr(ConnectTimeout)
    
        Else
    
            sRet = "Provider=" & Provider & ";" & _
                   "Data Source=" & DataSource & ";" & _
                   "Initial Catalog=" & InitialCatalog & ";" & _
                   "User ID=" & UserID & ";" & _
                   "Password=" & Password & ";" & _
                   "Connect Timeout=" & CStr(ConnectTimeout)
    
        End If
    
    Case "MICROSOFT.JET.OLEDB.4.0", "MICROSOFT.JET.OLEDB.3.51"

        sRet = "Provider=" & Provider & ";" & _
               "Data Source=" & DataSource & ";" & _
               IIf(Password <> "", "Jet OLEDB:Database Password=" & Password & ";", "")

    Case "MYSQL ODBC 3.51 DRIVER", "MYSQL ODBC 5.1 DRIVER"
        
        sRet = "DRIVER={" & UCase$(Provider) & "};" & _
               "SERVER=" & DataSource & ";" & _
               "DATABASE=" & InitialCatalog & ";" & _
               "UID=" & UserID & ";" & _
               "PWD=" & Password & ";" & _
               "OPTION=3"

    Case "MYODBC"
        
        sRet = "Driver={mySQL};" & _
               "Server=" & DataSource & ";" & _
               "Port=3306;" & _
               "Option=;" & _
               "Stmt=;" & _
               "Database=" & InitialCatalog & ";" & _
               "Uid=" & UserID & ";" & _
               "Pwd=" & Password

    Case Else
        sRet = ""

    End Select
    
    BuildDbConnectionString = sRet
    
End Function

