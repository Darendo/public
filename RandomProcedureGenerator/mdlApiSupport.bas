Attribute VB_Name = "ApiSupport"
Option Explicit

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal _
    lpApplicationName As String, ByVal _
    lpKeyName As Any, ByVal _
    lpDefault As String, ByVal _
    lpReturnedString As String, ByVal _
    nSize As Long, ByVal _
    lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal _
    lpApplicationName As String, ByVal _
    lpKeyName As Any, ByVal _
    lpString As Any, ByVal _
    lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Sub SleepEx(Optional Milliseconds As Long = 1)

On Error Resume Next

Dim lTicks As Long

    lTicks = GetTickCount + Milliseconds

    Do
        Sleep IIf(Milliseconds = 1 Or Milliseconds = 0, 0, 1)
        DoEvents
    Loop While GetTickCount < lTicks

End Sub


Public Function GetRandomNumber(MinValue As Double, MaxValue As Double) As Long

On Error Resume Next

Dim lRet As Long

    Randomize
    lRet = CLng(Round(((MaxValue - MinValue + IIf(MinValue > 0, 1, 0)) * Rnd) + MinValue))
    If lRet > MaxValue Then lRet = MaxValue
    If lRet < MinValue Then lRet = MinValue
    
    GetRandomNumber = lRet

End Function


Public Function FormatAge(dtmDOB As Date) As String

Dim lAgeInDays As Long
    
    lAgeInDays = DateDiff("d", dtmDOB, Now()) + 1
    
    If lAgeInDays > 730 Then
        ' Format AGE in years
        FormatAge = DateDiff("yyyy", dtmDOB, Now()) & " years"
    
    ElseIf lAgeInDays > 90 Then
        ' Format AGE in months
        FormatAge = DateDiff("m", dtmDOB, Now()) & " months"
    
    ElseIf lAgeInDays > 1 Then
        ' Format AGE in days
        FormatAge = lAgeInDays & " days"
    
    Else
        ' Format AGE in hours
        FormatAge = DateDiff("h", dtmDOB, Now()) & " hours"
    
    End If

End Function


Public Function FormatHeight(lValue As Long, sFormat As eFormatUnits) As String
    
    Select Case sFormat
    
    Case huEnglish
        
        If lValue > 36 Then
            FormatHeight = lValue \ 12 & "' " & lValue Mod 12 & """"
        
        Else
            FormatHeight = lValue & """"
        
        End If
    
    Case huMetric
        
        If lValue > 100 Then
            FormatHeight = lValue \ 100 & "m " & lValue Mod 100 & "cm"
        
        Else
            FormatHeight = lValue & "cm"
        
        End If
    
    Case Else
        
        FormatHeight = CStr(lValue)
    
    End Select

End Function


Public Function FormatWeight(lValue As Long, sFormat As eFormatUnits) As String
    
    Select Case sFormat
    
    Case huEnglish
        
        ' lb.s and oz.s
        If lValue > 16 Then
            FormatWeight = lValue \ 16 & " lb.s"
            If lValue Mod 16 > 0 Then
                FormatWeight = FormatWeight & ", " & lValue Mod 16 & " oz.s"
            End If
        
        Else
            FormatWeight = lValue & " oz.s"
        
        End If
    
    Case huMetric
    
    Case Else
        FormatWeight = CStr(lValue)
    
    End Select

End Function


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

    Case "MYSQL ODBC 5.1 DRIVER", "MYSQL ODBC 3.51 DRIVER"
        
        sRet = "DRIVER=" & Provider & ";" & _
               "SERVER=" & DataSource & ";" & _
               "DATABASE=" & InitialCatalog & ";" & _
               "UID=" & UserID & ";PWD=" & Password

    Case Else
        sRet = ""

    End Select
    
    BuildDbConnectionString = sRet
    
End Function


Public Function GetMachineName() As String
    
Dim dwLen As Long
Dim strString As String
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    
    'get only the actual data
    strString = Left(strString, dwLen)
    
    'Return the computer name
    GetMachineName = strString

End Function


Public Function NowEx(Optional sDelimeter As String = ":", Optional bPadWithZeroes As Boolean = True, Optional bUseDecimalForMillisecond As Boolean = False) As String
    NowEx = TimeToMillisecond(sDelimeter, bPadWithZeroes, bUseDecimalForMillisecond)
End Function


Public Function TimeToMillisecond(Optional sDelimeter As String = ":", Optional bPadWithZeroes As Boolean = True, Optional bUseDecimalForMillisecond As Boolean = False) As String
    
    Dim sFineTime As String
    Dim sFormatString As String
    
    If bPadWithZeroes Then
        sFormatString = "00"
    Else
        sFormatString = "0"
    End If
    
    On Error Resume Next
    sFineTime = Format$(Hour(Now), sFormatString) & sDelimeter & Format$(Minute(Now), sFormatString) & sDelimeter & Format$(Second(Now), sFormatString) & IIf(bUseDecimalForMillisecond, ".", sDelimeter) & Format$(GetMilliseconds, sFormatString & sFormatString)
    TimeToMillisecond = sFineTime

End Function


Public Function GetMilliseconds() As Integer

Dim typTime As SYSTEMTIME
    
    GetSystemTime typTime
    
    GetMilliseconds = typTime.wMilliseconds

End Function


Public Function GetGMT() As Date
    
Dim typTime As SYSTEMTIME
Dim sRet As Date
    
    GetSystemTime typTime
    
    sRet = typTime.wDay & "/" & typTime.wMonth & "/" & typTime.wYear & " " & typTime.wHour & ":" & typTime.wMinute & ":" & typTime.wSecond
    
    GetGMT = CDate(sRet)
    
End Function


Public Function AddSlash(strDir As String) As String
    
'On Error GoTo AddSlash_EH
Dim strErrMsg As String
    
    AddSlash = strDir & IIf(Right(strDir, 1) = "\", "", "\")
    
Exit Function


AddSlash_EH:
    
    strErrMsg = "Error [" & Err.Number & "] " & Err.Description & vbCrLf & " in AddSlash(" & strDir & ")"
    MsgBox strErrMsg & " in AddSlash()", vbExclamation, "Error Message"

End Function


Function AppendString(strFileName As String, strText As String) As Boolean

'On Error GoTo AppendString_EH
Dim strErrMsg As String

Dim intFileNumber As Integer

    intFileNumber = FreeFile ' Get the next free file number
    
    Open strFileName For Append As intFileNumber ' Open the file for output.
        Print #intFileNumber, strText ' Append the string
    Close intFileNumber ' Close the file
        
    AppendString = True ' Exit clean
Exit Function


AppendString_EH:

    strErrMsg = "Error [" & Err.Number & "] " & Err.Description
    MsgBox strErrMsg & " in AppendString()", vbExclamation, "Error Message"

End Function


Public Function GetINISetting(ByVal strSection As String, ByVal strKey As String, ByVal strDefault As String, ByVal strINIFile As String) As String

'On Error GoTo GetINISetting_EH
Dim strErrMsg As String

Dim ret As String, numChars As Long

ret = String(255, 0)
numChars = GetPrivateProfileString(strSection, strKey, strDefault, ret, 255, strINIFile)
If numChars > 0 Then
    GetINISetting = Left(ret, numChars)
End If

Exit Function


GetINISetting_EH:

    strErrMsg = "Error [" & Err.Number & "] " & Err.Description
    strErrMsg = strErrMsg & vbCrLf & vbCrLf
    strErrMsg = strErrMsg & "Section: " & strSection & vbCrLf
    strErrMsg = strErrMsg & "Key: " & strKey & vbCrLf
    strErrMsg = strErrMsg & "Default: " & strDefault & vbCrLf
    strErrMsg = strErrMsg & "INI File: " & strINIFile & vbCrLf
    
    MsgBox strErrMsg & " in GetINISetting()", vbExclamation, "Error Message"

End Function


Public Function SaveINISetting(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, ByVal strINIFile As String) As Boolean

'On Error GoTo SaveINISetting_EH
Dim strErrMsg As String

Dim numChars As Long

    numChars = WritePrivateProfileString(strSection, strKey, strValue, strINIFile)
    If numChars > 0 Then SaveINISetting = True

Exit Function


SaveINISetting_EH:
    
    strErrMsg = "Error [" & Err.Number & "] " & Err.Description
    strErrMsg = strErrMsg & vbCrLf & vbCrLf
    strErrMsg = strErrMsg & "Section: " & strSection & vbCrLf
    strErrMsg = strErrMsg & "Key: " & strKey & vbCrLf
    strErrMsg = strErrMsg & "Value: " & strValue & vbCrLf
    strErrMsg = strErrMsg & "INI File: " & strINIFile & vbCrLf
    
    MsgBox strErrMsg & " in SaveINISetting()", vbExclamation, "Error Message"

End Function


Public Function IsFormLoaded(ByVal FormName As String) As Boolean

Dim frm As Form
    
    For Each frm In Forms
        If UCase$(frm.Name) = UCase$(FormName) Then
            IsFormLoaded = True
            Exit For
        End If
    Next frm

End Function

