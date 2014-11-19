Attribute VB_Name = "FileSupport"
Option Explicit

Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

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

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Type FILE_INFO
    FileName As String
    FilePath As String
    'FileParentFolder As String
    'FileSize As Currency
    FileSize As Double
    FileTimeCreated As Date
    FileTimeLastModified As Date
    FileTimeLastAccessed As Date
    FileAttributes As Long
    'IsFolder As Boolean
End Type

Public CancelFindFiles As Boolean


Public Function FileExists(ByVal FileName As String) As Boolean

On Error Resume Next

Dim lFileHandle As Long, WFD As WIN32_FIND_DATA
Dim bFileFound As Boolean
    
    lFileHandle = FindFirstFile(FileName, WFD)
    bFileFound = (lFileHandle <> INVALID_HANDLE_VALUE)
        
    Call FindClose(lFileHandle)

    FileExists = bFileFound
    
End Function


Public Function FolderExists(ByVal FolderName As String) As Boolean

    On Error GoTo FolderExists_EH

    Dim bFound As Boolean, sRet As String

100     If Right$(FolderName, 1) = "\" Then
102         FolderName = Left$(FolderName, Len(FolderName) - 1)
        End If
    
104     If Len(FolderName) > 2 Then
106         bFound = FileExists(FolderName)
108         If Not bFound Then bFound = Len(Dir$(AddSlash(FolderName)))
        
        Else
110         FolderName = Left$(FolderName, 1)
112         sRet = ListDrives
114         bFound = (InStr(1, sRet, FolderName, vbTextCompare) > 0)

        End If

FolderExists_End:

        On Error Resume Next
    
116     FolderExists = bFound

        Exit Function
    

FolderExists_EH:

118     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in FileSupport.FolderExists()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
120     bFound = False
    
122     Resume FolderExists_End
    
End Function


Public Function GetParentDir(ByVal Path As String, Optional IncludeTrailingBackslash As Boolean = True) As String

On Error Resume Next

    If Len(Path) > 3 Then
        Path = IIf(Right$(Path, 1) = "\", Left$(Path, Len(Path) - 1), Path)
        Path = Left$(Path, InStrRev(Path, "\") - 1)
        If IncludeTrailingBackslash Then Path = AddSlash(Path)
    End If
    
    GetParentDir = Path

End Function


Public Function MkDirEx(ByVal Path As String) As Boolean

On Error GoTo MkDirEx_EH

Dim iCtr As Integer
Dim lRet As Long
Dim sCurrentDir As String
Dim SecAttrib As SECURITY_ATTRIBUTES
    
    ' Make sure the path has a "\" at the end
    Path = Path & IIf(Right$(Path, 1) = "\", "", "\")
    
    ' Is the path on a local/mapped drive OR a UNC share?
    Select Case Mid$(Path, 2, 1)
    Case "\"
        iCtr = 3 ' this puts us out past the "\\"
        iCtr = InStr(iCtr, Path, "\") ' this puts us out past the computer name
        iCtr = InStr(iCtr + 1, Path, "\") ' this puts us out past the share name
        
    Case ":"
        iCtr = 0
        
    Case Else
        MkDirEx = False
        Exit Function
        
    End Select
    
    ' Walk the string, creating directories as needed ...
    Do Until InStr(iCtr + 1, Path, "\") = 0
    
        iCtr = InStr(iCtr + 1, Path, "\")
        sCurrentDir = Left$(Path, iCtr)
        
        If Len(sCurrentDir) > 3 Then
            
            If Not FileExists(IIf(Right$(sCurrentDir, 1) = "\", Left$(sCurrentDir, Len(sCurrentDir) - 1), sCurrentDir)) Then
                
                ' Create the directory
                With SecAttrib
                    .lpSecurityDescriptor = &O0
                    .bInheritHandle = False
                    .nLength = Len(SecAttrib)
                End With
                lRet = CreateDirectory(sCurrentDir, SecAttrib)
                
                ' If we fail, there is no reason to continue
                If lRet = 0 Then Exit Function
            
            End If
        
        End If
        
        iCtr = iCtr + 1
    
    Loop
    
    ' If we've made it here, then we must have suceeded
    MkDirEx = True
    
    Exit Function


MkDirEx_EH:
    
    Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in MkDirEx(" & Path & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
    MkDirEx = False
    
End Function


Public Function ListDrives() As String

Dim sRet As String, lRet As Long, lCtr As Long, sTmp As String
Dim sDrives As String

    'Create a buffer to store all the drives
    sRet = String(256, Chr$(0))
    'Get all the drives
    lRet = GetLogicalDriveStrings(255, sRet)
    
    'Extract the drives from the buffer and print them on the form
    For lCtr = 1 To lRet Step 4
        sTmp = Mid$(sRet, lCtr, 3)
        If Left$(sTmp, 1) = Chr$(0) Then Exit For
        sDrives = sDrives & Left$(sTmp, 1)
    Next
    
    ListDrives = sDrives

End Function


Public Function FindFiles(ByVal FilePath As String, ByRef FileSpecInclude() As String, ByVal FileSpecExclude As String, ByVal UseArchiveBit As Boolean, ByVal Recursive As Boolean, ByVal StartChangedDate As Date, ByVal EndChangedDate As Date, ByRef FilesFound() As FILE_INFO, ByRef Results As String) As Long

    On Error GoTo FindFiles_EH

    Dim lFileHandle As Long, WFD As WIN32_FIND_DATA, lBaseCnt As Long, lFoundCnt As Long, lIncludeCtr As Long
    Dim FTM As FILETIME, STM As SYSTEMTIME, dtmStartChangedDate As Date, dtmEndChangedDate As Date, dtmTmp(2) As Date, udtFI As FILE_INFO
    Dim bCont As Boolean, sRet As String, lRet As Long
    Dim sErrMsg As String

100     If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
        'FileSpecInclude = LCase$(FileSpecInclude)
        'If FileSpecInclude = "" Then FileSpecInclude = "*"
102     FileSpecExclude = LCase$(FileSpecExclude)

104     If StartChangedDate <> CDate("12:00:00 AM") Then
106         dtmStartChangedDate = StartChangedDate
        Else
108         dtmStartChangedDate = CDate("1/1/1971")
        End If
        
110     dtmEndChangedDate = EndChangedDate

112     bCont = True
    
        'RaiseEvent Searching(FilePath & sRet, m_bCancel)
        'Debug.Print "Searching " & FilePath & sRet & " ..."
114     frmMain.UpdateStatus "Searching " & FilePath & sRet & " ...", False, CancelFindFiles

116     lFileHandle = FindFirstFile(FilePath & "*.*", WFD)
    
118     If lFileHandle <> INVALID_HANDLE_VALUE Then
        
120         lBaseCnt = UBoundEx(FilesFound)

122         Do While bCont And (Not CancelFindFiles)

124             sRet = LCase$(TrimNulls(WFD.cFileName))
            
                ' Ignore the directory dots
126             If (sRet <> ".") And (sRet <> "..") Then

                    ' Check for directory with bitwise comparison.
128                 If Not ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) Then
                
130                     For lIncludeCtr = 0 To UBound(FileSpecInclude())
    
132                         If (sRet Like FileSpecInclude(lIncludeCtr)) And (Not sRet Like FileSpecExclude) Then
        
134                             lRet = FileTimeToLocalFileTime(WFD.ftCreationTime, FTM)
136                             lRet = FileTimeToSystemTime(FTM, STM)
138                             dtmTmp(0) = STimeToVBTime(STM)
        
140                             lRet = FileTimeToLocalFileTime(WFD.ftLastWriteTime, FTM)
142                             lRet = FileTimeToSystemTime(FTM, STM)
144                             dtmTmp(1) = STimeToVBTime(STM)
                            
146                             lRet = FileTimeToLocalFileTime(WFD.ftLastAccessTime, FTM)
148                             lRet = FileTimeToSystemTime(FTM, STM)
150                             dtmTmp(2) = STimeToVBTime(STM)
        
152                             If EndChangedDate = CDate("12:00:00 AM") Then dtmEndChangedDate = Now()
        
154                             If (((dtmTmp(0) >= dtmStartChangedDate) Or (dtmTmp(1) >= dtmStartChangedDate)) And _
                                    ((dtmTmp(0) <= dtmEndChangedDate) Or (dtmTmp(1) <= dtmEndChangedDate)) And _
                                    (IIf(UseArchiveBit, (WFD.dwFileAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE, True))) Then
        
156                                 lFoundCnt = lFoundCnt + 1
158                                 ReDim Preserve FilesFound(lBaseCnt + lFoundCnt)
            
160                                 sRet = TrimNulls(WFD.cFileName)
            
162                                 With udtFI
                                    
                                        'RaiseEvent FoundFile(FilePath & sRet, m_bCancel)
        
164                                     FilesFound(lBaseCnt + lFoundCnt - 1).FileName = sRet
166                                     FilesFound(lBaseCnt + lFoundCnt - 1).FilePath = FilePath
168                                     FilesFound(lBaseCnt + lFoundCnt - 1).FileAttributes = WFD.dwFileAttributes
170                                     FilesFound(lBaseCnt + lFoundCnt - 1).FileTimeCreated = dtmTmp(0)
172                                     FilesFound(lBaseCnt + lFoundCnt - 1).FileTimeLastModified = dtmTmp(1)
174                                     FilesFound(lBaseCnt + lFoundCnt - 1).FileTimeLastAccessed = dtmTmp(2)
176                                     FilesFound(lBaseCnt + lFoundCnt - 1).FileSize = FileSizeFromWFD(WFD.nFileSizeHigh, WFD.nFileSizeLow)
                                    
                                    End With
        
                                End If
        
                            End If
    
                        Next

                    Else
                
178                     If Recursive Then
                        
180                         lRet = FindFiles(FilePath & TrimNulls(WFD.cFileName), FileSpecInclude, FileSpecExclude, UseArchiveBit, Recursive, StartChangedDate, EndChangedDate, FilesFound, sRet)
    
182                         If lRet >= 0 Then
184                             lFoundCnt = lFoundCnt + lRet
                        
                            Else
186                             GoTo FindFiles_End
    
                            End If
    
                        End If
                
                    End If
                
                End If

188             bCont = (Not bCloseApp) And (FindNextFile(lFileHandle, WFD) <> 0) ' Get next entry

            Loop
    
190         If Not bCont Then If bCloseApp Then Debug.Print "Cancelling FindFiles(" & FilePath & ") because app is closing"

        End If
       
FindFiles_End:

        On Error Resume Next
    
192     If lFileHandle <> 0 Then
194         Call FindClose(lFileHandle)
196         lFileHandle = 0
    
        End If

198     Results = sErrMsg

200     FindFiles = lFoundCnt

        Exit Function
    
    
FindFiles_EH:

        'MsgBox "Error [" & Err.Number & "] " & Err.Description & IIf(Erl <> 0, " at line " & Erl, "") & " where" & vbCrLf & vbCrLf & _
               "   Path = " & FilePath & FileSpecInclude, vbExclamation Or vbOKOnly, "An unexpected error has occurred in FindFiles() ..."

202     sErrMsg = "Error [" & Err.Number & "] " & Err.Description & IIf(Erl <> 0, " at line " & Erl, "") & " occurred in FindFiles()" ' & FilePath & FileSpec & ")"
        'RaiseEvent Error("FindFiles(" & FilePath & FileSpecInclude & ")", Err.Number, Err.Description, m_bCancel)

204     lFoundCnt = -1

206     Resume FindFiles_End
    
End Function


Private Function FileSizeFromWFD(FileSizeHigh As Long, FileSizeLow As Long) As Double
        
    On Error GoTo FileSizeFromWFD_EH
        
    Dim dblRet As Double

100     dblRet = CDbl((FileSizeHigh * 2147483647) + FileSizeLow)
    
FileSizeFromWFD_End:
    
        On Error Resume Next

102     FileSizeFromWFD = dblRet
    
        Exit Function


FileSizeFromWFD_EH:

104     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in FileSupport.FileSizeFromWFD()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next

106     dblRet = -1

End Function


Public Function TrimNulls(ByVal Value As String) As String

' strips any nulls from the end of a string

On Error Resume Next

Dim nFirstNull As Long

    'Value = Trim$(Value)
    
    If Len(Value) Then
        
        If Left$(Value, 1) <> vbNullChar Then
        
            nFirstNull = InStr(1, Value, vbNullChar)
            
            If nFirstNull > 0 Then
                TrimNulls = Left$(Value, nFirstNull - 1)
            
            Else
                TrimNulls = Value
            
            End If
        
        Else
            TrimNulls = ""
        
        End If
    
    Else
        TrimNulls = ""

    End If

End Function


Private Function STimeToVBTime(ByRef Value As SYSTEMTIME) As Date

On Error Resume Next

    With Value
        STimeToVBTime = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond)
    End With

End Function


Public Function UBoundEx(Target() As FILE_INFO) As Long

On Error GoTo UBoundEx_EH

Dim lRet As Long

    lRet = UBound(Target)

UBoundEx_End:

    On Error Resume Next

    UBoundEx = lRet
    
    Exit Function
    

UBoundEx_EH:

    lRet = 0
    
    Resume UBoundEx_End

End Function


