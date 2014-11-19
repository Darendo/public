VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ERROR_BAD_FILENAME        As Long = 1     ' Bad input file name specified
Private Const ERROR_NO_DATA             As Long = 2     ' No data to save
Private Const ERROR_NO_FILE_SPECIFIED   As Long = 3     ' No file specified when saving
Private Const ERROR_FILE_WRITE          As Long = 4     ' Could not write to file
Private Const ERROR_CREATE_NEW          As Long = 5     ' Error creating new data


Private m_lBlockSize    As Long         ' Size of buffer when reading file; Default = 32K, Max = 64K

Private szData()        As String       ' Contains the file data
Private lLines          As Long         ' Number of lines
Private lCurLine        As Long         ' Holds the current line when searching
Private lCurPos         As Long         ' Holds the current position in the current line when searching

Private szSearch        As String       ' String being searched
Private nSearchCase     As Integer      ' Case of search term

Private m_lErrorNum     As Long         ' Error number
Private m_sErrorMsg     As String       ' Error message

Private Sub Class_Initialize()
    
    m_lErrorNum = 0
    m_sErrorMsg = ""
    
    m_lBlockSize = 32768 ' Initialize the block size to 32K

End Sub



Public Sub CreateNew(lLineCount As Long)

On Error Resume Next
    
    m_lErrorNum = 0
    m_sErrorMsg = ""
    
    ReDim szData(1 To lLineCount) As String
    If Err Then
        m_lErrorNum = vbObjectError + ERROR_CREATE_NEW
        m_sErrorMsg = Error
        Exit Sub
    End If
    
    lLines = lLineCount
    
End Sub


Public Function Load(ByVal szFileName As String) As Boolean

'-- Load the contents of a text file into memory
'   This routine will handle any line that ends with
'   a carriage return, a linefeed, or both.
    
Dim szBuffer    As String
Dim lFileNum    As Long
Dim lFileLen    As Long
Dim lNumBlocks  As Long
Dim lRemainder  As Long
Dim lIndex      As Long
Dim lPos        As Long
    
    m_lErrorNum = 0
    m_sErrorMsg = ""
    
    '-- Open the file
    lFileNum = FreeFile
    Open szFileName For Binary As lFileNum
    lFileLen = LOF(lFileNum)
    
    '-- Does the file exist?
    If lFileLen = 0 Then
        Close lFileNum
        m_lErrorNum = vbObjectError + ERROR_BAD_FILENAME
        m_sErrorMsg = "File Does Not Exist"
        Load = False
        Exit Function
    End If
    
    '-- Clear the current array
    Erase szData
    lLines = 0
    
    '-- Validate the buffer size
    If m_lBlockSize <= 0 Then
        m_lBlockSize = 32768
    ElseIf m_lBlockSize > 65535 Then
        m_lBlockSize = 65535
    End If
    
    '-- Get the number of blocks
    lNumBlocks = lFileLen \ m_lBlockSize
    
    '-- Anything left over?
    lRemainder = lFileLen Mod m_lBlockSize
    
    '-- Read and process each block
    For lIndex = 1 To lNumBlocks
        szBuffer = Space$(m_lBlockSize)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    Next
    
    '-- Process whatever's left
    If lRemainder Then
        szBuffer = Space$(lRemainder)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    End If

    '-- Close the file and exit
    Close lFileNum

    Load = True
    
    Exit Function


ProcessData:
    
    Do
        '-- Find the next CR
        lPos = InStr(szBuffer, vbCr)
        
        If lPos Then
            
            '-- Copy the text up to the CRLF into szData
            lLines = lLines + 1
            ReDim Preserve szData(1 To lLines) As String
            szData(lLines) = Left$(szBuffer, lPos - 1)
            '-- If the next character is a linefeed, skip over it.
            If Asc(Mid$(szBuffer, lPos + 1, 1)) = 10 Then
                szBuffer = Mid$(szBuffer, lPos + 2)
            Else
                szBuffer = Mid$(szBuffer, lPos + 1)
            End If
        
        Else
            
            '-- Find the next LF
            lPos = InStr(szBuffer, vbLf)
            
            If lPos Then
                '-- Copy the text up to the CRLF into szData
                lLines = lLines + 1
                ReDim Preserve szData(1 To lLines) As String
                szData(lLines) = Left$(szBuffer, lPos - 1)
                szBuffer = Mid$(szBuffer, lPos + 1)
            
            Else
                '-- No more CRLFs. More data?
                If Len(szBuffer) Then
                    '-- Yep. Add it to szData
                    lLines = lLines + 1
                    ReDim Preserve szData(1 To lLines) As String
                    szData(lLines) = szBuffer
                End If
                '-- No more data. Exit the loop
                Exit Do
            
            End If
        
        End If
    
    Loop
    
    Return

End Function


Public Sub LoadListBox(ByVal szFileName As String, List1 As Control)

'-- Loads the contents of a text file into a list box
'   This routine will handle any line that ends with
'   a carriage return, a linefeed, or both.
    
    Dim szBuffer    As String
    Dim lFileNum    As Long
    Dim lFileLen    As Long
    Dim lNumBlocks  As Long
    Dim lRemainder  As Long
    Dim lIndex      As Long
    Dim lPos        As Long
    
    m_lErrorNum = 0
    m_sErrorMsg = ""
    
    '-- Open the file
    lFileNum = FreeFile
    Open szFileName For Binary As lFileNum
    lFileLen = LOF(lFileNum)
    
    '-- Does the file exist?
    If lFileLen = 0 Then
        Close lFileNum
        m_lErrorNum = vbObjectError + ERROR_BAD_FILENAME
        m_sErrorMsg = "File Does Not Exist"
        Exit Sub
    End If
    
    '-- Clear the List Box
    List1.Clear
    
    '-- Get the buffer size
    If m_lBlockSize = 0 Then
        m_lBlockSize = 32768
    ElseIf m_lBlockSize > 65535 Then
        m_lBlockSize = 65535
    End If
    
    '-- Get the number of blocks
    lNumBlocks = lFileLen \ m_lBlockSize
    
    '-- Anything left over?
    lRemainder = lFileLen Mod m_lBlockSize
    
    '-- Read and process each block
    For lIndex = 1 To lNumBlocks
        szBuffer = Space$(m_lBlockSize)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    Next
    
    '-- Process whatever's left
    If lRemainder Then
        szBuffer = Space$(lRemainder)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    End If

    '-- Close the file and exit
    Close lFileNum

    Exit Sub

ProcessData:
    
    Do
        '-- Find the next CR
        lPos = InStr(szBuffer, vbCr)
        If lPos Then
            '-- Copy the text up to the CRLF into List1
            List1.AddItem Left$(szBuffer, lPos - 1)
            '-- If the next character is a linefeed, skip over it.
            If Asc(Mid$(szBuffer, lPos + 1, 1)) = 10 Then
                szBuffer = Mid$(szBuffer, lPos + 2)
            Else
                szBuffer = Mid$(szBuffer, lPos + 1)
            End If
        Else
            '-- Find the next LF
            lPos = InStr(szBuffer, vbLf)
            If lPos Then
                '-- Copy the text up to the CRLF into List1
                List1.AddItem Left$(szBuffer, lPos - 1)
                szBuffer = Mid$(szBuffer, lPos + 1)
            Else
                '-- No more CRLFs. More data?
                If Len(szBuffer) Then
                    '-- Yep. Add it to the list box
                    List1.AddItem szBuffer
                End If
                '-- No more data. Exit the loop
                Exit Do
            End If
        End If
    Loop
    
    Return

End Sub


Public Sub Save(ByVal szFileName As String)

Dim szBuffer    As String
Dim lFileNum    As Long
Dim lFileLen    As Long
Dim lNumBlocks  As Long
Dim lRemainder  As Long
Dim lIndex      As Long
Dim lPos        As Long
        
    '-- Any data in this object?
    If lLines = 0 Then
        m_lErrorNum = vbObjectError + ERROR_NO_DATA
        m_sErrorMsg = "Save: There is nothing to save"
        Exit Sub
    End If
    
    '-- Was a filename specified?
    If Len(szFileName) = 0 Then
        m_lErrorNum = vbObjectError + ERROR_NO_FILE_SPECIFIED
        m_sErrorMsg = "Save: No File Specified"
        Exit Sub
    End If
    
    '-- Clear the errors
    m_lErrorNum = 0
    m_sErrorMsg = ""
    
    '-- Open the file
    lFileNum = FreeFile
    On Error Resume Next
    Open szFileName For Binary As lFileNum
    If Err Then
        m_lErrorNum = vbObjectError + ERROR_FILE_WRITE
        m_sErrorMsg = "Save: Could Not Write File " & szFileName
    End If

    '-- Get the buffer size
    If m_lBlockSize = 0 Then
        m_lBlockSize = 32768
    ElseIf m_lBlockSize > 65535 Then
        m_lBlockSize = 65535
    End If
    
    '-- Get the number of blocks
    lNumBlocks = lFileLen \ m_lBlockSize
    
    '-- Process each block
    For lIndex = 1 To lLines
        '-- make sure to add the CR/LF
        szBuffer = szBuffer & szData(lIndex) & vbCrLf
        If Len(szBuffer) >= m_lBlockSize Then
            Put #lFileNum, , szBuffer
            szBuffer = ""
        End If
    Next
    
    '-- Process whatever's left
    If Len(szBuffer) Then
        Put #lFileNum, , szBuffer
    End If

    '-- Close the file and exit
    Close lFileNum

End Sub


Public Function FindFirst(ByVal szText As String, ByVal nCase As Integer) As Long
'-- Searches for text in the file and returns the
'   line number where a match was found.

    Dim lIndex      As Long
    Dim lPos        As Long
    
    lCurPos = 0
    nSearchCase = nCase

    '-- Set case-sensitivity
    If nCase Then
        szSearch = szText
    Else
        szSearch = UCase$(szText)
    End If

    '-- Search from the top
    For lIndex = 1 To lLines
        If nCase Then
            lPos = InStr(szData(lIndex), szSearch)
        Else
            lPos = InStr(UCase$(szData(lIndex)), szSearch)
        End If
        If lPos Then
            lCurLine = lIndex
            FindFirst = lCurLine
            lCurPos = lPos
            Exit For
        End If
    Next

End Function


Public Function FindNext() As Long
'-- Searches for text in the file and returns the
'   line number where a match was found.

    Dim lIndex      As Long
    Dim szTemp      As String
    Dim lPos        As Long
    

    '-- Search from the current line
    For lIndex = lCurLine + 1 To lLines
        '-- If we're in the middle of a line,
        '   search from the end of the last find.
        If lCurPos Then
            szTemp = Mid$(szData(lIndex), lCurPos + Len(szSearch))
        Else
            '-- Normal search
            szTemp = szData(lIndex)
        End If
        
        '-- Search
        If nSearchCase Then
            lPos = InStr(szTemp, szSearch)
        Else
            lPos = InStr(UCase$(szTemp), szSearch)
        End If
        If lPos Then
            '-- Set the current line
            lCurLine = lIndex
            
            '-- Set the current position within the line
            If lCurPos Then
                lCurPos = lCurPos + Len(szSearch) + lPos
            Else
                lCurPos = lPos
            End If
            
            '-- Return the current line
            FindNext = lCurLine
            Exit Function
        Else
            '-- No match
            'szSearch = ""
            'lCurLine = 0
            lCurPos = 0
        End If
    Next

End Function


Public Property Get BlockSize() As Long
    BlockSize = m_lBlockSize
End Property

Public Property Let BlockSize(Value As Long)
    If Value <= 0 Then
        m_lBlockSize = 1024
    ElseIf Value > 65535 Then
        m_lBlockSize = 65535
    Else
        m_lBlockSize = Value
    End If
End Property


Public Property Get ErrorNum() As Long
    ErrorNum = m_lErrorNum
End Property

Public Property Get ErrorMsg() As String
    ErrorMsg = m_sErrorMsg
End Property


Public Property Get Lines() As Long
    Lines = lLines
End Property


Public Property Get Line(ByVal lIndex As Long) As String

'-- Retreives a line of text from the file.

    '-- Trap errors
    On Error Resume Next
    Line = szData(lIndex)

End Property

'Public Property Let Line(ByVal lIndex As Long, ByVal szText As String)
Public Sub SetLine(ByVal lIndex As Long, ByVal szText As String)

    '-- Trim Cr and LF chars
    If Right$(szText, 1) = vbLf Then
        szText = Left$(szText, Len(szText) - 1)
    End If
    
    If Right$(szText, 1) = vbCr Then
        szText = Left$(szText, Len(szText) - 1)
    End If
    
    '-- Return the array element (no CR/LF)
    szData(lIndex) = szText

End Sub


Public Property Get FoundPos() As Integer

'-- Return the current position within the current line of
'   the searched for and found text.
    
    FoundPos = lCurPos
    
End Property

