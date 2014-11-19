Attribute VB_Name = "AppSupport"
Option Explicit

Public Const APP_ALIAS As String = "EventsMonitor"
Public Const DEFAULT_CAPTION As String = "SyncFile " & APP_ALIAS
Public Const COPY_LOG_TABLE As String = "daq_Dataworxs.tbl_Dw2DTS_CopyLog"

Public INI_FILE As String
Public APP_LOG As String
Public JOB_LOG As String

Public oCnDwLog As ADODB.Connection, oCnDtsDAQ As ADODB.Connection

Public bInIDE As Boolean, bCloseApp As Boolean
Private m_bEndAppCalled As Boolean


Sub Main()

    On Error GoTo Main_EH

    Dim sDbProfile As String
    Dim bRet As Boolean, sRet As String
    
100     bInIDE = InIDE
        
102     INI_FILE = AddSlash(App.Path) & "SyncFileAgents.ini"
104     APP_LOG = AddSlash(App.Path) & "Logs\" & App.EXEName & " - " & Format$(Now(), "yyyy-mm-dd") & ".log"
106     JOB_LOG = AddSlash(App.Path) & "Logs\" & App.EXEName & " - " & Format$(Now(), "yyyy-mm-dd") & ".csv"
    

108     Load frmMain
110     frmMain.Show
112     DoEvents

    
114     frmMain.UpdateStatus "", True
116     frmMain.UpdateStatus "", True
118     frmMain.UpdateStatus "Loading " & App.Title & " ...", True
        

120     sDbProfile = GetINISetting("Databases", "DataCollector", "", INI_FILE)
122     If sDbProfile <> "" Then
124         frmMain.UpdateStatus "  Connecting to " & sDbProfile & " database ..."
126         bRet = ConnectDb(oCnDtsDAQ, sDbProfile, sRet)
128         If bRet Then
130             frmMain.UpdateStatus "  Connected to " & sDbProfile & " database OK"
        
            Else
132             sRet = "Failed to connect to " & sDbProfile & " database" & IIf(sRet <> "", ": " & sRet, "")
134             GoTo Main_End
            
            End If
    
        Else
136         bRet = False
138         sRet = "The DataCollector database profile is not specified."
    
        End If
    
    
140     sDbProfile = GetINISetting("Databases", "DataworxsLog", "", INI_FILE)
142     If sDbProfile <> "" Then
144         frmMain.UpdateStatus "  Connecting to " & sDbProfile & " database ..."
146         bRet = ConnectDb(oCnDwLog, sDbProfile, sRet)
148         If bRet Then
150             frmMain.UpdateStatus "  Connected to " & sDbProfile & " database OK"
                
            Else
152             sRet = "Failed to connect to " & sDbProfile & " database" & IIf(sRet <> "", ": " & sRet, "")
154             GoTo Main_End
            
            End If
    
        Else
156         bRet = False
158         sRet = "The DataworxsLog database profile is not specified."
    
        End If
    
    
Main_End:

160     If bRet Then
162         frmMain.UpdateStatus "Loaded " & App.Title & " OK", True
164         frmMain.UpdateStatus "", True
166         frmMain.moptAutoRun.Checked = (Not bInIDE)
168         frmMain.tmrTimer.Enabled = True
    
        Else
170         frmMain.UpdateStatus "Failed to load " & App.Title & IIf(sRet <> "", ": " & sRet, ""), True
172         MsgBox sRet & vbCrLf & "Closing " & App.Title, vbCritical Or vbOKOnly, "A critical error has occurred ..."
174         EndApp
    
        End If

        Exit Sub
    

Main_EH:

176     MsgBox "In Main()" & IIf(Erl <> 0, " at line " & Erl, "") & vbCrLf & vbCrLf & _
               vbTab & "Error [" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & _
               "Please forward to technical support.", vbExclamation Or vbOKOnly, "An unexpected error has occurred ..."
    
178     End
    
End Sub


Private Function ConnectDb(TargetDb As ADODB.Connection, DbProfile As String, Optional Results As String) As Boolean

    On Error GoTo ConnectDb_EH

    Dim sCN As String
    Dim lCtr As Long
    Dim sRet As String, bRet As Boolean

        'frmMain.UpdateStatus "Connecting to " & DbProfile & " database ...", True

100     sCN = BuildDbConnectionString(GetINISetting(DbProfile, "Provider", "MYSQL ODBC 3.51 DRIVER", INI_FILE), _
                                      GetINISetting(DbProfile, "Server", "", INI_FILE), _
                                      GetINISetting(DbProfile, "Database", "", INI_FILE), _
                                      CBoolEx(GetINISetting(DbProfile, "UseNTSecurity", "False", INI_FILE), False), _
                                      GetINISetting(DbProfile, "Username", "", INI_FILE), _
                                      GetINISetting(DbProfile, "Password", "", INI_FILE), _
                                      CLng(Val(GetINISetting(DbProfile, "Timeout", "30", INI_FILE))))
        'Debug.Print sCN

102     Set TargetDb = New ADODB.Connection

104     TargetDb.Open sCN

        'Debug.Print TargetDb.ConnectionString
        
106     If TargetDb.State = adStateOpen Then
            'frmMain.UpdateStatus "Connected OK", True
            'frmMain.UpdateStatus

        Else
108         sRet = IIf(sRet <> "", sRet & vbCrLf, "") & "Failed to connect to database"
110         For lCtr = 0 To TargetDb.Errors.Count - 1
112             sRet = sRet & vbCrLf & vbTab & TargetDb.Errors(lCtr).Number & ": " & TargetDb.Errors(lCtr).Description
            Next
        
        End If

114     bRet = (sRet = "")
    
ConnectDb_End:

116     Results = sRet
    
118     ConnectDb = bRet
        
        Exit Function
    

ConnectDb_EH:

120     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in ConnectDb(" & DbProfile & ")" & IIf(Erl <> 0, " at line " & Erl, "")

122     bRet = False

124     Resume ConnectDb_End

End Function


Private Sub DisconnectDb(TargetDb As ADODB.Connection, Optional Results As String = "")

    On Error GoTo DisconnectDb_EH

    Dim sRet As String

100     If Not TargetDb Is Nothing Then
102         If TargetDb.State <> adStateClosed Then TargetDb.Close
104         Set TargetDb = Nothing
        End If
    
DisconnectDb_End:

        On Error Resume Next
        
106     Results = sRet
        
        Exit Sub
    

DisconnectDb_EH:

108     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in DisconnectDb()" & IIf(Erl <> 0, " at line " & Erl, "")

        On Error Resume Next
        
110     If IsFormLoaded("frmMain") Then
112         frmMain.UpdateStatus sRet, True
        
        End If
    
114     Resume DisconnectDb_End
        
End Sub


Public Sub EndApp()

On Error Resume Next

Dim oFrm As Form, sRet As String

    If m_bEndAppCalled Then
        Exit Sub
    
    Else
        m_bEndAppCalled = True
    
    End If

    For Each oFrm In Forms
        Unload oFrm
        Set oFrm = Nothing
    
    Next
    
    DisconnectDb oCnDtsDAQ, sRet
    If sRet <> "" Then Debug.Print "DisconnectDb(oCnDtsDAQ) exited with error: " & sRet

    DisconnectDb oCnDwLog, sRet
    If sRet <> "" Then Debug.Print "DisconnectDb(oCnDwLog) exited with error: " & sRet

    End
    
End Sub


Public Function LoadSQLFile(FileName As String, Optional Results As String) As String

    On Error GoTo LoadSQLFile_EH

    Dim iFileNum As Integer, bFileOpen As Boolean, aryBytes() As Byte
    Dim sSQL As String
    Dim sRet As String

100     iFileNum = FreeFile
102     Open FileName For Binary As iFileNum
104     bFileOpen = True
        
106     ReDim aryBytes(1 To LOF(iFileNum))
108     Get iFileNum, , aryBytes()

110     Close iFileNum
112     bFileOpen = False

114     sSQL = StrConv(aryBytes, vbUnicode)

LoadSQLFile_End:

        On Error Resume Next
    
116     If bFileOpen Then
118         Close iFileNum
120         bFileOpen = False
        End If
    
122     Results = sRet
    
124     LoadSQLFile = sSQL
    
        Exit Function
    
    
LoadSQLFile_EH:

126     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in LoadSQLFile()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
128     sSQL = ""
    
130     Resume LoadSQLFile_End
    
End Function

