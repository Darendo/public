VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   Caption         =   "SyncFile Event Monitor"
   ClientHeight    =   4590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   4260
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16960
            Text            =   "Idle"
            TextSave        =   "Idle"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1755
      Top             =   1260
   End
   Begin VB.TextBox txtOutput 
      Height          =   3300
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   225
      Width           =   7395
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mcmdRunNow 
         Caption         =   "&Run now"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mcmdResend 
         Caption         =   "Resend ..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mcmdExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu moptAutoRun 
         Caption         =   "&Auto-run"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOptionBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mcmdSettings 
         Caption         =   "&Settings"
         Enabled         =   0   'False
      End
      Begin VB.Menu moptDebugMode 
         Caption         =   "&Debug mode"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lFrequency As Long, m_lTimerCtr As Long
Private m_bBusy As Boolean

Private Sub Form_Load()

    On Error Resume Next

    Dim sRet As String

100     m_lFrequency = 60
102     sRet = GetINISetting("General", "Frequency", CStr(m_lFrequency), INI_FILE)
104     If IsNumeric(sRet) Then
106         m_lFrequency = CLng(sRet)
        End If

108     If FileExists(APP_LOG) Then LoadLog APP_LOG

End Sub

Private Sub Form_Resize()

On Error Resume Next

    If WindowState <> vbMinimized Then
        txtOutput.Move ScaleLeft, ScaleTop, ScaleWidth, ScaleHeight - (ScaleTop + StatusBar1.Height)
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next

    bCloseApp = True
    
    If m_bBusy Then
        Cancel = 1
        If moptDebugMode.Checked Then
            UpdateStatus "Form_QueryUnload() event fired while busy. Waiting to cancel current operation ...", True
        Else
            Visible = False
        End If
    End If
    
End Sub

Private Sub mcmdExit_Click()
    
On Error Resume Next

    mcmdRunNow.Enabled = False
    moptAutoRun.Enabled = False
    moptAutoRun.Checked = False
    DoEvents
    
    If Not m_bBusy Then
        EndApp
    
    Else
        bCloseApp = True

    End If

End Sub

Private Sub mcmdRunNow_Click()
    On Error Resume Next
    If Not m_bBusy Then m_lTimerCtr = (m_lFrequency + 1)
End Sub

Private Sub mcmdResend_Click()

On Error Resume Next

Dim sRet As String, bRet As Boolean

    sRet = InputBox("Enter ReportID:", "Queue report for resend ...", "")
    If sRet <> "" Then AppendString AddSlash(App.Path) & "Resends.txt", sRet & "|||" & sRet & ".doc", False

End Sub

Private Sub mcmdSettings_Click()
    
On Error Resume Next

    Load frmSettings
    
    frmSettings.Frequency = m_lFrequency
    
    frmSettings.Show vbModal, Me
    
    If Not frmSettings.Cancel Then
        
        m_lFrequency = frmSettings.Frequency
        SaveINISetting "General", "Frequency", CStr(m_lFrequency), INI_FILE
    
    End If
    
    Unload frmSettings
    Set frmSettings = Nothing

End Sub

Private Sub moptAutorun_Click()
On Error Resume Next
    moptAutoRun.Checked = (Not moptAutoRun.Checked)
    m_lTimerCtr = 0
End Sub

Private Sub moptDebugMode_Click()
On Error Resume Next
    moptDebugMode.Checked = (Not moptDebugMode.Checked)
End Sub

Private Sub tmrTimer_Timer()
'Private Sub MySub()

    On Error GoTo tmrTimer_Timer_EH
    
    Dim lRet As Long, sRet As String

100     tmrTimer.Enabled = False
    
102     If bCloseApp Then
104         EndApp
            Exit Sub
        
        End If
            
106     If Not IsFormLoaded("frmSettings") Then
    
108         m_lTimerCtr = m_lTimerCtr + 1
        
110         If moptAutoRun.Checked Then
        
112             lRet = m_lFrequency - m_lTimerCtr
            
114             If lRet <= 0 Then
116                 SyncFiles
117                 HandleFilesNotFound
118                 m_lTimerCtr = 0
                
                Else
120                 UpdateStatus "Idle - " & lRet & " second" & IIf(lRet = 1, "", "s") & " to run ...", False
            
                End If
    
            Else
        
122             lRet = (5 * IIf(bInIDE, 1, 60)) - m_lTimerCtr
            
124             If lRet > 0 Then
126                 UpdateStatus "Paused - Auto-resume in " & lRet & " second" & IIf(lRet = 1, "", "s") & " ...", False
                
                Else
128                 moptAutoRun.Checked = True
130                 m_lTimerCtr = 0

                End If
        
            End If
        
        Else
            ' Ignore for now

        End If

tmrTimer_Timer_End:

        On Error Resume Next

132     tmrTimer.Enabled = True

        Exit Sub
    

tmrTimer_Timer_EH:

134     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in frmMain.tmrTimer_Timer()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next

136     UpdateStatus sRet, True

138     Resume tmrTimer_Timer_End

End Sub

Private Sub SyncFiles()

    On Error GoTo SyncFiles_EH

    Dim aryProfiles() As String, lFreq As Long, dtmNextRun As Date
    Dim lFilesQueued As Long, lFilesFailed As Long
    Dim sRet As String, bRet As Boolean, lRet As Long, lCtr As Long, sglRet As Single

100     m_bBusy = True
102     EnableUI False
    
104     sRet = GetINISetting(APP_ALIAS, "Profiles", "", INI_FILE)
    
106     If sRet <> "" Then
    
108         If InStr(1, sRet, ";") > 0 Then
110             aryProfiles() = Split(sRet, ";")
        
            Else
112             ReDim aryProfiles(0)
114             aryProfiles(0) = sRet
            
            End If
        
116         For lCtr = 0 To UBound(aryProfiles())
        
118             lFreq = 0
120             sRet = GetINISetting(aryProfiles(lCtr), "Frequency", CStr(lFreq), INI_FILE)
122             If IsNumeric(sRet) Then lFreq = Abs(CLng(sRet))
            
124             dtmNextRun = CDate("12:00:00 AM")
126             sRet = GetINISetting(aryProfiles(lCtr), "NextRun", CStr(dtmNextRun), INI_FILE)
128             If IsDate(sRet) Then dtmNextRun = CDate(sRet)
130             sglRet = DateDiff("s", Now(), dtmNextRun)
            
132             If lFreq > 0 And sglRet <= 0 Then
            
134                 UpdateStatus "Scanning " & aryProfiles(lCtr) & " events ...", True
            
136                 bRet = ScanEvents(aryProfiles(lCtr), lFilesQueued, lFilesFailed, sRet)
            
138                 UpdateStatus "  Queued " & Format$(lFilesQueued, "#,##0") & " file" & IIf(lFilesQueued = 1, "", "s") & " OK", True
                    'UpdateStatus "  Failed to queue " & Format$(lFilesFailed, "#,##0") & " file" & IIf(lFilesFailed = 1, "", "s"), True
            
140                 If bRet Then
142                     UpdateStatus "Completed scanning " & aryProfiles(lCtr) & " events " & IIf(sRet <> "", "with comments: " & sRet, "OK"), True
144                     dtmNextRun = DateAdd("n", lFreq, dtmNextRun)
146                     SaveINISetting aryProfiles(lCtr), "NextRun", CStr(dtmNextRun), INI_FILE
                
                    Else
148                     UpdateStatus "Failed to scan " & aryProfiles(lCtr) & " events" & IIf(sRet <> "", ": " & sRet, ""), True
            
                    End If
                
150                 UpdateStatus "", True

                End If
            
            Next
        
        Else
152         UpdateStatus "WARNING: No profiles set", True
    
        End If
    

SyncFiles_End:

        On Error Resume Next

154     UpdateStatus "Idle", False
    
156     EnableUI True
158     m_bBusy = False

        Exit Sub
    

SyncFiles_EH:

160     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in frmMain.SyncFiles()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next

162     UpdateStatus sRet, True

164     Resume SyncFiles_End

End Sub

Private Function ScanEvents(EventProfile As String, Optional FilesQueued As Long, Optional FilesFailed As Long, Optional Results As String = "") As Boolean

    On Error GoTo ScanEvents_EH

    Dim aryProfiles() As String, lNumProfiles As Long, pCtr As Long, sProfileINI As String, sSourceFolder As String, lDateModifiedOffset As Long
    Dim dtmOffset As Long, dtmLastRun As Date, dtmNextRun As Date, lNumEvents As Long, lEvent As Long

    Dim oRS As ADODB.Recordset, sSqlFile As String, strSQL As String

    Dim udtFileInfo As FILE_INFO, sReportID As String, sFileName As String, sFileInfo As String
    Dim aryFiles() As FILE_INFO, lFilesFound As Long

    Dim sRet As String, bRet As Boolean, lCtr As Long, sTmp As String, lRet As Long, bCancel As Boolean, sglStart As Single, sglEnd As Single

        'sTmp = GetINISetting("General", "TimeZoneOffset", "", INI_FILE)
        'If IsNumeric(sTmp) Then
        '    dtmOffset = CLng(sTmp)
        'Else
        '    dtmOffset = 0
        'End If
100     dtmOffset = 0   ' <-- This is now abandoned in favor of letting MySQL translate to GMT
    
102     sTmp = GetINISetting(EventProfile, "LastRun", "", INI_FILE)
104     If IsDate(sTmp) Then
106         dtmLastRun = CDate(sTmp)
108         If dtmOffset <> 0 Then
110             dtmLastRun = DateAdd("h", dtmOffset, dtmLastRun)
            End If
112         UpdateStatus "  Querying for events since " & dtmLastRun
    
        Else
114         sRet = "Invalid date for last run"
116         GoTo ScanEvents_End

        End If
    
        'strSQL = EVENT_QUERY
        'strSQL = Replace$(strSQL, "#LastQueryDateTime#", Format$(dtmLastRun, "yyyy-mm-dd hh:nn:ss"))
118     sSqlFile = GetINISetting(EventProfile, "Query", "", INI_FILE)
120     If sSqlFile <> "" Then
122         If InStr(1, sSqlFile, "\\") = 0 And InStr(1, sSqlFile, ":") = 0 Then
124             sSqlFile = AddSlash(App.Path) & sSqlFile
            End If
126         strSQL = LoadSQLFile(sSqlFile, sRet)
128         If strSQL <> "" Then
130             strSQL = Replace$(strSQL, "#CopyLogTable#", COPY_LOG_TABLE)
132             strSQL = Replace$(strSQL, "#CurrentDateTime#", Format$(Now(), "yyyy-mm-dd hh:nn:ss"))
134             strSQL = Replace$(strSQL, "#LastQueryDateTime#", Format$(dtmLastRun, "yyyy-mm-dd hh:nn:ss"))
                'Debug.Print "strSQL = " & strSQL
            Else
136             sRet = "Failed to load SQL query from " & GetFileName(sSqlFile) & IIf(sRet <> "", ": " & sRet, "")
138             GoTo ScanEvents_End
            End If
        Else
140         sRet = "SQL query file not specified"
142         GoTo ScanEvents_End
        End If
        'UpdateStatus "strSQL = " & strSQL, True

    
144     dtmNextRun = Now()

146     UpdateStatus "  Opening recordset ...", False
148     Set oRS = OpenRS(oCnDwLog, strSQL, , , , sRet)
    
150     If Not oRS Is Nothing Then
        
152         UpdateStatus "  Recordset opened OK", False
154         strSQL = ""

156         With oRS
            
158             If Not .EOF Then
            
                    '.MoveLast
                    'lNumEvents = .RecordCount
                    '.MoveFirst
    
160                 Do While Not .EOF
                        
162                     If (Not IsNull(!ReportID)) Then
164                         sReportID = !ReportID
                        Else
166                         sReportID = "Null"
                        End If
                    
168                     If (Not IsNull(!FileName)) Then
170                         sFileName = !FileName
                        Else
172                         sFileName = "Null"
                        End If
                        
174                     If sReportID <> "Null" And sFileName <> "Null" Then
                    
                            'lEvent = lEvent + 1
                            'UpdateStatus "  Analyzing event " & Format$(lEvent, "#,##0") & ": " & sReportID & " ... ", False, bCancel
176                         UpdateStatus "  Analyzing " & sReportID & " ...", False, bCancel

178                         sRet = GetINISetting("General", "SourceProfiles", "", INI_FILE)
180                         If sRet <> "" Then
182                             If InStr(1, sRet, ";") > 0 Then
184                                 aryProfiles() = Split(sRet, ";")
186                                 lNumProfiles = UBound(aryProfiles()) + 1
                                Else
188                                 ReDim aryProfiles(0)
190                                 aryProfiles(0) = sRet
192                                 lNumProfiles = 1
                                End If
                            Else
194                             lNumProfiles = 0
                            End If

196                         For pCtr = 0 To (lNumProfiles - 1)

198                             sRet = ""

200                             UpdateStatus "  Analyzing " & sReportID & " in " & aryProfiles(pCtr) & " ...", False, bCancel

202                             sProfileINI = AddSlash(App.Path) & aryProfiles(pCtr) & ".ini"
204                             sSourceFolder = GetINISetting("Locations", "SourcePattern", "", sProfileINI)
206                             If sSourceFolder = "" Then GoTo Skip_File_Check
208                             lDateModifiedOffset = 10
210                             sTmp = GetINISetting("Options", "DateModifiedOffset", CStr(lDateModifiedOffset), sProfileINI)
212                             If IsNumeric(sTmp) Then
214                                 lDateModifiedOffset = CLng(sTmp)
                                End If

216                             udtFileInfo.FileName = sFileName
218                             udtFileInfo.FilePath = sSourceFolder
220                             udtFileInfo.FileTimeLastModified = "12:00:00"
222                             udtFileInfo.FileSize = 0
224                             sRet = ""
                            
226                             If (InStr(1, udtFileInfo.FilePath, "#TranGroup#") > 0 And IsNull(!TranGroup)) Then
228                                 sRet = "TranGroup expected but missing"
                                End If
230                             If (InStr(1, udtFileInfo.FilePath, "#TranUser#") > 0 And IsNull(!TranUser)) Then
232                                 sRet = IIf(sRet <> "", sRet & "; ", "") & "TranUser expected but missing"
                                End If
234                             If (InStr(1, udtFileInfo.FilePath, "#SubFolder#") > 0 And IsNull(!SubFolder)) Then
236                                 sRet = IIf(sRet <> "", sRet & "; ", "") & "SubFolder expected but missing"
                                End If
238                             If sRet <> "" Then GoTo Skip_File_Check
                            
240                             udtFileInfo.FileName = sFileName
                                'udtFileInfo.FilePath = sSourceFolder
242                             udtFileInfo.FilePath = Replace$(udtFileInfo.FilePath, "#TranGroup#", !TranGroup)
244                             udtFileInfo.FilePath = Replace$(udtFileInfo.FilePath, "#TranUser#", !TranUser)
246                             udtFileInfo.FilePath = Replace$(udtFileInfo.FilePath, "#SubFolder#", !SubFolder)
                            
248                             If FolderExists(udtFileInfo.FilePath) Then
                            
250                                 If FileExists(udtFileInfo.FilePath & udtFileInfo.FileName) Then
252                                     udtFileInfo.FileTimeLastModified = FileDateTime(udtFileInfo.FilePath & udtFileInfo.FileName)
254                                     udtFileInfo.FileSize = FileLen(udtFileInfo.FilePath & udtFileInfo.FileName)
                                
                                    Else
'255                                     UpdateStatus "  File not found: " & udtFileInfo.FileName & " in " & udtFileInfo.FilePath, True
255                                     sRet = "File not found: " & udtFileInfo.FileName & " in " & udtFileInfo.FilePath
                                        
                                    End If
                                
                                Else
'256                                 UpdateStatus "  Folder not found: " & udtFileInfo.FilePath, True
256                                 sRet = "Folder not found: " & udtFileInfo.FilePath

                                End If
        
257                             If udtFileInfo.FileSize > 0 Then

                                    ''strSQL = "INSERT INTO tbl_Dw2DTS_CopyLog_EM (ReportID, Profile, SourceFolder, FileName, FileSize, DateFileModified, DateFileFound, DateFileCopied) VALUES " & _
                                    '         "('" & sReportID & "', '" & GetProfileFromFolder(udtFileInfo.FilePath) & "', '" & Replace$(udtFileInfo.FilePath, "\", "\\") & "', '" & udtFileInfo.FileName & "', " & udtFileInfo.FileSize & ", '" & Format$(udtFileInfo.FileTimeLastModified, "yyyy-mm-dd hh:nn:ss") & "', '" & Format$(DateAdd("h", -dtmOffset, !LastEvent), "yyyy-mm-dd hh:nn:ss") & "', '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "');"
                                    'strSQL = "INSERT INTO " & COPY_LOG_TABLE & " (ReportID, Profile, SourceFolder, FileName, FileSize, DateFileModified, DateFileFound, DateFileCopied, DateFileToCopy) VALUES " & _
                                    '         "('" & sReportID & "', '" & aryProfiles(pCtr) & "', '" & Replace$(udtFileInfo.FilePath, "\", "\\") & "'," & _
                                    '         " '" & udtFileInfo.FileName & "', " & udtFileInfo.FileSize & ", '" & Format$(udtFileInfo.FileTimeLastModified, "yyyy-mm-dd hh:nn:ss") & "'," & _
                                    '         " '" & Format$(DateAdd("h", -dtmOffset, !LastEvent), "yyyy-mm-dd hh:nn:ss") & "', '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "'," & _
                                    '         " '" & Format$(DateAdd("n", lDateModifiedOffset, Now()), "yyyy-mm-dd hh:nn:ss") & "');"
                                    'Debug.Print "strSQL = " & strSQL
                                    '
                                    'oCnDtsDAQ.Execute strSQL, lRet
                                    '
                                    'If lRet <> 1 Then
                                    '    UpdateStatus "  Failed to execute SQL command: " & strSQL, True, bCancel
                                    '
                                    'Else
                                    '    UpdateStatus "  Queued " & sReportID & " for copy", True, bCancel
                                    '
                                    'End If
                                    'strSQL = ""
    
    
258                                 lFilesFound = lFilesFound + 1
                                    'ReDim Preserve aryFiles(lFilesFound - 1)
                                    'aryFiles(lFilesFound - 1) = udtFileInfo
260                                 ReDim aryFiles(0)
262                                 aryFiles(0) = udtFileInfo

                                
264                                 UpdateStatus "  Analyzing " & sReportID & " in " & aryProfiles(pCtr) & " ....", False, bCancel
                                
266                                 sglStart = Timer
                                    'bRet = SaveListToDb(aryProfiles(pCtr), aryFiles(), lFilesFound, lRet, lDateModifiedOffset, GetResends, sRet)
268                                 bRet = SaveListToDb(aryProfiles(pCtr), aryFiles(), 1, lRet, lDateModifiedOffset, sRet)
270                                 sglEnd = Timer
                                
272                                 FilesQueued = FilesQueued + lRet
                                
274                                 If bRet Then
                        
276                                     If sglEnd >= sglStart Then
278                                         sglEnd = sglEnd - sglStart
                                        Else
280                                         sglEnd = 0
                                        End If
                                    
282                                     If lRet = 1 Then
                                            'UpdateStatus "  Queued " & sReportID & " OK in " & Format$(sglEnd, "#,##0.000") & " second" & IIf(sglEnd = 1, "", "s"), False, bCancel
284                                         UpdateStatus "  " & sRet & " " & sReportID & " in " & aryProfiles(pCtr) & " OK", False, bCancel
                                        
                                        Else
286                                         UpdateStatus "  " & sReportID & " in " & aryProfiles(pCtr) & " already queued", False, bCancel
                                        
                                        End If
287                                     sRet = ""

288                                     If bCancel Then
290                                         sRet = "Operation cancelled by user"
292                                         GoTo ScanEvents_End
                                        End If
                        
                                    Else
294                                     sRet = "  FAILED to queue " & sReportID & " in " & aryProfiles(pCtr) & IIf(sRet <> "", ": " & sRet, "")
296                                     GoTo Skip_File_Check
                                
                                    End If

298                                 If moptDebugMode.Checked Then
300                                     sFileInfo = "'" & sReportID & "','" & udtFileInfo.FilePath & "','" & udtFileInfo.FileName & "'," & _
                                                    udtFileInfo.FileSize & "," & IIf(IsDate(udtFileInfo.FileTimeLastModified), "'" & Format$(udtFileInfo.FileTimeLastModified, "yyyy-mm-dd hh:nn:ss") & "'", "Null") & "," & _
                                                    "'" & DateAdd("h", -dtmOffset, !LastEvent) & "','" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "'," & CStr(lRet)
301                                     AppendString JOB_LOG, sFileInfo, False
                                    End If

                                Else
                                    ' At this point, we will assume the situation has been handled/logged.
                                    'UpdateStatus "  Analyzing " & sReportID & " in " & aryProfiles(pCtr) & " ....", False, bCancel
302                                 sTmp = sRet
303                                 sRet = ""
304                                 bRet = FileNotFound(aryProfiles(pCtr), udtFileInfo, sTmp, sRet)

                                End If

Skip_File_Check:
                        
305                             If sRet <> "" Then
306                                 UpdateStatus "  Unable to process " & sReportID & " from " & aryProfiles(pCtr) & " profile: " & sRet, True, bCancel

                                End If
                            
                            Next
                    
                        Else
308                         UpdateStatus "  Unable to process event; ReportID = '" & sReportID & "' and FileName = '" & sFileName & "'", True, bCancel
    
                        End If
                    
310                     If bCancel Then
312                         sRet = "Operation cancelled by user"
314                         GoTo ScanEvents_End
                        End If
                    
316                     .MoveNext
                
                    Loop
            
                Else
318                 sRet = "No events found"

                End If
            
320             UpdateStatus "  Closing recordset ...", False
            
322             .Close
        
            End With

324         SaveINISetting EventProfile, "LastRun", dtmNextRun, INI_FILE
        
326         bRet = True

        Else
328         bRet = False
330         sRet = "Failed to open the recordset" & IIf(sRet <> "", ": " & sRet, "")

        End If
    
ScanEvents_End:

        On Error Resume Next
    
332     If Not oRS Is Nothing Then
334         If oRS.State <> adStateClosed Then oRS.Close
336         Set oRS = Nothing
        End If
    
338     Results = sRet
    
340     ScanEvents = bRet
    
        Exit Function


ScanEvents_EH:

342     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in ScanEvents()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
        
344     If strSQL <> "" Then sRet = sRet & vbCrLf & "strSQL = " & strSQL
        
346     bRet = False
    
348     Resume ScanEvents_End
    
End Function

Private Function HandleFilesNotFound(Optional Results As String = "") As Boolean

On Error GoTo HandleFilesNotFound_EH

Dim oRS As ADODB.Recordset, strSQL As String
Dim sSourceFolder As String, udtFileInfo As FILE_INFO, sReportID As String, sFileName As String, sFileInfo As String, sComments As String
Dim sRet As String, bRet As Boolean, lCtr As Long, sTmp As String, lRet As Long, bCancel As Boolean, sglStart As Single, sglEnd As Single

    strSQL = "SELECT * FROM " & COPY_LOG_TABLE & " WHERE DateFileFound >= '" & Format$(DateAdd("m", -1, Now()), "yyyy-mm-dd hh:nn:ss") & "' AND (CopiedOK = 0) AND (Comments LIKE 'File not found%') ORDER BY DateFileFound ASC;"
    'UpdateStatus "strSQL = " & strSQL, True

    UpdateStatus "Handling files not found ...", True

    UpdateStatus "  Opening recordset ...", False
    Set oRS = New ADODB.Recordset
    
    If Not oRS Is Nothing Then
        
        With oRS
            
            .Open strSQL, oCnDtsDAQ, adOpenStatic, adLockOptimistic, adCmdText
    
            If Not .EOF Then
            
                UpdateStatus "  Recordset opened OK", False
                strSQL = ""
        
                Do While Not .EOF
                        
                    If (Not IsNull(!ReportID)) Then
                        sReportID = !ReportID
                    Else
                        sReportID = "Null"
                    End If
                    
                    If (Not IsNull(!FileName)) Then
                        sFileName = !FileName
                    Else
                        sFileName = "Null"
                    End If
                        
                    If Not IsNull(!SourceFolder) Then
                        sSourceFolder = !SourceFolder
                    Else
                        sSourceFolder = "Null"
                    End If
                    
                    If sReportID <> "Null" And sFileName <> "Null" And sSourceFolder <> "Null" Then
                    
                        UpdateStatus "  Analyzing " & sReportID & " ...", False, bCancel

                        udtFileInfo.FileName = sFileName
                        udtFileInfo.FilePath = sSourceFolder
                        udtFileInfo.FileTimeLastModified = "12:00:00"
                        udtFileInfo.FileSize = 0
                        sRet = ""
                        
                        If FolderExists(udtFileInfo.FilePath) Then
                        
                            If FileExists(udtFileInfo.FilePath & udtFileInfo.FileName) Then
                                udtFileInfo.FileTimeLastModified = FileDateTime(udtFileInfo.FilePath & udtFileInfo.FileName)
                                udtFileInfo.FileSize = FileLen(udtFileInfo.FilePath & udtFileInfo.FileName)
                            
                            Else
                                'UpdateStatus "  File not found: " & udtFileInfo.FileName & " in " & udtFileInfo.FilePath, True
                                sRet = "File not found: " & udtFileInfo.FileName & " in " & udtFileInfo.FilePath
                                    
                            End If
                            
                        Else
                            'UpdateStatus "  Folder not found: " & udtFileInfo.FilePath, True
                            sRet = "Folder not found: " & udtFileInfo.FilePath

                        End If
    
                        If udtFileInfo.FileSize > 0 Then

                            UpdateStatus "  Requeuing " & sReportID & " ...", False, bCancel
                            
                            sglStart = Timer
                            
                            !Source = APP_ALIAS '& " - " & Profile
                            !DateFileModified = Format$(udtFileInfo.FileTimeLastModified, "yyyy-mm-dd hh:nn:ss")
                            !FileSize = udtFileInfo.FileSize
                            !DateFileToCopy = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
                            !DateFileCopied = Null
                            !CopiedOK = Null
                            sComments = "requeued " & Format$(Now(), "yyyymmddhhnnss")
                            If Not IsNull(!Comments) Then
                                sComments = !Comments & "; " & sComments
                                If Len(sComments) > 255 Then sComments = Right$(sComments, 255)
                            End If
                            !Comments = sComments
                            .Update
                            
                            sglEnd = Timer
                            
                            'FilesQueued = FilesQueued + lRet
                            
                            If sglEnd >= sglStart Then
                                sglEnd = sglEnd - sglStart
                            Else
                                sglEnd = 0
                            End If
                            
                            'UpdateStatus "  Queued " & sReportID & " OK in " & Format$(sglEnd, "#,##0.000") & " second" & IIf(sglEnd = 1, "", "s"), False, bCancel
                            UpdateStatus "  " & sReportID & " requeued OK", False, bCancel
                            sRet = ""

                            If bCancel Then
                                sRet = "Operation cancelled by user"
                                GoTo HandleFilesNotFound_End
                            End If

                            If moptDebugMode.Checked Then
                                sFileInfo = "'" & sReportID & "','" & udtFileInfo.FilePath & "','" & udtFileInfo.FileName & "'," & _
                                            udtFileInfo.FileSize & "," & IIf(IsDate(udtFileInfo.FileTimeLastModified), "'" & Format$(udtFileInfo.FileTimeLastModified, "yyyy-mm-dd hh:nn:ss") & "'", "Null") & "," & _
                                            "'','" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "'," & CStr(lRet)
                                AppendString JOB_LOG, sFileInfo, False
                            End If

                        Else
                            ' Already denoted as a "file not found"

                        End If

Skip_File_Check:
                    
                        If sRet <> "" Then
                            UpdateStatus "  Unable to process " & sReportID & ": " & sRet, True, bCancel

                        End If
                    
                    Else
                        UpdateStatus "  Unable to process event; ReportID = '" & sReportID & "' and FileName = '" & sFileName & "'", True, bCancel
    
                    End If
                    
                    If bCancel Then
                        sRet = "Operation cancelled by user"
                        GoTo HandleFilesNotFound_End
                    End If
                    
                    .MoveNext
                
                Loop
            
            Else
                sRet = "No events found"

            End If
            
            UpdateStatus "  Closing recordset ...", False
            
            .Close
        
        End With

        bRet = True

    Else
        bRet = False
        sRet = "Failed to open the recordset" & IIf(sRet <> "", ": " & sRet, "")
        UpdateStatus sRet, True

    End If
    
HandleFilesNotFound_End:

    On Error Resume Next
    
    If Not oRS Is Nothing Then
        If oRS.State <> adStateClosed Then oRS.Close
        Set oRS = Nothing
    End If
    
    Results = sRet

    UpdateStatus "Finished handling files not found" & IIf(sRet <> "", ": " & sRet, ""), True

    HandleFilesNotFound = bRet
    
    Exit Function


HandleFilesNotFound_EH:

    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in HandleFilesNotFound()" & IIf(Erl <> 0, " at line " & Erl, "")
    
    On Error Resume Next
        
    If strSQL <> "" Then sRet = sRet & vbCrLf & "strSQL = " & strSQL
        
    bRet = False
    
    Resume HandleFilesNotFound_End
    
End Function

Private Function SaveListToDb(Profile As String, Files() As FILE_INFO, NumFiles As Long, FilesQueued As Long, DateModifiedOffset As Long, Optional Results As String = "") As Boolean

    On Error GoTo SaveListToDb_EH

    Dim strSQL As String, oRS As ADODB.Recordset
    Dim fCtr As Long, sReportID As String, sComments As String
    Dim sRet As String, bRet As Boolean, lRet As Long, lCtr As Long
    Dim sErrMsg As String

        ' Assume success
100     bRet = True
102     FilesQueued = 0

104     For fCtr = 0 To (NumFiles - 1)
    
106         strSQL = "SELECT * FROM " & COPY_LOG_TABLE & " " & _
                     "WHERE Profile = '" & Profile & "' AND SourceFolder = '" & Replace$(Files(fCtr).FilePath, "\", "\\") & "' AND " & _
                     "      FileName = '" & Replace$(Files(fCtr).FileName, "'", "''") & "';"

108         If oRS Is Nothing Then
110             Set oRS = New ADODB.Recordset
        
            Else
112             If oRS.State <> adStateClosed Then
114                 oRS.Close
                End If
            
            End If
        
116         With oRS
        
118             .Open strSQL, oCnDtsDAQ, adOpenStatic, adLockOptimistic, adCmdText
            
120             If Not .EOF Then
                    'strSQL = "UPDATE " & COPY_LOG_TABLE & " SET DateFileModified = '" & Files(fCtr).FileTimeLastModified & "' WHERE Profile = '" & Profile & "' AND SourceFolder = '" & Files(fCtr).FilePath & "' AND Filename = '" & Files(fCtr).FileName & "';"
                    'Debug.Print !DateFileModified & " <> " & Files(fCtr).FileTimeLastModified & " = " & (!DateFileModified <> Files(fCtr).FileTimeLastModified)
122                 If (((!DateFileModified <> Files(fCtr).FileTimeLastModified) Or IsNull(!DateFileModified)) Or _
                        ((!FileSize <> Files(fCtr).FileSize) Or IsNull(!FileSize))) And _
                        (Not IsNull(!CopiedOK)) Then
124                     !Source = APP_ALIAS '& " - " & Profile
126                     !DateFileModified = Format$(Files(fCtr).FileTimeLastModified, "yyyy-mm-dd hh:nn:ss")
128                     !FileSize = Files(fCtr).FileSize
                        '!DateFileFound = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
130                     !DateFileToCopy = Format$(DateAdd("n", DateModifiedOffset, Now()), "yyyy-mm-dd hh:nn:ss")
132                     !DateFileCopied = Null
134                     !CopiedOK = Null
136                     sComments = "requeued " & Format$(Now(), "yyyymmddhhnnss")
138                     If Not IsNull(!Comments) Then
140                         sComments = !Comments & "; " & sComments
142                         If Len(sComments) > 255 Then sComments = Right$(sComments, 255)
                        End If
144                     !Comments = sComments
146                     .Update
148                     FilesQueued = FilesQueued + 1
149                     sRet = "Requeued"
                    End If

                Else
150                 sReportID = Trim$(GetFileName(Files(fCtr).FileName, False))
152                 If LCase$(Right$(Files(fCtr).FileName, 3)) = "wav" And InStr(1, sReportID, "(0)") > 0 Then
154                     sReportID = Replace$(sReportID, "(0)", "")
                    End If
156                 strSQL = "INSERT INTO " & COPY_LOG_TABLE & " (Source, ReportID, Profile, SourceFolder, Filename, DateFileModified, FileSize, DateFileFound, DateFileToCopy) " & _
                             "VALUES ('" & APP_ALIAS & "', '" & Replace$(sReportID, "'", "''") & "', '" & Profile & "', '" & Replace$(Files(fCtr).FilePath, "\", "\\") & "', " & _
                             "        '" & Replace$(Files(fCtr).FileName, "'", "''") & "', '" & Format$(Files(fCtr).FileTimeLastModified, "yyyy-mm-dd hh:nn:ss") & "', " & Files(fCtr).FileSize & ", " & _
                             "        '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "', '" & Format$(DateAdd("n", DateModifiedOffset, Now()), "yyyy-mm-dd hh:nn:ss") & "');"
158                 oCnDtsDAQ.Execute strSQL, lRet
160                 If lRet = 1 Then
162                     FilesQueued = FilesQueued + 1
163                     sRet = "Queued"

                    Else
164                     bRet = False
166                     sRet = IIf(sRet <> "", sRet & vbCrLf, "") & "Failed to execute SQL statement: " & strSQL
168                     oCnDtsDAQ.Errors.Refresh
170                     If oCnDtsDAQ.Errors.Count > 0 Then
172                         For lCtr = 0 To (oCnDtsDAQ.Errors.Count - 1)
174                             sRet = sRet & vbCrLf & "    Error " & oCnDtsDAQ.Errors(lCtr).Number & ": " & oCnDtsDAQ.Errors(lCtr).Description
                            Next
176                         oCnDtsDAQ.Errors.Clear
                        End If
                    
                    End If

                End If
            
178             .Close

            End With
        
        Next
    
SaveListToDb_End:

        On Error Resume Next
    
180     If Not oRS Is Nothing Then
182         If oRS.State <> adStateClosed Then oRS.Close
184         Set oRS = Nothing
        End If

186     Results = sRet
    
188     SaveListToDb = bRet
    
        Exit Function
    

SaveListToDb_EH:

190     sErrMsg = "Unexpected error occurred in SaveListToDb()" & IIf(Erl <> 0, " at line " & Erl, "") & vbCrLf & _
                  "    Error [" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & _
                  IIf(strSQL <> "", "    SQL = " & strSQL, "")

        On Error Resume Next

192     sRet = IIf(sRet <> "", sRet & vbCrLf, "") & sErrMsg

194     bRet = False
    
196     Resume SaveListToDb_End

End Function

Private Function FileNotFound(Profile As String, TargetFile As FILE_INFO, Comments As String, Optional Results As String = "") As Boolean

    On Error GoTo FileNotFound_EH

    Dim strSQL As String, oRS As ADODB.Recordset
    Dim fCtr As Long, sReportID As String, sComments As String
    Dim sRet As String, bRet As Boolean, lRet As Long, lCtr As Long
    Dim sErrMsg As String

        ' Assume success
100     bRet = True
        
102     strSQL = "SELECT * FROM " & COPY_LOG_TABLE & " " & _
                 "WHERE Profile = '" & Profile & "' AND SourceFolder = '" & Replace$(TargetFile.FilePath, "\", "\\") & "' AND " & _
                 "      FileName = '" & Replace$(TargetFile.FileName, "'", "''") & "';"

104     If oRS Is Nothing Then
106         Set oRS = New ADODB.Recordset
    
        Else
108         If oRS.State <> adStateClosed Then
110             oRS.Close
            End If
        
        End If
    
112     With oRS
    
114         .Open strSQL, oCnDtsDAQ, adOpenStatic, adLockOptimistic, adCmdText
        
116         If Not .EOF Then
                'strSQL = "UPDATE " & COPY_LOG_TABLE & " SET DateFileModified = '" & TargetFile.FileTimeLastModified & "' WHERE Profile = '" & Profile & "' AND SourceFolder = '" & TargetFile.FilePath & "' AND Filename = '" & TargetFile.FileName & "';"
                'Debug.Print !DateFileModified & " <> " & TargetFile.FileTimeLastModified & " = " & (!DateFileModified <> TargetFile.FileTimeLastModified)
118             If Not IsNull(!CopiedOK) Then
120                 !Source = APP_ALIAS '& " - " & Profile
                    '!DateFileModified = Format$(TargetFile.FileTimeLastModified, "yyyy-mm-dd hh:nn:ss")
                    '!FileSize = TargetFile.FileSize
121                 !DateFileFound = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
                    '!DateFileToCopy = Format$(DateAdd("n", DateModifiedOffset, Now()), "yyyy-mm-dd hh:nn:ss")
                    '!DateFileCopied = Null
122                 !CopiedOK = 0
124                 sComments = Comments
126                 If Not IsNull(!Comments) Then
128                     sComments = !Comments & "; " & sComments
130                     If Len(sComments) > 255 Then sComments = Right$(sComments, 255)
                    End If
132                 !Comments = sComments
134                 .Update
                End If

            Else
136             sReportID = Trim$(GetFileName(TargetFile.FileName, False))
138             If LCase$(Right$(TargetFile.FileName, 3)) = "wav" And InStr(1, sReportID, "(0)") > 0 Then
140                 sReportID = Replace$(sReportID, "(0)", "")
                End If
142             strSQL = "INSERT INTO " & COPY_LOG_TABLE & " (Source, ReportID, Profile, SourceFolder, Filename, DateFileFound, CopiedOK, Comments) " & _
                         "VALUES ('" & APP_ALIAS & "', '" & Replace$(sReportID, "'", "''") & "', '" & Profile & "', '" & Replace$(TargetFile.FilePath, "\", "\\") & "', " & _
                         "        '" & Replace$(TargetFile.FileName, "'", "''") & "', '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "', 0, '" & Replace$(Replace$(Comments, "'", "''"), "\", "\\") & "');"
144             oCnDtsDAQ.Execute strSQL, lRet
146             If lRet <> 1 Then
148                 bRet = False
150                 sRet = IIf(sRet <> "", sRet & vbCrLf, "") & "Failed to execute SQL statement: " & strSQL
152                 oCnDtsDAQ.Errors.Refresh
154                 If oCnDtsDAQ.Errors.Count > 0 Then
156                     For lCtr = 0 To (oCnDtsDAQ.Errors.Count - 1)
158                         sRet = sRet & vbCrLf & "    Error " & oCnDtsDAQ.Errors(lCtr).Number & ": " & oCnDtsDAQ.Errors(lCtr).Description
                        Next
160                     oCnDtsDAQ.Errors.Clear
                    End If
                
                End If

            End If
        
162         .Close

        End With

FileNotFound_End:

        On Error Resume Next
    
164     If Not oRS Is Nothing Then
166         If oRS.State <> adStateClosed Then oRS.Close
168         Set oRS = Nothing
        End If

170     Results = sRet
    
172     FileNotFound = bRet
    
        Exit Function
    

FileNotFound_EH:

174     sErrMsg = "Unexpected error occurred in FileNotFound()" & IIf(Erl <> 0, " at line " & Erl, "") & vbCrLf & _
                  "    Error [" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & _
                  IIf(strSQL <> "", "    SQL = " & strSQL, "")

        On Error Resume Next

176     sRet = IIf(sRet <> "", sRet & vbCrLf, "") & sErrMsg

178     bRet = False
    
180     Resume FileNotFound_End

End Function

Private Function OpenRS(DatabaseCN As ADODB.Connection, QueryString As String, Optional CursorType As ADODB.CursorTypeEnum = adOpenStatic, Optional LockType As ADODB.LockTypeEnum = adLockReadOnly, Optional CommandOptions As ADODB.CommandTypeEnum = adCmdText, Optional Results As String = "") As ADODB.Recordset

    On Error GoTo OpenRS_EH

    Dim oRS As ADODB.Recordset
    Dim sRet As String, lCtr As Long

100     Set oRS = New ADODB.Recordset

102     oRS.Open QueryString, DatabaseCN, CursorType, LockType, CommandOptions
    
OpenRS_End:

        On Error Resume Next
    
104     Results = sRet

106     Set OpenRS = oRS
    
        Exit Function
    

OpenRS_EH:

108     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in OpenRS(" & QueryString & ")" & IIf(Erl <> 0, " at line " & Erl, "")

        On Error Resume Next

110     If DatabaseCN.Errors.Count > 0 Then
112         For lCtr = 0 To DatabaseCN.Errors.Count - 1
114             Debug.Print "DbConnection error [" & DatabaseCN.Errors(lCtr).Number & "] " & DatabaseCN.Errors(lCtr).Description
            Next
        End If
        
116     Resume OpenRS_End
    
End Function

Private Sub EnableUI(Enable As Boolean)

On Error Resume Next

    mcmdRunNow.Enabled = Enable
    mcmdResend.Enabled = Enable
    moptAutoRun.Enabled = Enable
    mcmdSettings.Enabled = Enable
    'moptDebugMode.Enabled = Enable

End Sub

Public Sub UpdateStatus(Optional ByVal Message As String = "", Optional ByVal Permenant As Boolean = True, Optional ByRef Cancel As Boolean = False)

On Error Resume Next

Dim sMsg As String

    StatusBar1.Panels(1).Text = Message

    If Permenant Or moptDebugMode.Checked Then
        APP_LOG = AddSlash(App.Path) & "Logs\" & App.EXEName & " - " & Format$(Now(), "yyyy-mm-dd") & ".log"
        AppendString APP_LOG, Message, True
        sMsg = IIf(txtOutput.Text <> "", txtOutput.Text & vbCrLf, "") & Message
        If Len(sMsg) > 65535 Then sMsg = Right$(sMsg, 65535)
        txtOutput.Text = sMsg
        txtOutput.SelStart = Len(sMsg)
    End If

    SleepEx
    
    Cancel = bCloseApp

End Sub

Private Sub LoadLog(LogFile As String)

    On Error GoTo LoadLog_EH

    Dim oFR As FileReader, sMsg As String
    Dim bRet As Boolean, lCtr As Long
    Dim sLine As String

100     UpdateStatus "Loading log file ...", False
    
102     Set oFR = New FileReader
104     bRet = oFR.Load(LogFile)
    
106     If bRet Then
108         For lCtr = 1 To oFR.Lines
110             sLine = oFR.Line(lCtr)
112             If InStr(1, sLine, vbTab) > 0 Then
114                 sLine = Mid$(sLine, InStr(1, sLine, vbTab) + 1)
                End If
116             sMsg = IIf(txtOutput.Text <> "", txtOutput.Text & vbCrLf, "") & sLine
118             If Len(sMsg) > 65535 Then sMsg = Right$(sMsg, 65535)
120             txtOutput.Text = sMsg
        
            Next
122         txtOutput.SelStart = Len(sMsg)

        Else
124         UpdateStatus "Error [" & oFR.ErrorNum & "] " & oFR.ErrorMsg & " occurred in FileReader.Load(" & LogFile & ") while loading log file", True

        End If

LoadLog_End:

        On Error Resume Next
    
126     Set oFR = Nothing
    
        'UpdateStatus "Finished loading log file", False
128     UpdateStatus , False
    
        Exit Sub


LoadLog_EH:

130     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in LoadLog(" & LogFile & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
132     Resume LoadLog_End
    
End Sub

