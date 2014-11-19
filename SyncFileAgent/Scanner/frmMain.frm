VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "SyncFile Scanner"
   ClientHeight    =   4590
   ClientLeft      =   165
   ClientTop       =   855
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
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   4275
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1140
      Top             =   3180
   End
   Begin VB.TextBox txtOutput 
      Height          =   3555
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   8730
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
      Begin VB.Menu mnuOptionsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mcmdSettings 
         Caption         =   "&Settings ..."
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
102     sRet = GetINISetting(APP_ALIAS, "Frequency", CStr(m_lFrequency), INI_FILE)
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
        SaveINISetting APP_ALIAS, "Frequency", CStr(m_lFrequency), INI_FILE
        
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
118                 m_lTimerCtr = 0
                
                Else
120                 UpdateStatus "Idle - Run in " & lRet & " second" & IIf(lRet = 1, "", "s") & " ...", False

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

    Dim lFilesQueued As Long, lFilesFailed As Long
    Dim sRet As String, bRet As Boolean

100     m_bBusy = True
102     EnableUI False
    
    
104     If FileExists(AddSlash(App.Path) & "Resends.txt") Then
        
106         UpdateStatus "Queuing resends ...", True
        
108         bRet = ScanFiles(True, lFilesQueued, lFilesFailed, sRet)
        
110         If bRet Then
112             Name AddSlash(App.Path) & "Resends.txt" As AddSlash(App.Path) & "Resends - " & Format$(Now, "yyyymmddhhnnss") & ".txt"
114             UpdateStatus "Completed queuing " & Format$(lFilesQueued, "#,##0") & " resend" & IIf(lFilesQueued = 1, "", "s") & " " & IIf(sRet <> "", "with comments: " & sRet, "OK"), (sRet <> ""), True
        
            Else
116             UpdateStatus "Failed to queue resends" & IIf(sRet <> "", ": " & sRet, ""), True
    
            End If
    
        End If
    
    
117     UpdateStatus "Scanning files ...", True
    
118     lFilesQueued = 0
119     lFilesFailed = 0

120     bRet = ScanFiles(False, lFilesQueued, lFilesFailed, sRet)

122     UpdateStatus "  Queued " & Format$(lFilesQueued, "#,##0") & " file" & IIf(lFilesQueued = 1, "", "s") & " OK", True

124     If bRet Then
126         UpdateStatus "Completed scanning files " & IIf(sRet <> "", "with comments: " & sRet, "OK"), True
    
        Else
128         UpdateStatus "Failed to scan files" & IIf(sRet <> "", ": " & sRet, ""), True

        End If

130     UpdateStatus "", True


SyncFiles_End:

        On Error Resume Next

132     UpdateStatus "Idle", False
    
134     EnableUI True
136     m_bBusy = False

        Exit Sub
    

SyncFiles_EH:

138     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in frmMain.SyncFiles()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next

140     UpdateStatus sRet, True

142     Resume SyncFiles_End

End Sub

Private Function ScanFiles(Optional GetResends As Boolean = False, Optional FilesQueued As Long, Optional FilesFailed As Long, Optional Results As String) As Boolean

On Error GoTo ScanFiles_EH

Dim aryProfiles() As String, lNumProfiles As Long, pCtr As Long
Dim sProfileINI As String, sName As String, dtmLastRun As Date, dtmEndRun As Date, dtmNextRun As Date
Dim sSource As String, aryInclude() As String, sExclude As String, bUseArchiveBit As Boolean, bRecursive As Boolean, lDateModifiedOffset As Long
Dim aryFiles() As FILE_INFO
Dim lFilesFound As Long
Dim sglStart As Single, sglEnd As Single, bCancel As Boolean
Dim lRet As Long, sRet As String, bRet As Boolean

    'lNumProfiles = GetINISection("Profiles", aryProfiles(), INI_FILE)
    sRet = GetINISetting("General", "SourceProfiles", "", INI_FILE)
    If sRet <> "" Then
        If InStr(1, sRet, ";") > 0 Then
            aryProfiles() = Split(sRet, ";")
            lNumProfiles = UBound(aryProfiles()) + 1
        Else
            ReDim aryProfiles(0)
            aryProfiles(0) = sRet
            lNumProfiles = 1
        End If
    Else
        lNumProfiles = 0
    End If
    
    If lNumProfiles > 0 Then
    
        For pCtr = 0 To lNumProfiles - 1
        
            ' Manage settings
            sProfileINI = AddSlash(App.Path) & aryProfiles(pCtr) & ".ini"
            sName = aryProfiles(pCtr)
            sRet = ""
            sSource = ""
            Erase aryInclude()
            sExclude = ""
            bUseArchiveBit = True
            bRecursive = True
            lDateModifiedOffset = 10
            Erase aryFiles()
            FilesQueued = 0

            bRet = LoadProfile(sProfileINI, sName, GetResends, dtmLastRun, dtmEndRun, sSource, aryInclude(), sExclude, bUseArchiveBit, bRecursive, lDateModifiedOffset, sRet)
            If bRet Then
                Caption = DEFAULT_CAPTION & " - " & sName
                SleepEx
            Else
                UpdateStatus "  Failed to load profile " & aryProfiles(pCtr) & IIf(sRet <> "", ": " & sRet, ""), True, bCancel
                If bCancel Then
                    sRet = "Operation cancelled by user"
                    GoTo ScanFiles_End
                Else
                    GoTo Profile_End
                End If
            End If


            ' Scan for new/modified files
            UpdateStatus "  Searching for files " & IIf(GetResends, "to resend ", "") & "in " & sName & " modified on or after " & CStr(dtmLastRun) & " ...", True, bCancel
            If bCancel Then
                sRet = "Operation cancelled by user"
                GoTo ScanFiles_End
            End If
            dtmNextRun = Now()

            sglStart = Timer
            lFilesFound = FindFiles(sSource, aryInclude(), sExclude, bUseArchiveBit, bRecursive, dtmLastRun, dtmEndRun, aryFiles(), sRet)
            sglEnd = Timer


            ' Gracefully exit if necesarry
            If bCloseApp Then
                bCancel = True
                GoTo Profile_End
            End If


            ' Log new/modified files
            If lFilesFound >= 0 Then
                
                '''FilesQueued = lRet
                
                If sglEnd >= sglStart Then
                    sglEnd = sglEnd - sglStart
                Else
                    sglEnd = 0
                End If

                UpdateStatus "  Found " & Format$(lFilesFound, "#,##0") & " file" & IIf(lFilesFound = 1, "", "s") & IIf(GetResends, " to resend", "") & " in " & sName & _
                             " - time elapsed: " & Format$(sglEnd, "#,##0.000") & " second" & IIf(sglEnd = 1, "", "s"), True, bCancel

                If bCancel Then
                    sRet = "Operation cancelled by user"
                    GoTo ScanFiles_End
                End If

                If lFilesFound > 0 Then
                    
                    UpdateStatus "  Queuing " & Format$(lFilesFound, "#,##0") & " file" & IIf(lFilesFound = 1, "", "s") & " ...", True, bCancel
                    
                    sglStart = Timer
                    bRet = SaveListToDb(aryProfiles(pCtr), aryFiles(), lFilesFound, lRet, lDateModifiedOffset, GetResends, sRet)
                    sglEnd = Timer
                    
                    FilesQueued = FilesQueued + lRet
                    
                    If bRet Then

                        If sglEnd >= sglStart Then
                            sglEnd = sglEnd - sglStart
                        Else
                            sglEnd = 0
                        End If
                        
                        UpdateStatus "  Queued " & Format$(lRet, "#,##0") & " file" & IIf(lRet = 1, "", "s") & " OK in " & Format$(sglEnd, "#,##0.000") & " second" & IIf(sglEnd = 1, "", "s"), True, bCancel

                        If bCancel Then
                            sRet = "Operation cancelled by user"
                            GoTo ScanFiles_End
                        End If

                    Else
                        UpdateStatus "  FAILED to commit files to database in " & sName & IIf(sRet <> "", ": " & sRet, ""), True, bCancel
                        GoTo Profile_End
                    
                    End If
                
                End If

            Else
                UpdateStatus "  ERROR finding files in " & sName & IIf(sRet <> "", ":" & sRet, ""), True, bCancel
                GoTo Profile_End

            End If
            

            ' Clean up
            If (Not GetResends) Then SaveINISetting "General", "LastRun", dtmNextRun, sProfileINI
            UpdateStatus "  Completed " & sName & " OK", True, bCancel
            sName = ""

Profile_End:

            If sName <> "" Then UpdateStatus "  Failed to complete " & sName, True, bCancel

            If bCancel Then
                sRet = "Operation cancelled by user"
                GoTo ScanFiles_End
            End If
            
        Next
        
    Else
        UpdateStatus "  No active source profiles found.", True

    End If

    sRet = ""
    bRet = True
    
ScanFiles_End:

    On Error Resume Next

    Caption = DEFAULT_CAPTION
        
    Results = sRet
    
    ScanFiles = bRet
    
    Exit Function


ScanFiles_EH:

    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in ScanFiles(" & CStr(GetResends) & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
    On Error Resume Next
    
    bRet = False
    
    Resume ScanFiles_End

End Function

Private Function LoadProfile(ProfileINI As String, ProfileName As String, GetResends As Boolean, DateLastRun As Date, DateEndRun As Date, SourceFolder As String, _
                             IncludeFileSpec() As String, ExcludeFileSpec As String, UseArchiveBit As Boolean, ScanRecursively As Boolean, DateModifiedOffset As Long, _
                             Optional Results As String = "") As Boolean

    On Error GoTo LoadProfile_EH

    Dim aryDest() As String, dCtr As Long
    Dim sTmp As String
    Dim sRet As String, bRet As Boolean, lRet As Long

100     ProfileName = GetINISetting("General", "Name", ProfileName, ProfileINI)
    
102     Caption = DEFAULT_CAPTION & " - " & ProfileName
        
104     UpdateStatus "  Loading " & ProfileName & " ...", False


106     SourceFolder = GetINISetting("Locations", "Source", "", ProfileINI)
108     If SourceFolder <> "" Then
110         SourceFolder = AddSlash(SourceFolder)
112         If Not FolderExists(SourceFolder) Then
114             sRet = "Source folder for " & ProfileName & " not found: " & SourceFolder
116             GoTo LoadProfile_End
            End If
        Else
118         sRet = "Source folder for " & ProfileName & " is invalid" & IIf(SourceFolder <> "", ": " & SourceFolder, "")
120         GoTo LoadProfile_End
        End If


122     If (Not GetResends) Then
124         sTmp = GetINISetting("Options", "Include", "*", ProfileINI)
126         If sTmp = "" Then sTmp = "*"
128         ReDim IncludeFileSpec(0)
130         IncludeFileSpec(0) = sTmp
        
        Else
132         lRet = LoadResends(AddSlash(App.Path) & "Resends.txt", IncludeFileSpec(), sRet)
134         If lRet <= 0 Then
136             If sRet = "" Then sRet = "Failed to load resends file"
140             GoTo LoadProfile_End

            End If

        End If
    
142     ExcludeFileSpec = GetINISetting("Options", "Exclude", "", ProfileINI)


144     If (Not GetResends) Then
146         sTmp = GetINISetting("Options", "UseArchiveBit", "True", ProfileINI)
148         UseArchiveBit = CBoolEx(sTmp, False)
        Else
150         UseArchiveBit = False
        End If
    
152     sTmp = GetINISetting("Options", "Recursive", "True", ProfileINI)
154     ScanRecursively = CBoolEx(sTmp, True)

    
156     If (Not GetResends) Then
            ' DateLastRun
158         sTmp = GetINISetting("General", "LastRun", "1/28/2009", ProfileINI)
160         If IsDate(sTmp) Then DateLastRun = CDate(sTmp)
            ' DateEndRun  <-- Used when backfilling
162         sTmp = GetINISetting("General", "EndRun", "12:00:00 AM", ProfileINI)
164         If sTmp <> "" And IsDate(sTmp) Then DateEndRun = CDate(sTmp)
        Else
166         DateLastRun = "1/28/2009"
168         DateEndRun = "12:00:00 AM"
        End If
    
170     DateModifiedOffset = 10
172     sTmp = GetINISetting("Options", "DateModifiedOffset", CStr(DateModifiedOffset), ProfileINI)
174     If IsNumeric(sTmp) Then
176         DateModifiedOffset = CLng(sTmp)
        End If

177     sRet = ""
178     bRet = True

LoadProfile_End:

        On Error Resume Next
    
179     Results = sRet
    
180     LoadProfile = bRet
    
        Exit Function
    
    
LoadProfile_EH:
    
182     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in LoadProfile(" & ProfileName & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
184     bRet = False
    
186     Resume LoadProfile_End
    
End Function

Private Function LoadResends(InputFile As String, FilesArray() As String, Optional Results As String = "") As Long

    On Error GoTo LoadResends_EH

    Dim oFR As FileReader, lAdded As Long
    Dim bRet As Boolean, lCtr As Long, lRet As Long, sRet As String

100     Set oFR = New FileReader
102     bRet = oFR.Load(InputFile)
    
104     If bRet Then
106         For lCtr = 1 To oFR.Lines
108             If InStr(1, oFR.Line(lCtr), "|||") > 0 Then
110                 lAdded = lAdded + 1
112                 ReDim Preserve FilesArray(lAdded - 1)
114                 FilesArray(lAdded - 1) = LCase$(Mid$(oFR.Line(lCtr), InStr(1, oFR.Line(lCtr), "|||") + 3))
                End If
            Next
116         lRet = lAdded

        Else
118         sRet = "In LoadResends, FileReader error [" & oFR.ErrorNum & "] " & oFR.ErrorMsg
120         lRet = -1
        
        End If

LoadResends_End:

        On Error Resume Next
    
122     If Not oFR Is Nothing Then Set oFR = Nothing

124     Results = sRet
    
126     LoadResends = lRet
    
        Exit Function
    

LoadResends_EH:

128     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in LoadResends()" & IIf(Erl <> 0, " at line " & Erl, "")
    
130     lRet = -1
    
132     Resume LoadResends_End

End Function

Private Function SaveListToDb(Profile As String, Files() As FILE_INFO, NumFiles As Long, FilesQueued As Long, DateModifiedOffset As Long, Optional ForceResend As Boolean = False, Optional Results As String = "") As Boolean

    On Error GoTo SaveListToDb_EH

    Dim strSQL As String, oRS As ADODB.Recordset
    Dim fCtr As Long, sReportID As String
    Dim sRet As String, bRet As Boolean, lRet As Long, lCtr As Long
    Dim sErrMsg As String

        ' Assume success
100     bRet = True
101     FilesQueued = 0

102     For fCtr = 0 To (NumFiles - 1)
    
104         strSQL = "SELECT * FROM " & COPY_LOG_TABLE & " " & _
                     "WHERE Profile = '" & Profile & "' AND SourceFolder = '" & Replace$(Files(fCtr).FilePath, "\", "\\") & "' AND " & _
                     "      FileName = '" & Replace$(Files(fCtr).FileName, "'", "''") & "';"

106         If oRS Is Nothing Then
108             Set oRS = New ADODB.Recordset
        
            Else
110             If oRS.State <> adStateClosed Then
112                 oRS.Close
                End If
            
            End If
        
114         With oRS
        
116             .Open strSQL, oCnDtsDAQ, adOpenStatic, adLockOptimistic, adCmdText
            
118             If Not .EOF Then
                    'strSQL = "UPDATE " & COPY_LOG_TABLE & " SET DateFileModified = '" & Files(fCtr).FileTimeLastModified & "' WHERE Profile = '" & Profile & "' AND SourceFolder = '" & Files(fCtr).FilePath & "' AND Filename = '" & Files(fCtr).FileName & "';"
120                 If (!DateFileModified <> Files(fCtr).FileTimeLastModified) Or _
                       (!FileSize <> Files(fCtr).FileSize) Or _
                       (!CopiedOK = 0) Or _
                       ForceResend Then
121                     !Source = APP_ALIAS & " - " & Profile
122                     !DateFileModified = Format$(Files(fCtr).FileTimeLastModified, "yyyy-mm-dd hh:nn:ss")
123                     !FileSize = Files(fCtr).FileSize
124                     !DateFileFound = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
126                     !DateFileToCopy = Format$(DateAdd("n", DateModifiedOffset, Now()), "yyyy-mm-dd hh:nn:ss")
128                     !DateFileCopied = Null
130                     !CopiedOK = Null
132                     .Update
133                     FilesQueued = FilesQueued + 1
                    End If

                Else
134                 sReportID = Trim$(GetFileName(Files(fCtr).FileName, False))
136                 If LCase$(Right$(Files(fCtr).FileName, 3)) = "wav" And InStr(1, sReportID, "(0)") > 0 Then
138                     sReportID = Replace$(sReportID, "(0)", "")
                    End If
140                 strSQL = "INSERT INTO " & COPY_LOG_TABLE & " (Source, ReportID, Profile, SourceFolder, Filename, DateFileModified, FileSize, DateFileFound, DateFileToCopy) " & _
                             "VALUES ('" & APP_ALIAS & " - " & Profile & "', '" & Replace$(sReportID, "'", "''") & "', '" & Profile & "', '" & Replace$(Files(fCtr).FilePath, "\", "\\") & "', " & _
                             "        '" & Replace$(Files(fCtr).FileName, "'", "''") & "', '" & Format$(Files(fCtr).FileTimeLastModified, "yyyy-mm-dd hh:nn:ss") & "', " & Files(fCtr).FileSize & ", " & _
                             "        '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "', '" & Format$(DateAdd("n", DateModifiedOffset, Now()), "yyyy-mm-dd hh:nn:ss") & "');"
142                 oCnDtsDAQ.Execute strSQL, lRet
144                 If lRet = 1 Then
145                     FilesQueued = FilesQueued + 1

                    Else
146                     bRet = False
148                     sRet = IIf(sRet <> "", sRet & vbCrLf, "") & "Failed to execute SQL statement: " & strSQL
150                     oCnDtsDAQ.Errors.Refresh
152                     If oCnDtsDAQ.Errors.Count > 0 Then
154                         For lCtr = 0 To (oCnDtsDAQ.Errors.Count - 1)
156                             sRet = sRet & vbCrLf & "    Error " & oCnDtsDAQ.Errors(lCtr).Number & ": " & oCnDtsDAQ.Errors(lCtr).Description
                            Next
158                         oCnDtsDAQ.Errors.Clear
                        End If
                    
                    End If

                End If
            
160             .Close

            End With
        
        Next
    
SaveListToDb_End:

        On Error Resume Next
    
162     If Not oRS Is Nothing Then
164         If oRS.State <> adStateClosed Then oRS.Close
166         Set oRS = Nothing
        End If

168     Results = sRet
    
170     SaveListToDb = bRet
    
        Exit Function
    

SaveListToDb_EH:

172     sErrMsg = "Unexpected error occurred in SaveListToDb()" & IIf(Erl <> 0, " at line " & Erl, "") & vbCrLf & _
                  "    Error [" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & _
                  IIf(strSQL <> "", "    SQL = " & strSQL, "")

        On Error Resume Next

174     sRet = IIf(sRet <> "", sRet & vbCrLf, "") & sErrMsg

176     bRet = False
    
178     Resume SaveListToDb_End

End Function

Private Sub EnableUI(Enable As Boolean)

On Error Resume Next

    mcmdResend.Enabled = Enable
    moptAutoRun.Enabled = Enable
    mcmdRunNow.Enabled = Enable
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
128     UpdateStatus
    
        Exit Sub


LoadLog_EH:

130     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in LoadLog(" & LogFile & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
132     Resume LoadLog_End
    
End Sub

