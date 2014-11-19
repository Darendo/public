VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "SyncFile Copier"
   ClientHeight    =   4620
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9945
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
   ScaleHeight     =   4620
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   4305
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13414
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Copied OK"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Failed to copy"
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
      Height          =   3645
      Left            =   450
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   135
      Width           =   7515
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

    Dim lFilesCopied As Long, lFilesFailed As Long
    Dim sRet As String, bRet As Boolean

100     m_bBusy = True
102     EnableUI False
    
104     UpdateStatus "Copying files ...", True
    
106     bRet = CopyFiles(lFilesCopied, lFilesFailed, sRet)

108     UpdateStatus "  Copied " & Format$(lFilesCopied, "#,##0") & " file" & IIf(lFilesCopied = 1, "", "s") & " OK", True
110     If lFilesFailed > 0 Then UpdateStatus "  Failed to copy " & Format$(lFilesFailed, "#,##0") & " file" & IIf(lFilesFailed = 1, "", "s"), True
    
112     If bRet Then
114         UpdateStatus "Completed copying files " & IIf(sRet <> "", "with comments: " & sRet, "OK"), True

        Else
116         UpdateStatus "Failed to copy files" & IIf(sRet <> "", ": " & sRet, ""), True

        End If

118     UpdateStatus "", True


SyncFiles_End:

        On Error Resume Next

120     UpdateStatus "Idle", False
    
122     EnableUI True
124     m_bBusy = False

        Exit Sub
    

SyncFiles_EH:

126     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in frmMain.SyncFiles()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next

128     UpdateStatus sRet, True

130     Resume SyncFiles_End

End Sub


Private Function CopyFiles(Optional FilesCopied As Long, Optional FilesFailed As Long, Optional Results As String = "") As Boolean

    On Error GoTo CopyFiles_EH

    Dim sSqlFile As String, strSQL As String, oRS As ADODB.Recordset
    Dim sProfileINI As String, sName As String
    Dim sSource As String, bUseArchiveBit As Boolean
    Dim aryDest() As String, sDest As String, sTmp As String, sComments As String
    Dim fCtr As Long, lFailureCtr As Long
    Dim sglStart As Single, sglEnd As Single, bCancel As Boolean
    Dim sRet As String, bRet As Boolean

        'strSQL = "SELECT * FROM " & COPY_LOG_TABLE & " WHERE " & _
                 "DateFileToCopy <= '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "' AND " & _
                 "CopiedOK IS NULL ORDER BY DateFileToCopy ASC;"
    
100     sSqlFile = GetINISetting(APP_ALIAS, "Query", "", INI_FILE)
102     If sSqlFile <> "" Then
104         If InStr(1, sSqlFile, "\\") = 0 And InStr(1, sSqlFile, ":") = 0 Then
106             sSqlFile = AddSlash(App.Path) & sSqlFile
            End If
108         strSQL = LoadSQLFile(sSqlFile, sRet)
110         If strSQL <> "" Then
112             strSQL = Replace$(strSQL, "#CopyLogTable#", COPY_LOG_TABLE)
114             strSQL = Replace$(strSQL, "#CurrentDateTime#", Format$(Now(), "yyyy-mm-dd hh:nn:ss"))
            Else
116             sRet = "Failed to load SQL query from " & GetFileName(sSqlFile) & IIf(sRet <> "", ": " & sRet, "")
118             GoTo CopyFiles_End
            End If
        Else
120         sRet = "SQL query file not specified"
122         GoTo CopyFiles_End
        End If
    
124     If oRS Is Nothing Then Set oRS = New ADODB.Recordset

126     With oRS
    
128         If .State <> adStateClosed Then .Close

130         .Open strSQL, oCnDtsDAQ, adOpenStatic, adLockOptimistic, adCmdText

132         Do While Not .EOF

134             If !Profile <> sName Then
136                 sName = !Profile
138                 sProfileINI = AddSlash(App.Path) & sName & ".ini"
140                 bRet = LoadProfile(sProfileINI, sName, sSource, aryDest(), bUseArchiveBit, sRet)
142                 If bRet Then
144                     Caption = DEFAULT_CAPTION & " - " & sName
146                     SleepEx
                    Else
148                     UpdateStatus " FAILED to copy queue ID " & !clUID & ": unable to load profile " & sName & IIf(sRet <> "", ": " & sRet, ""), True
150                     GoTo FileCopy_End
                    End If
                End If
            
152             For fCtr = 0 To UBound(aryDest())
                
154                 sDest = aryDest(fCtr)

156                 UpdateStatus "  Copying " & !FileName & " to " & sDest & " ...", False, bCancel
158                 If bCancel Then
160                     sRet = "Operation cancelled by user"
162                     GoTo CopyFiles_End
                    End If
                
164                 sTmp = ""
166                 If Not IsTestServer(sDest) Then
                        ' The file is likely going to a DAQ
168                     If Len(sSource) <> Len(!SourceFolder) Then
170                         sTmp = Mid$(!SourceFolder, Len(sSource) + 1)
172                         If Right$(sTmp, 1) = "\" Then sTmp = Left$(sTmp, Len(sTmp) - 1)
174                         If Left$(sTmp, 1) = "\" Then sTmp = Mid$(sTmp, 2)
176                         If sTmp <> "" Then sTmp = Replace$(sTmp, "\", "~") & "~"
                        End If
178                     sTmp = sDest & NowEx("yyyymmdd", "", "hhnnss", "") & "_" & sTmp & !FileName
                
                    Else
                        ' The files is going to a backup/test server
180                     If Len(sSource) <> Len(!SourceFolder) Then
182                         sTmp = AddSlash(Mid$(!SourceFolder, Len(sSource) + 1))
                        End If
184                     sTmp = sDest & sTmp & !FileName
                
                    End If
                
186                 lFailureCtr = 0

FileCopy_Start:

188                 sRet = ""
190                 sComments = ""
                
                    'Debug.Print !FileName & vbTab & !DateFileModified
                    'GoTo FileCopy_End

                    'On Error Resume Next
                    'FileCopy !SourceFolder & !FileName, sTmp
                    'On Error GoTo CopyFiles_EH

192                 If Not MkDirEx(GetParentDir(sTmp)) Then
194                     UpdateStatus "  ERROR: Failed to create destination folder: " & GetParentDir(sTmp), True, bCancel
196                     If Not bCancel Then
198                         sRet = "Operation cancelled by user"
200                         GoTo FileCopy_End
                        Else
202                         GoTo CopyFiles_End
                        End If
                    End If
                
204                 UpdateStatus "  Copying " & !SourceFolder & !FileName & " to " & sTmp, False, bCancel
206                 If bCancel Then
208                     sRet = "Operation cancelled by user"
210                     GoTo CopyFiles_End
                    End If

212                 bRet = CopyFile(!SourceFolder & !FileName, sTmp, sRet)

214                 If bRet Then
                    
216                     !DateFileCopied = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
218                     !CopiedOK = 1
220                     If sRet <> "" Then
222                         sComments = sRet
224                         If Not IsNull(!Comments) Then
226                             sComments = !Comments & "; " & sComments
228                             If Len(sComments) > 255 Then sComments = Right$(sComments, 255)
                            End If
230                         !Comments = sComments
                        End If
232                     .Update
                    
234                     FilesCopied = FilesCopied + 1
236                     If bInIDE Then AppendString JOB_LOG, NowEx(, , , ",") & "," & sName & "," & !SourceFolder & "," & sDest & "," & !FileName & "," & !DateFileModified & "," & "1" & ",", False
238                     UpdateStatus "  Copied " & !SourceFolder & !FileName & " to " & sDest & " OK", (lFailureCtr > 0)
240                     StatusBar1.Panels(2).Text = Format$(CLng(Val(StatusBar1.Panels(2).Text)) + 1, "#,##0")
242                     If bUseArchiveBit Then
244                         UpdateStatus "  Clearing archive bit for " & !SourceFolder & !FileName, False, bCancel
246                         bRet = SetArchiveBit(!SourceFolder & !FileName, False)
248                         If Not bRet Then
250                             UpdateStatus "  Failed to clear archive bit for " & !SourceFolder & !FileName, True, bCancel
252                             If bCancel Then
254                                 sRet = "Operation cancelled by user"
256                                 GoTo CopyFiles_End
                                End If
                            End If
                        End If
                
                    Else
258                     lFailureCtr = lFailureCtr + 1
260                     If lFailureCtr < 3 Then
262                         UpdateStatus "  Failed to copy " & !SourceFolder & !FileName & IIf(sRet <> "", ": " & sRet, "") & "; waiting " & lFailureCtr & " second" & IIf(lFailureCtr = 1, "", "s") & " ...", True, bCancel
264                         If bCancel Then
266                             sRet = "Operation cancelled by user"
268                             GoTo CopyFiles_End
                            End If
270                         SleepEx (1000 * lFailureCtr)
272                         GoTo FileCopy_Start

                        Else
                        
274                         !DateFileCopied = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
276                         !CopiedOK = 0
278                         If sRet <> "" Then
280                             sComments = sRet
282                             If Not IsNull(!Comments) Then
284                                 sComments = !Comments & "; " & sComments
286                                 If Len(sComments) > 255 Then sComments = Right$(sComments, 255)
                                End If
288                             !Comments = sComments
                            End If
290                         .Update
                        
292                         FilesFailed = FilesFailed + 1
294                         If bInIDE Then AppendString JOB_LOG, NowEx(, , , ",") & "," & sName & "," & !SourceFolder & "," & sDest & "," & !FileName & "," & !DateFileModified & "," & "0" & "," & sRet, False
296                         StatusBar1.Panels(3).Text = Format$(CLng(Val(StatusBar1.Panels(3).Text)) + 1, "#,##0")
298                         UpdateStatus "  Failed to copy " & !SourceFolder & !FileName & " to " & sTmp & IIf(sRet <> "", ": " & sRet, ""), True, bCancel
                            'GoTo Profile_End
300                         If bCancel Then
302                             sRet = "Operation cancelled by user"
304                             GoTo CopyFiles_End
                            End If

                        End If

                    End If

FileCopy_End:
    
306                 If bCancel Then
308                     sRet = "Operation cancelled by user"
310                     GoTo CopyFiles_End
                    End If

312             Next fCtr
    
314             If bCancel Then
316                 sRet = "Operation cancelled by user"
318                 GoTo CopyFiles_End
                End If
            
320             .MoveNext
        
            Loop
    
        End With

322     sRet = ""
324     bRet = True

CopyFiles_End:

        On Error Resume Next

326     If Not oRS Is Nothing Then
328         If oRS.State <> adStateClosed Then oRS.Close
330         Set oRS = Nothing
        End If

332     Caption = DEFAULT_CAPTION
    
334     Results = sRet
    
336     CopyFiles = bRet
    
        Exit Function


CopyFiles_EH:

338     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in frmMain.CopyFiles()" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
340     bRet = False
    
342     Resume CopyFiles_End

End Function


Private Function LoadProfile(ProfileINI As String, ProfileName As String, _
                             SourceFolder As String, DestArray() As String, UseArchiveBit As Boolean, _
                             Optional Results As String = "") As Boolean

    On Error GoTo LoadProfile_EH

    Dim aryDest() As String, sDest As String, dCtr As Long
    Dim sTmp As String
    Dim bCancel As Boolean, sRet As String, bRet As Boolean, lRet As Long

100     ProfileName = GetINISetting("General", "Name", ProfileName, ProfileINI)
    
104     UpdateStatus "  Loading " & ProfileName & " ...", False, bCancel


106     SourceFolder = GetINISetting("Locations", "Source", "", ProfileINI)
108     If SourceFolder <> "" Then
110             SourceFolder = AddSlash(SourceFolder)
112             If Not FolderExists(SourceFolder) Then
114             sRet = "Source folder for " & ProfileName & " not found: " & SourceFolder
116             GoTo LoadProfile_End
            End If
        Else
118         sRet = "Source folder for " & ProfileName & " is invalid" & IIf(SourceFolder <> "", ": " & SourceFolder, "")
120         GoTo LoadProfile_End
        End If
    
    
122     sTmp = GetINISetting("Locations", "Dest", "", ProfileINI)
124     If sTmp <> "" Then
126         If InStr(1, sTmp, ";") > 0 Then
128             DestArray() = Split(sTmp, ";")
            Else
130             ReDim DestArray(0)
132             DestArray(0) = sTmp
            End If
134         lRet = 0
136         For dCtr = 0 To UBound(DestArray())
138             DestArray(dCtr) = AddSlash(DestArray(dCtr))
140             sDest = DestArray(dCtr)
142             If FolderExists(sDest) Then
144                 lRet = lRet + 1
                End If
146         Next dCtr
148         If lRet = 0 Then
150             sRet = "One or more destination folders for " & ProfileName & " not found"
152             GoTo LoadProfile_End
            End If
        Else
154         sRet = "Destination folder for " & ProfileName & " is invalid" & IIf(sDest <> "", ": " & sDest, "")
156         GoTo LoadProfile_End
        End If

    
158     sTmp = GetINISetting("Options", "UseArchiveBit", "True", ProfileINI)
160     UseArchiveBit = CBoolEx(sTmp, False)

    
162     sRet = ""
164     bRet = True

LoadProfile_End:

        On Error Resume Next
    
166     Results = sRet
    
168     LoadProfile = bRet
    
        Exit Function
    
    
LoadProfile_EH:
    
170     sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in LoadProfile(" & ProfileName & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
        On Error Resume Next
    
172     bRet = False
    
174     Resume LoadProfile_End
    
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

