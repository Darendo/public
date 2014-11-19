VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RAG"
   ClientHeight    =   2265
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDelay 
      Caption         =   "Delay:"
      Height          =   1035
      Left            =   150
      TabIndex        =   2
      Top             =   210
      Width           =   4635
      Begin MSComctlLib.Slider sldDelay 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   1000
         SelStart        =   10
         TickStyle       =   1
         Value           =   10
         TextPosition    =   1
      End
      Begin VB.Label lblDelay 
         Alignment       =   2  'Center
         Caption         =   "Max"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblDelay 
         Alignment       =   2  'Center
         Caption         =   "Min"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   435
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Click here or press F5 to begin."
      Top             =   1350
      Width           =   4635
   End
   Begin VB.Timer tmrMTHandler 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4365
      Top             =   45
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   1935
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5821
            Text            =   "Idle"
            TextSave        =   "Idle"
            Key             =   "Status"
            Object.ToolTipText     =   "Current operational status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "0"
            TextSave        =   "0"
            Key             =   "ADT"
            Object.ToolTipText     =   "Total number of ADT records"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "0"
            TextSave        =   "0"
            Key             =   "Problems"
            Object.ToolTipText     =   "Total number of problems"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mcmdGo 
         Caption         =   "&Go"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mcmdGoEx 
         Caption         =   "G&o ..."
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mcmdExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu moptIncludeExistingPeople 
         Caption         =   "Include existing &people"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lWait As Long, bBusy As Boolean
Private WithEvents oRAG As clsRAG
Attribute oRAG.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Set oRAG = New clsRAG

End Sub

Public Sub InitDisplay()

    With Me
        
        .Caption = App.Title & " for Docuworxs v." & Format$(App.Major, "0") & IIf(App.Minor <> 0 And App.Revision <> 0, "." & Format$(App.Minor, "0"), "") & IIf(App.Revision <> 0, "." & Format$(App.Revision, "0"), "")
        
        sldDelay.Value = lngCreateProcDelay
        moptIncludeExistingPeople.Checked = False
        
        Me.UpdateStatus "Idle"
        Me.StatusBar.Panels("ADT").Text = "0"
        Me.StatusBar.Panels("Problems").Text = "0"
        
    End With
    
    DoEvents

End Sub

Private Sub mcmdGoEx_Click()

Dim sRet As String, lRet As Long

    sRet = InputBox("Enter the number of rows to generate:", "Generate ADT rows ...", "10")
    
    If IsNumeric(sRet) Then lRet = CLng(sRet)
    
    If lRet > 0 Then
        GenerateADT lRet

    Else
        MsgBox "Please enter a positive whole number.", vbExclamation Or vbOKOnly, "Unable to continue ..."
        
    End If
    
End Sub

Private Sub moptIncludeExistingPeople_Click()
    moptIncludeExistingPeople.Checked = (Not moptIncludeExistingPeople.Checked)
End Sub

Private Sub sldDelay_Change()
    lngCreateProcDelay = sldDelay.Value '* 10
End Sub

Private Sub mcmdGo_Click()
    GoControl
End Sub

Private Sub cmdControl_Click()
    GoControl
End Sub

Private Sub mcmdExit_Click()
    Unload Me
    CloseDbResources
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Need to allow for running config

    ' Close the recordset(s) and db connection(s)
    CloseDbResources
    
    If Not oRAG Is Nothing Then Set oRAG = Nothing
    
    ' End the application
    End

End Sub



Private Sub GoControl()
    If cmdControl.Caption = "&Go" Then
        cmdControl.Caption = "&Stop"
        mcmdGo.Caption = "&Stop"
        lWait = 0
    Else
        cmdControl.Caption = "&Go"
        mcmdGo.Caption = "&Go"
        lWait = lngCreateProcDelay
    End If
End Sub

Private Sub tmrMTHandler_Timer()

    tmrMTHandler.Enabled = False
    
    If mcmdGo.Caption = "&Stop" Then
    
        If lWait = 0 Then
            GenerateADT
            lWait = lngCreateProcDelay
        
        Else
            lWait = lWait - 1
        
        End If

    End If

    tmrMTHandler.Enabled = True

End Sub


Private Sub GenerateADT(Optional NumRecords As Long = 0)

Dim strSQL As String
Dim udtAdtInfo As AdtInfo, lCtr As Long, bCont As Boolean
Dim bRet As Boolean, sRet As String, lRet As Long

    If Not bBusy Then
        bBusy = True
    
    Else
        Exit Sub

    End If
    
    Do
    
        ' The CustomerID
        oRAG.GroupNbr = sGroupNbr
        
        ' To Include Existing Patients or Not
        oRAG.IncludeExistingPatients = moptIncludeExistingPeople.Checked
        
        ' Generate a procedure
        bRet = oRAG.MakePerson(sRet)
        lCtr = lCtr + 1
        SleepEx

        If bRet Then
        
            With oRAG
        
                'Debug.Print "GroupNbr = " & .GroupNbr               ' GroupNbr
                'Debug.Print "PhysAttending = " & .PhysAttending     ' AttPhy
                'Debug.Print "PhysAdmitting = " & .PhysAdmitting     ' AdmitPhy
                'Debug.Print "PhysReferring = " & .PhysReferring     ' RefPhy
                '
                'Debug.Print "AdmitDate = " & .AdmitDate             ' AdmitDate
                'Debug.Print "DischargeDate = " & .DischargeDate     ' DiscDate
                'Debug.Print "Room = " & .Room                       ' Room
                '
                'Debug.Print "MRN = " & .MRN                         ' MRN
                'Debug.Print "AccountNbr = " & .AccountNbr           ' AccountNbr
                'Debug.Print "BillingNbr = " & .BillingNbr           ' BillingNbr
                'Debug.Print "VisitNbr = " & .VisitNbr
                'Debug.Print "NameLast = " & .NameLast               ' PatLast
                'Debug.Print "NameFirst = " & .NameFirst             ' PatFirst
                'Debug.Print "NameMiddle = " & .NameMiddle
                'Debug.Print "DOB = " & .DOB                         ' DOBDate
                'Debug.Print "Gender = " & .Gender                   ' Sex
                'Debug.Print "Age = " & .Age                         ' Age
                'Debug.Print "Height = " & .Height
                'Debug.Print "Weight = " & .Weight
                '
                'Debug.Print "DateCreated = " & .DateCreated         ' AddDate
                'Debug.Print "DateModified = " & .DateModified       ' UpdateDate
            
                '  ------------------------------------
                ' | Append the new patient information |
                '  ------------------------------------
                strSQL = "INSERT INTO tblADT (GroupNbr, AdmitDate, DiscDate, " & _
                                      IIf(.PhysAttending > 0, "AttPhy, ", "") & IIf(.PhysAdmitting > 0, "AdmitPhy, ", "") & IIf(.Location <> "", "Location, ", "") & IIf(.Room <> "", "Room, ", "") & _
                                      "MRN, AccountNbr, BillingNbr, PatLast, PatFirst, Sex, " & IIf(.DOB <> "", "DOBDate, ", "") & IIf(.Age <> "", "Age, ", "") & _
                                      "AddDate " & _
                         ") VALUES ('" & .GroupNbr & "', '" & .AdmitDate & "', '" & .DischargeDate & "', " & _
                                      IIf(.PhysAttending > 0, "'" & .PhysAttending & "', ", "") & IIf(.PhysAdmitting > 0, "'" & .PhysAdmitting & "', ", "") & IIf(.Location <> "", "'" & .Location & "', ", "") & IIf(.Room <> "", "'" & .Room & "', ", "") & _
                                      "'" & .MRN & "', '" & .AccountNbr & "', '" & .BillingNbr & "', '" & Replace(.NameLast, "'", "''") & "', '" & .NameFirst & "', '" & .Gender & "', " & IIf(.DOB <> "", "'" & Format$(.DOB, "yyyy-mm-dd") & "', ", "") & IIf(.Age <> "", "'" & .Age & "', ", "") & _
                                      "'" & Format$(.DateCreated, "yyyy-mm-dd") & "')"
            
                    Debug.Print "strSQL = " & strSQL
                    
            End With
            
            cnData.Execute strSQL, lRet
        
            If lRet = 1 Then
                Me.IncrementADT
            
            Else
                Me.IncrementProblems
            
            End If
        
        Else
            Me.IncrementProblems
        
        End If
    
        bCont = ((lCtr < NumRecords) Or (NumRecords = 0))

    Loop While bCont
    
    bBusy = False

End Sub

Private Sub oRAG_Status(MsgType As String, MsgString As String)
    
    Select Case MsgType
    Case "INFO"
        UpdateStatus MsgString
    Case "ERROR"
        MsgBox MsgString
        UpdateStatus "Error encountered. RPG Halted!"
    Case Else
        MsgBox "Unknown message type '" & MsgType & "' with value '" & MsgString & "'"
    End Select

End Sub



Public Sub UpdateStatus(strMsg As String)
    Me.StatusBar.Panels("Status").Text = strMsg
    DoEvents
End Sub

Public Sub IncrementADT()
    Me.StatusBar.Panels("ADT").Text = CLng(Me.StatusBar.Panels("ADT").Text) + 1
    DoEvents
End Sub

Public Sub IncrementProblems()
    Me.StatusBar.Panels("Problems").Text = CLng(Me.StatusBar.Panels("Problems").Text) + 1
    DoEvents
End Sub


