VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRAG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum enumProperCasePhraseTypes
    pcPhraseName = 1
    pcPhraseSentence = 2
End Enum

Private m_IncludeExistingPatients As Boolean
Private udtAdtInfo As AdtInfo

Private oRS As ADODB.Recordset, strSQL As String


Public Event Status(MsgType As String, MsgString As String)


Private Sub Class_Initialize()

On Error GoTo Class_Initialize_EH

    'InitializeStatusList
    'InitializeSpecialty

    Set oRS = New ADODB.Recordset
    
    Exit Sub
    

Class_Initialize_EH:

    Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in RAG.clsRAG.Class_Initialize()" & IIf(Erl <> 0, " at line " & Erl, "")

End Sub


Private Sub Class_Terminate()

On Error Resume Next

    If Not oRS Is Nothing Then
        If oRS.State <> adStateClosed Then oRS.Close
        Set oRS = Nothing
    End If
    
End Sub


Public Property Get IncludeExistingPatients() As Boolean
    IncludeExistingPatients = m_IncludeExistingPatients
End Property

Public Property Let IncludeExistingPatients(Value As Boolean)
    m_IncludeExistingPatients = Value
End Property


Public Property Get GroupNbr() As String
    GroupNbr = udtAdtInfo.GroupNbr
End Property

Public Property Let GroupNbr(ByVal Value As String)
    If Value <> "" Then
        udtAdtInfo.GroupNbr = Value
    Else
        RaiseEvent Status("ERROR", "Invalid Group Number")
    End If
End Property


Public Property Get PhysAttending() As Long
    PhysAttending = udtAdtInfo.PhysAttending
End Property

Public Property Get PhysAdmitting() As Long
    PhysAdmitting = udtAdtInfo.PhysAdmitting
End Property

Public Property Get PhysReferring() As Long
    PhysReferring = udtAdtInfo.PhysReferring
End Property
    
Public Property Get AdmitDate() As String
    AdmitDate = udtAdtInfo.AdmitDate
End Property

Public Property Get DischargeDate() As String
    DischargeDate = udtAdtInfo.DischargeDate
End Property

Public Property Get Location() As String
    Location = udtAdtInfo.Location
End Property
    
Public Property Get Room() As String
    Room = udtAdtInfo.Room
End Property
    
Public Property Get MRN() As String
    MRN = udtAdtInfo.MRN
End Property

Public Property Get AccountNbr() As String
    AccountNbr = udtAdtInfo.AccountNbr
End Property
    
Public Property Get BillingNbr() As String
    BillingNbr = udtAdtInfo.BillingNbr
End Property
    
Public Property Get VisitNbr() As String
    VisitNbr = udtAdtInfo.VisitNbr
End Property

Public Property Get NameLast() As String
    NameLast = udtAdtInfo.NameLast
End Property

Public Property Get NameFirst() As String
    NameFirst = udtAdtInfo.NameFirst
End Property
    
Public Property Get NameMiddle() As String
    NameMiddle = udtAdtInfo.NameMiddle
End Property

Public Property Get DOB() As String
    DOB = udtAdtInfo.DOB
End Property

Public Property Get Gender() As String
    Gender = udtAdtInfo.Gender
End Property

Public Property Get Age() As String
    Age = udtAdtInfo.Age
End Property

Public Property Get Height() As String
    Height = udtAdtInfo.Height
End Property

Public Property Get Weight() As String
    Weight = udtAdtInfo.Weight
End Property

Public Property Get DateCreated() As String
    DateCreated = udtAdtInfo.DateCreated
End Property

Public Property Get DateModified() As String
    DateModified = udtAdtInfo.DateModified
End Property
    


Public Function MakePerson(Optional Results As String = "") As Boolean

On Error GoTo MakePerson_EH

Dim bRet As Boolean, sRet As String, lRet As Long

    RaiseEvent Status("INFO", "Generating a new person ...")
    

    With udtAdtInfo
        
        '.GroupNbr = ""
        .PhysAttending = 0
        .PhysAdmitting = 0
        .PhysReferring = 0
        
        .AdmitDate = ""
        .DischargeDate = ""
        .Location = ""
        .Room = ""
        
        .MRN = ""
        .AccountNbr = ""
        .BillingNbr = ""
        .VisitNbr = ""
        .NameLast = ""
        .NameFirst = ""
        .NameMiddle = ""
        .Gender = ""
        .DOB = ""
        .Age = ""
        .Height = ""
        .Weight = ""
        
        .DateCreated = Format$(Date, "yyyy-mm-dd")
        .DateModified = Format$(Date, "yyyy-mm-dd")
    
    End With

    
    ' ------------------------------------------------------------
    '  Gather case data
    ' ------------------------------------------------------------
    
    ' Get the patient ID, first name, last name, age, and gender
    RaiseEvent Status("INFO", "Getting patient information ...")
    If m_IncludeExistingPatients Then
        
        lRet = GetRandomNumber(1, 10)
        
        If lRet <= 7 Then
            ' Debug.Print "Creating a new patient"
            bRet = MakePersonEx(sRet)
        
        Else
            bRet = GetPerson(sRet)
            ' Debug.Print "Used existing patient " & udtAdtInfo.PatientID
        
        End If
    
    Else
        ' Debug.Print "Creating a new patient"
        bRet = MakePersonEx(sRet)
    
    End If
    
    If Not bRet Then
        ' We encountered an error getting the patient information
        RaiseEvent Status("ERROR", "An error was encountered while getting patient information" & IIf(sRet <> "", ": " & sRet, ""))
        GoTo MakePerson_End
    
    End If
    
    ' Update the main UI
    RaiseEvent Status("INFO", "Done")
    

MakePerson_End:
    
    On Error Resume Next

    Results = sRet

    MakePerson = bRet
    
    Exit Function


MakePerson_EH:
    
    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in RAG.clsRAG.MakePerson()" & IIf(Erl <> 0, " at line " & Erl, "")
    
    On Error Resume Next

    bRet = False
    
    RaiseEvent Status("ERROR", sRet)
    
End Function


Private Function MakePersonEx(Results As String) As Boolean
    
On Error GoTo MakePersonEx_EH

Dim yBirth As Long, mBirth As Long, dBirth As Long, sDateBirth As String
Dim lHeight As Long, lWeight As Long
Dim dblMin As Double, dblMax As Double, lRet As Long
Dim bRet As Boolean, sRet As String
Dim lCtr As Long

    '  --------------------------------
    ' | Generate our Patient ID Number |
    '  --------------------------------
    ' lRet = GetRandomNumber(1000, 9999)
    ' udtAdtInfo.PatientID = udtAdtInfo.PatientID & "-" & CStr(lRet)


    '  -----------------------------------------
    ' | Get the Patient's Medical Record Number |
    '  -----------------------------------------
    udtAdtInfo.MRN = Format$(Time, "HHMMSS") & GetMilliseconds()
    udtAdtInfo.AccountNbr = udtAdtInfo.MRN
    udtAdtInfo.BillingNbr = udtAdtInfo.MRN

    lRet = GetRandomNumber(1, 10)
    If lRet <= 4 Then
        udtAdtInfo.PhysAttending = 2222
    Else
        udtAdtInfo.PhysAttending = 1098
    End If
    
    udtAdtInfo.AdmitDate = Format$(Now(), "yyyy-mm-dd")
    udtAdtInfo.DischargeDate = Format$(Now(), "yyyy-mm-dd")
    
    lRet = GetRandomNumber(0, 9)
    udtAdtInfo.Location = aryLocations(lRet)
    
    lRet = GetRandomNumber(1, 24)
    udtAdtInfo.Room = lRet
    lRet = GetRandomNumber(65, 68)
    udtAdtInfo.Room = udtAdtInfo.Room & Chr$(lRet)


    '  ----------------
    ' | Get the GENDER |
    '  ----------------
    dblMin = 1
    dblMax = 20
    lRet = GetRandomNumber(dblMin, dblMax)
    Select Case lRet
    Case 1 To 9
        udtAdtInfo.Gender = "M"
    Case 10 To 20
        udtAdtInfo.Gender = "F"
    End Select
    
    
    '  -------------
    ' | Get the DOB |
    '  -------------
    ' Year
    dblMin = 1931
    dblMax = CDbl(Format$(DateAdd("y", -1, GetGMT()), "YYYY"))
    yBirth = GetRandomNumber(dblMin, dblMax)
    ' Month
    dblMin = 1
    dblMax = 12
    mBirth = GetRandomNumber(dblMin, dblMax)
    ' Day
    dblMin = 1
    dblMax = CDbl(Format$(DateAdd("d", -1, DateAdd("m", 1, CDate(mBirth & "/1/" & yBirth))), "d"))
    dBirth = GetRandomNumber(dblMin, dblMax)
    udtAdtInfo.DOB = CStr(mBirth) & "/" & CStr(dBirth) & "/" & CStr(yBirth)
    ' Debug.Print "DateBirth = " & udtAdtInfo.DOB
    
    udtAdtInfo.Age = FormatAge(CDate(udtAdtInfo.DOB))
    
    lHeight = CalcHeight(udtAdtInfo.Gender, CDate(udtAdtInfo.DOB))
    udtAdtInfo.Height = FormatHeight(lHeight, 0)
    
    lWeight = CalcWeight(udtAdtInfo.Gender, CDate(udtAdtInfo.DOB), lHeight)
    udtAdtInfo.Weight = FormatWeight(lWeight, 0)
    
    
    '  --------------------
    ' | Get the FIRST name |
    '  --------------------
    ' Count how many names we have with the appropriate gender
    strSQL = "SELECT Count(ID) as NumRows FROM FirstNames WHERE Gender = '" & udtAdtInfo.Gender & "'"
    oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
    dblMax = oRS!NumRows
    oRS.Close
    dblMin = 1
    ' Randonly select one
    lRet = GetRandomNumber(dblMin, dblMax)

    strSQL = "SELECT * FROM FirstNames WHERE Gender = '" & udtAdtInfo.Gender & "'"
    oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not oRS.EOF Then
        If lRet > 1 Then oRS.Move (lRet - 1)
        udtAdtInfo.NameFirst = ProperCase(oRS!Name, pcPhraseName)
        udtAdtInfo.Gender = oRS!Gender
    Else
        RaiseEvent Status("ERROR", "Unexpected EOF while getting patient first name.")
        GoTo MakePersonEx_EH
    End If
    oRS.Close
    
    
    '  ------------------------
    ' | Get the middle initial |
    '  ------------------------
    dblMin = 65
    dblMax = 90
    lRet = GetRandomNumber(dblMin, dblMax)
    udtAdtInfo.NameMiddle = Chr$(lRet)

    
    '  -------------------
    ' | Get the LAST name |
    '  -------------------
    strSQL = "SELECT Count(ID) as NumRows FROM LastNames"
    oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
    dblMin = 1
    dblMax = oRS!NumRows
    oRS.Close
    
    lRet = GetRandomNumber(dblMin, dblMax)

    strSQL = "SELECT * FROM LastNames WHERE ID = " & lRet
    oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not oRS.EOF Then
        udtAdtInfo.NameLast = ProperCase(oRS!Name, pcPhraseName)
    Else
        RaiseEvent Status("ERROR", "Unexpected EOF while getting patient last name.")
        GoTo MakePersonEx_EH
    End If
    oRS.Close
    
    bRet = True
    
MakePersonEx_End:

    On Error Resume Next

    If Not oRS Is Nothing Then
        If oRS.State = adStateOpen Then oRS.Close
        'Set oRS = Nothing
    End If

    Results = sRet
    
    MakePersonEx = bRet
    
    Exit Function


MakePersonEx_EH:

    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in RAG.clsRAG.MakePersonEx()" & IIf(Erl <> 0, " at line " & Erl, "")
        
    bRet = False

    Resume MakePersonEx_End

End Function


Private Function GetPerson(Results As String) As Boolean

On Error GoTo GetPerson_EH

Dim sDateBirth As String, lHeight As Long, lWeight As Long
Dim dblMin As Double, dblMax As Double, lRet As Long
Dim sRet As String, bRet As Boolean

    '  ---------------------------
    ' | Randomly choose a Patient |
    '  ---------------------------
    strSQL = "SELECT Count(PatientID) as NumRows FROM Patients"
    oRS.Open strSQL, cnData, adOpenStatic, adLockReadOnly, adCmdText
    dblMin = 1
    dblMax = oRS!NumRows
    oRS.Close
    lRet = GetRandomNumber(dblMin, dblMax)

    
    '  -------------------------
    ' | Get Patient Information |
    '  -------------------------
    strSQL = "SELECT * FROM Patients"
    oRS.Open strSQL, cnData, adOpenStatic, adLockReadOnly, adCmdText
    If Not oRS.EOF Then
        
        If lRet > 1 Then oRS.Move (lRet - 1)
        
        udtAdtInfo.MRN = oRS!MedRecNo
        '''udtAdtInfo.PatientID = oRS!PatientID
        udtAdtInfo.NameFirst = oRS!NameFirst
        If Not IsNull(oRS!NameMiddle) Then udtAdtInfo.NameMiddle = oRS!NameMiddle
        udtAdtInfo.NameLast = oRS!NameLast
        udtAdtInfo.Gender = oRS!Gender
        If Not IsNull(oRS!DateBirth) Then
            sDateBirth = oRS!DateBirth
        End If
        oRS.Close
        
        If IsDate(sDateBirth) Then
            
            udtAdtInfo.DOB = sDateBirth
            udtAdtInfo.Age = FormatAge(CDate(udtAdtInfo.DOB))

            lHeight = CalcHeight(udtAdtInfo.Gender, CDate(udtAdtInfo.DOB))
            udtAdtInfo.Height = FormatHeight(lHeight, 0)
            
            lWeight = CalcWeight(udtAdtInfo.Gender, CDate(udtAdtInfo.DOB), lHeight)
            udtAdtInfo.Weight = FormatWeight(lWeight, 1)
        
        Else
            udtAdtInfo.Age = ""
            udtAdtInfo.Height = ""
            udtAdtInfo.Weight = ""
        
        End If
    
    Else
        oRS.Close
        bRet = MakePersonEx(sRet)
    
    End If

GetPerson_End:

    On Error Resume Next
    
    If Not oRS Is Nothing Then
        If oRS.State = adStateOpen Then oRS.Close
        'Set oRS = Nothing
    End If
    
    Results = sRet
    
    GetPerson = bRet
    
    Exit Function


GetPerson_EH:

    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in RAG.clsRAG.GetPerson()" & IIf(Erl <> 0, " at line " & Erl, "")
    
    On Error Resume Next

    bRet = False

    Resume GetPerson_End
    
End Function


Private Function CalcHeight(Gender As String, DateOfBirth As Date, Optional Results As String = "") As Long

On Error GoTo CalcHeight_EH

Dim lAgeInDays As Long, dHeight As Double, lHeight As Long ' , sHeight As String
    
Dim rAvg As Double, rMin As Long, rMax As Long
Dim rArray() As Double ' , bFirstItem As Boolean
Dim c As Long, r As Long, n As Long
Dim s As Double, sm As Double
Dim lCtr As Long
Dim dblMin As Double, dblMax As Double, lRet As Long
Dim sRet As String

    lAgeInDays = DateDiff("d", DateOfBirth, Now()) + 1
    If lAgeInDays > 50000 Then
        lAgeInDays = 50000
    ElseIf lAgeInDays < 1 Then
        lAgeInDays = 1
    End If
    
    strSQL = "SELECT * FROM Height WHERE Gender = '" & Gender & "' AND maxAgeInDays >= " & lAgeInDays & " AND minAgeInDays <= " & lAgeInDays
    oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not oRS.EOF Then
        
        ' rAvg = Round(oRS!Average, 0)
        rAvg = oRS!Average
        rMin = Round(oRS!RangeMinus, 0)
        rMax = Round(oRS!RangePlus, 0)
        oRS.Close
        
        ' Set the below-average values
        c = 0
        For n = rMin To 0 Step -1
            ' Calculate the number of elements
            r = (rMin + 1) - n
            ' Redimension the Array
            ReDim Preserve rArray((r - 1) + c) As Double
            ' Set the values
            For lCtr = c To (c + (r - 1))
                rArray(lCtr) = (rAvg - n)
                c = c + 1
            Next lCtr
        Next n
    
        ' Set the above-average values
        For n = 0 To rMax
            ' Calculate the number of elements
            r = (rMax + 1) - n
            ' Redimension the Array
            ReDim Preserve rArray((r - 1) + c) As Double
            ' Set the values
            For lCtr = c To (c + (r - 1))
                rArray(lCtr) = (rAvg + n)
                c = c + 1
            Next lCtr
        Next n
    
        ' Randomly choose the height
        dblMin = LBound(rArray())
        dblMax = UBound(rArray())
        lRet = GetRandomNumber(dblMin, dblMax)
        dHeight = rArray(lRet)

        If lAgeInDays >= (60 * 365) Then
            ' A height-reducing algorithm for seniors"
            sm = 1
            If Gender = "M" Then
                ' Male
                s = 0.2125
            Else
                ' Female
                s = 0.1875
            End If
            For lCtr = 1 To (((lAgeInDays \ 365) + IIf((lAgeInDays Mod 365) > 0, 1, 0)) - 60)
                If lCtr Mod 10 = 0 Then sm = sm - 0.1
                dHeight = dHeight - (s * sm)
                ' Debug.Print "i = " & lCtr & ", dHeight = " & dHeight
            Next lCtr
        End If
        
    Else
        sRet = "Unexpected EOF"
        dHeight = 0
    End If
    
    lHeight = Round(dHeight, 0)
    ' Debug.Print "dHeight = " & dHeight & " or " & FormatHeight(lHeight, huEnglish)
    
CalcHeight_End:

    On Error Resume Next

    If Not oRS Is Nothing Then
        If oRS.State = adStateOpen Then oRS.Close
        'Set oRS = Nothing
    End If
    
    Results = sRet
    
    CalcHeight = lHeight

    Exit Function
    
    
CalcHeight_EH:

    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in RAG.clsRAG.CalcHeight(" & Gender & ", " & DateOfBirth & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
    On Error Resume Next
    
    lHeight = -1
    
    Resume CalcHeight_End

End Function


Private Function CalcWeight(ByVal Gender As String, ByVal DateOfBirth As Date, ByVal Height As Long, Optional ByVal UseWeightedAverage As Boolean = False, Optional ByRef Results As String = "") As Long

Dim lWeight As Long
Dim lAgeInDays As Long
Dim lHeightMultiplier As Double, iFrame As Integer
Dim rAvg As Long, rMin As Long, rMax As Long
Dim rArray() As Long
Dim c As Long, r As Long, n As Long, lCtr As Long
Dim dblMin As Double, dblMax As Double, lRet As Long
Dim sRet As String

    lAgeInDays = DateDiff("d", DateOfBirth, Now()) + 1
    
    If lAgeInDays <= 9124 Then
        
        ' *** Adolescent
        
        strSQL = "SELECT * FROM WeightAdolescent WHERE Gender = '" & Gender & "' AND maxAgeInDays >= " & lAgeInDays & " AND minAgeInDays <= " & lAgeInDays
        oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
        
        If Not oRS.EOF Then
            
            rAvg = oRS!Average
            rMin = oRS!RangeMinus
            rMax = oRS!RangePlus
            oRS.Close
        
            ' Set the below-average values
            c = 0
            For n = rMin To 0 Step -1
                ' Calculate the number of elements
                r = (rMin + 1) - n
                ' Redimension the Array
                ReDim Preserve rArray((r - 1) + c) As Long
                ' Set the values
                If UseWeightedAverage Then
                    For lCtr = c To (c + (r - 1))
                        rArray(lCtr) = (rAvg - n)
                        c = c + 1
                    Next lCtr
                Else
                    rArray(c) = (rAvg - n)
                    c = c + 1
                End If
            Next n
        
            ' Set the above-average values
            For n = 0 To rMax
                ' Calculate the number of elements
                r = (rMax + 1) - n
                ' Redimension the Array
                ReDim Preserve rArray((r - 1) + c) As Long
                ' Set the values
                If UseWeightedAverage Then
                    For lCtr = c To (c + (r - 1))
                        rArray(lCtr) = (rAvg + n)
                        c = c + 1
                    Next lCtr
                Else
                    rArray(c) = (rAvg + n)
                    c = c + 1
                End If
            Next n
        
            ' Randomly choose the height
            dblMin = LBound(rArray())
            dblMax = UBound(rArray())
            lRet = GetRandomNumber(dblMin, dblMax)
            lWeight = rArray(lRet)

        Else
            sRet = "Unexpected EOF in adolecents table"
            lWeight = -1
            GoTo CalcWeight_End
        
        End If

        
    Else
        
        ' *** Adult

        ' Determine the frame size
        dblMin = 1
        dblMax = 4
        lRet = GetRandomNumber(dblMin, dblMax)
        Select Case lRet
        Case 1
            ' Small
            iFrame = 0
        Case 2, 3
            ' Medium
            iFrame = 1
        Case 4
            ' Large
            iFrame = 2
        End Select
            
        ' No records exist for heights below the following values _
          and so we must correct for them
        If Gender = "M" And Height < 62 Then
            ' Short Male
            lHeightMultiplier = Height / 62
            Height = 62
        ElseIf Gender = "F" And Height < 58 Then
            ' Short Female
            lHeightMultiplier = Height / 58
            Height = 58
        Else
            lHeightMultiplier = 1
        End If
            
        ' Get the range
        strSQL = "SELECT * FROM WeightAdult WHERE Gender = '" & Gender & "' AND Height = " & Height & " AND Frame = " & iFrame
        oRS.Open strSQL, cnNames, adOpenStatic, adLockReadOnly, adCmdText
        
        If Not oRS.EOF Then
            
            rMin = oRS!RangeMin
            rMax = oRS!RangeMax
            oRS.Close
            
            ' Randomly choose the height
            dblMin = rMin
            dblMax = rMax
            lRet = GetRandomNumber(dblMin, dblMax)
            lWeight = lRet
            
            ' Apply the height multiplier -- this "corrects" the weight for short people
            lWeight = lWeight * lHeightMultiplier

        Else
            sRet = "Unexpected EOF in adults table"
            lWeight = -1
            GoTo CalcWeight_End
        
        End If
    
    End If
    
    ' If the weight is more than 25 lb.s, round the weight
    If lWeight >= (25 * 16) Then
        lWeight = Round(lWeight / 16, 0) * 16
    End If

    ' Debug.Print FormatWeight(lWeight, huEnglish)

CalcWeight_End:
    
    On Error Resume Next
    
    If Not oRS Is Nothing Then
        If oRS.State = adStateOpen Then oRS.Close
        'Set oRS = Nothing
    End If
    
    Results = sRet

    CalcWeight = lWeight

    Exit Function


CalcWeight_EH:

    sRet = "Error [" & Err.Number & "] " & Err.Description & " occurred in RAG.clsRAG.CalcWeight(" & Gender & ", " & DateOfBirth & ", " & Height & ")" & IIf(Erl <> 0, " at line " & Erl, "")
    
    On Error Resume Next
    
    lWeight = -1
    
    Resume CalcWeight_End
    
End Function


Private Function GenPatientID() As String
    'GenPatientID = "PAT" & "-" & udtAdtInfo.GroupNbr & "-" & m_FacilityID & "-" & m_MachineID & "-" & Format(GetGMT, "YYYYMMDDHHNNSS") & CStr(GetMilliseconds())
    GenPatientID = Format(GetGMT, "HHNNSS") & CStr(GetMilliseconds())
End Function


'Private Function RandomizeStatus() As String
'
'Dim lMax As Long, lMin As Long, lCtr As Long
'
'    lMin = LBound(StatusList())
'    lMax = UBound(StatusList())
'
'    Randomize   ' Initialize random-number generator.
'    lCtr = Int((lMax - lMin + 1) * Rnd + lMin)
'
'    RandomizeStatus = StatusList(i)
'
'End Function

'Private Function InitializeStatusList() As Boolean
'
''On Error GoTo InitializeStatusList_EH
'
'Dim bFirstEntry As Boolean
'
'    ' Specialty
'
'    strSQL = "SELECT * FROM StatusList WHERE Weight <> 0 ORDER BY StatusID"
'    oRS.Open strSQL, cnData, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not oRS.EOF Then
'        bFirstEntry = True
'        Do While Not oRS.EOF
'            If bFirstEntry Then
'                bFirstEntry = False
'                ReDim StatusList(oRS!Weight - 1)
'            Else
'                ReDim Preserve StatusList(UBound(StatusList) + oRS!Weight)
'            End If
'            For lCtr = (UBound(StatusList) - (oRS!Weight - 1)) To UBound(StatusList)
'                ' StatusList(i) = oRS!cName
'                StatusList(i) = oRS!StatusID
'            Next lCtr
'            oRS.MoveNext
'        Loop
'    End If
'    oRS.Close
'
'    ' For debugging purposes only ...
'    ' For lCtr = lMin To lMax
'    '     AppendString AddSlash(App.Path) & "StatusList.txt", "i = " & lCtr & "    (" & StatusList(i) & ")"
'    ' Next lCtr
'
'    InitializeStatusList = True
'    Exit Function
'
'
'InitializeStatusList_EH:
'
'    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'
'End Function

'Private Function RandomizeSpecialty() As String
'
'Dim lMax As Long, lMin As Long, lCtr As Long
'
'    lMin = LBound(Specialty())
'    lMax = UBound(Specialty())
'
'    Randomize   ' Initialize random-number generator.
'    lCtr = Int((lMax - lMin + 1) * Rnd + lMin)
'
'    RandomizeSpecialty = Specialty(i)
'
'End Function

'Private Function InitializeSpecialty() As Boolean
'
''On Error GoTo InitializeSpecialty_EH
'
'Dim lMax As Long, lMin As Long, lCtr As Long
'
'    strSQL = "SELECT Count(ID) as CountOfSpecID FROM Specialties"
'    oRS.Open strSQL, cnData, adOpenStatic, adLockReadOnly, adCmdText
'    If Not oRS.EOF Then
'        ' lMax = (oRS!CountOfSpecID - 1)
'        ReDim Specialty(oRS!CountOfSpecID - 1)
'    Else
'        MsgBox "No Specialties found.", vbExclamation, "Message"
'        GoTo InitializeSpecialty_End
'    End If
'    oRS.Close
'
'    strSQL = "SELECT DISTINCT * FROM Specialties ORDER BY ID"
'    oRS.Open strSQL, cnData, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not oRS.EOF Then
'        lCtr = 0
'        Do While Not oRS.EOF
'            Specialty(i) = oRS!Specialty
'            lCtr = lCtr + 1
'            oRS.MoveNext
'        Loop
'    End If
'    oRS.Close
'
'    ' For debugging purposes only ...
'    ' For lCtr = LBound(Specialty()) To UBound(Specialty())
'    '     Debug.Print "i = " & lCtr & "    (" & Specialty(i) & ")"
'    ' Next lCtr
'
'    InitializeSpecialty = True
'
'InitializeSpecialty_End:
'    Exit Function
'
'
'InitializeSpecialty_EH:
'
'    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'
'End Function








Private Function ProperCase(ByVal Phrase As String, ByVal PhraseType As enumProperCasePhraseTypes) As String

' Changes case of strPhrase: capitalizes first letter/character and
' sets the rest to lower case for each word in strPhrase

Dim strProperPhrase As String, iCtr As Integer

    On Error GoTo ProperCase_EH

    strProperPhrase = Phrase
    
    If Len(strProperPhrase) > 1 Then
        
        ' Capitalize the first letter, lower-case the rest
        strProperPhrase = UCase$(Mid$(strProperPhrase, 1, 1)) & LCase$(Mid$(strProperPhrase, 2))
        If PhraseType = pcPhraseSentence Then GoTo ProperCase_End
        
        ' Work through the rest of the phrase
        For iCtr = 2 To Len(strProperPhrase)
            If Mid$(strProperPhrase, iCtr - 1, 1) = " " And PhraseType = pcPhraseName Then
                ' The character is the beginning of a new word
                Mid$(strProperPhrase, iCtr, 1) = UCase$(Mid$(strProperPhrase, iCtr, 1))
            ElseIf Mid$(strProperPhrase, iCtr - 1, 1) <> " " Then
                ' The character is in the middle of a word
                ' If iCtr > 3 Then
                '     If Mid$(strProperPhrase, iCtr - 3, 3) = " Mc" Then
                '         Mid$(strProperPhrase, iCtr, 1) = UCase$(Mid$(strProperPhrase, iCtr, 1))
                '     End If
                ' End If
                If iCtr > 2 And PhraseType = pcPhraseName Then
                    If Mid$(strProperPhrase, iCtr - 2, 2) = "Mc" Then
                        Mid$(strProperPhrase, iCtr, 1) = UCase$(Mid$(strProperPhrase, iCtr, 1))
                    End If
                End If
            End If
        Next
    
    Else
        strProperPhrase = UCase$(strProperPhrase)

    End If
    
ProperCase_End:

    ProperCase = strProperPhrase
    Exit Function


ProperCase_EH:
    Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred during ProperCase(" & Phrase & ", " & PhraseType & ")"
    ProperCase = Phrase

End Function

