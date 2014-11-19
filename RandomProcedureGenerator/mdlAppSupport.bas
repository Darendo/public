Attribute VB_Name = "AppSupport"
Option Explicit

Public Const FORMAT_UNITS_ENGLISH = 0
Public Const FORMAT_UNITS_METRIC = 1

Public Enum eFormatUnits
    huEnglish = FORMAT_UNITS_ENGLISH
    huMetric = FORMAT_UNITS_METRIC
End Enum

Public aryLocations() As String

Public Type AdtInfo
    
    'ID                             ' ID  <-- auto-number
    DateCreated As String           ' AddDate
    DateModified As String          ' UpdateDate
    
    GroupNbr As String              ' GroupNbr
    PhysAttending As Long           ' AttPhy
    PhysAdmitting As Long           ' AdmitPhy
    PhysReferring As Long           ' RefPhy
    
    AdmitDate As String             ' AdmitDate
    Location As String              ' Location
    Room As String                  ' Room
    DischargeDate As String         ' DiscDate
    
    MRN As String                   ' MRN
    AccountNbr As String            ' AccountNbr
    BillingNbr As String            ' BillingNbr
    VisitNbr As String
    NameLast As String              ' PatLast
    NameFirst As String             ' PatFirst
    NameMiddle As String
    DOB As String                   ' DOBDate
    Gender As String                ' Sex
    Age As String                   ' Age
    Height As String
    Weight As String

End Type


Public INI_FILE_PATH As String

Public sGroupNbr As String, lngCreateProcDelay As Long
Public cnNames As ADODB.Connection, NamesProvider As String, NamesDataSource As String, NamesInitialCatalog As String, NamesUsername As String, NamesPassword As String
Public cnData As ADODB.Connection, DataProvider As String, DataDataSource As String, DataInitialCatalog As String, DataUsername As String, DataPassword As String
Public StartTop As Long, StartLeft As Long ', StartHeight As Long, StartWidth As Long


Sub Main()

    On Error GoTo Main_EH

        ' Initialize variables
100     Init
102     ConnectDb
    
        ' Setup the main form ...
104     With frmMain
106         .InitDisplay
108         .Show
        End With

        Exit Sub
    

Main_EH:

110     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in RPG.AppSupport.Main()" & IIf(Erl <> 0, " at line " & Erl, "")
    
End Sub


Sub Init()

    On Error GoTo Init_EH

        ' Set the local application's INI
100     INI_FILE_PATH = AddSlash(App.Path) & App.Title & ".ini"
    
        ' General
102     sGroupNbr = CLng(GetINISetting("General", "GroupNumber", "", INI_FILE_PATH))
104     lngCreateProcDelay = CLng(GetINISetting("General", "CreateProcDelay", "1", INI_FILE_PATH))

        ' Databases
106     NamesProvider = GetINISetting("NamesDatabase", "Provider", "", INI_FILE_PATH)
108     NamesDataSource = GetINISetting("NamesDatabase", "DataSource", "", INI_FILE_PATH)
110     NamesInitialCatalog = GetINISetting("NamesDatabase", "InitialCatalog", "", INI_FILE_PATH)
112     NamesUsername = GetINISetting("NamesDatabase", "UID", "", INI_FILE_PATH)
114     NamesPassword = GetINISetting("NamesDatabase", "PWD", "", INI_FILE_PATH)

116     DataProvider = GetINISetting("DataDatabase", "Provider", "", INI_FILE_PATH)
118     DataDataSource = GetINISetting("DataDatabase", "DataSource", "", INI_FILE_PATH)
120     DataInitialCatalog = GetINISetting("DataDatabase", "InitialCatalog", "", INI_FILE_PATH)
122     DataUsername = GetINISetting("DataDatabase", "UID", "", INI_FILE_PATH)
124     DataPassword = GetINISetting("DataDatabase", "PWD", "", INI_FILE_PATH)

        ReDim aryLocations(9)
        aryLocations(0) = "ER"
        aryLocations(1) = "Main"
        aryLocations(2) = "CCU"
        aryLocations(3) = "Pain"
        aryLocations(4) = "RadOnc"
        aryLocations(5) = "Surgery"
        aryLocations(6) = "Admit"
        aryLocations(7) = "Main"
        aryLocations(8) = "Prep"
        aryLocations(9) = "Rehab"
        
        Exit Sub
    

Init_EH:

130     Debug.Print "Error [" & Err.Number & "] " & Err.Description & " occurred in RPG.AppSupport.Init()" & IIf(Erl <> 0, " at line " & Erl, "")

End Sub


Public Sub ConnectDb()

    On Error GoTo ConnectDb_EH
    
        ' Setup the Patient database connection
100     If cnNames Is Nothing Then Set cnNames = New ADODB.Connection
102     With cnNames
104         .ConnectionString = BuildDbConnectionString(NamesProvider, NamesDataSource, NamesInitialCatalog, , NamesUsername, NamesPassword)
106         If .State <> adStateClosed Then .Close
108         .Open
        End With
    
        ' Setup the Patient database connection
110     If cnData Is Nothing Then Set cnData = New ADODB.Connection
112     With cnData
114         .ConnectionString = BuildDbConnectionString(DataProvider, DataDataSource, DataInitialCatalog, , DataUsername, DataPassword)
116         If .State <> adStateClosed Then .Close
118         .Open
        End With

        Exit Sub
    

ConnectDb_EH:

120     MsgBox "In RPG.AppSupport.ConnectDb()" & IIf(Erl <> 0, " at line " & Erl, "") & vbCrLf & vbCrLf & _
               vbTab & "Error [" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & _
               "Please forward to technical support.", vbExclamation Or vbOKOnly, "An unexpected error has occurred ..."
           
End Sub


Sub CloseDbResources()

On Error Resume Next

    If Not cnNames Is Nothing Then
        If cnNames.State <> adStateClosed Then cnNames.Close
        Set cnNames = Nothing
    End If
    
    If Not cnData Is Nothing Then
        If cnData.State <> adStateClosed Then cnData.Close
        Set cnData = Nothing
    End If

End Sub

