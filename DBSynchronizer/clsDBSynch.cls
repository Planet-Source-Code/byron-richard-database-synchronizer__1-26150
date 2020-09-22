VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBSynch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sMServer As String
Public sMUser As String
Public sMPass As String
Public sMDBName As String

Public sSServer As String
Public sSUser As String
Public sSPass As String
Public sSDBName As String

Public ReturnStatus As String
Public sScriptFile As String
Public bScriptFix As Boolean
Public bUseSystemObjects As Boolean
Public bUseIndexesWithUnderscore As Boolean
Public bUseBulkCopyForDataXfer As Boolean

Dim sScript As String
Dim sBCPFile As String
Dim sBCPErrFile As String

Dim lMConnID As Long
Dim lSConnID As Long

Dim sSConnString As String
Dim sMConnString As String
Dim cnS As ADODB.Connection
Dim cnM As ADODB.Connection
Dim FSO As FileSystemObject
Dim TS As TextStream

Public oMServer As SQLDMO.SQLServer
Dim oMDatabase As SQLDMO.Database
Dim oMTable As SQLDMO.Table
Dim oMCol As SQLDMO.Column
Dim oMProc As SQLDMO.StoredProcedure
Dim oMView As SQLDMO.View
Dim oMCheck As SQLDMO.Check
Dim oMIndex As SQLDMO.Index
Dim oMKey As SQLDMO.Key
Dim oMTrigger As SQLDMO.Trigger

Public oSServer As SQLDMO.SQLServer
Dim oSDatabase As SQLDMO.Database
Dim oSTable As SQLDMO.Table
Dim oSCol As SQLDMO.Column
Dim oSProc As SQLDMO.StoredProcedure
Dim oSView As SQLDMO.View
Dim oSCheck As SQLDMO.Check
Dim oSIndex As SQLDMO.Index
Dim oSKey As SQLDMO.Key
Dim oSTrigger As SQLDMO.Trigger

Private Const TABLEMISS As String = "Table - "
Private Const FIELDMISS As String = "Field - "
Private Const TYPEDIFF As String = "Field Type - "
Private Const TABLEADD As String = "Add Table - "
Private Const FIELDADD As String = "Add Field - "
Private Const TYPECHANGE As String = "Change Field Type - "
Private Const VIEWMISS As String = "View Missing - "
Private Const VIEWDIFF As String = "View Diff - "
Private Const VIEWADD As String = "Add View - "
Private Const VIEWCHANGE As String = "Change View - "
Private Const PROCMISS As String = "Proc Miss - "
Private Const PROCDIFF As String = "Proc Diff - "
Private Const PROCADD As String = "Add Proc - "
Private Const PROCCHANGE As String = "Change Proc - "
Private Const CHECKMISS As String = "Constr Miss - "
Private Const CHECKDIFF As String = "Constr Diff - "
Private Const CHECKADD As String = "Add Constr - "
Private Const CHECKCHANGE As String = "Change Constr - "
Private Const INDEXMISS As String = "Index Miss - "
Private Const INDEXDIFF As String = "Index Diff - "
Private Const INDEXADD As String = "Add Index - "
Private Const INDEXCHANGE As String = "Change Index - "
Private Const KEYMISS As String = "Key Miss - "
Private Const KEYDIFF As String = "Key Diff - "
Private Const KEYADD As String = "Add Key - "
Private Const KEYCHANGE As String = "Change Key - "
Private Const TRIGGERMISS As String = "Trigger Miss - "
Private Const TRIGGERDIFF As String = "Trigger Diff - "
Private Const TRIGGERADD As String = "Add Trigger - "
Private Const TRIGGERCHANGE As String = "Change Trigger - "
Private Const CODEFILEMISS As String = "Code File Rec Miss - "
Private Const CODEFILEADD As String = "Add Rec/Data - "
Private Const USERDEFDIFF As String = "UserDef Data Diff - "
Private Const USERDEFCHANGE As String = "Change Data - "

Private Const USERDEFFIELD1 As String = "FieldDesc"
Private Const USERDEFFIELD2 As String = "Row"

Private Const ERR_WHILE_FIXING As String = "Error-"
Private Const FIXED_OK As String = "Fixed"
Private Const SCRIPTED As String = "Scripted"
Private Const UNSCRIPTABLE As String = "Couldn't generate create script from Master"
Private Const BCP_ERROR As String = "Error doing BCP: "
Private Const TOKEN1 As String = ":"
Private Const CODE_FILE_TABLE As String = "tblCodeFile"
Private Const USER_DEF_TABLE As String = "tblUserDefFields"

Private Const TMP_ROOT_DIR As String = "C:\Temp\"

Public vDBObjectsList As Variant
Public Enum DBObjects
  OBJ_TABLE
  OBJ_VIEW
  OBJ_STORED_PROC
  OBJ_CONSTRAINT
  OBJ_INDEX
  OBJ_KEY
  OBJ_TRIGGER
  OBJ_ALL_DB
End Enum
Private Const ALL_OBJ As Integer = 7

Public vDBTablesList As Variant
Public Enum DBTables
  OBJ_CODEFILE = 10
  OBJ_USERDEFTBL = 11
  OBJ_ALL_TABLE = 12
End Enum
Private Const ALL_TABLES As Integer = 12

Public Enum DiffType
  TYPE_DBOBJECTS
  TYPE_DBTABLES
End Enum

Private Sub Class_Initialize()

Set oMServer = New SQLDMO.SQLServer
Set oMDatabase = New SQLDMO.Database
Set oMTable = New SQLDMO.Table
Set oMCol = New SQLDMO.Column
Set oMProc = New SQLDMO.StoredProcedure
Set oMCheck = New SQLDMO.Check
Set oMView = New SQLDMO.View
Set oMIndex = New SQLDMO.Index
Set oMKey = New SQLDMO.Key
Set oMTrigger = New SQLDMO.Trigger

Set oSServer = New SQLDMO.SQLServer
Set oSDatabase = New SQLDMO.Database
Set oSTable = New SQLDMO.Table
Set oSCol = New SQLDMO.Column
Set oSProc = New SQLDMO.StoredProcedure
Set oSView = New SQLDMO.View
Set oSCheck = New SQLDMO.Check
Set oSIndex = New SQLDMO.Index
Set oSKey = New SQLDMO.Key
Set oSTrigger = New SQLDMO.Trigger

Set FSO = New FileSystemObject
If Not FSO.FolderExists(TMP_ROOT_DIR) Then
  FSO.CreateFolder TMP_ROOT_DIR
End If

bUseSystemObjects = False
bUseIndexesWithUnderscore = False
bUseBulkCopyForDataXfer = False

ReDim vDBObjectsList(ALL_OBJ + 1)
vDBObjectsList(0) = "Table/Fields"
vDBObjectsList(1) = "Views"
vDBObjectsList(2) = "Stored Procs"
vDBObjectsList(3) = "Constraints"
vDBObjectsList(4) = "Indexes"
vDBObjectsList(5) = "Keys"
vDBObjectsList(6) = "Triggers"
vDBObjectsList(7) = "All"

ReDim vDBTablesList(ALL_TABLES)
vDBTablesList(10) = "Code File"
vDBTablesList(11) = "User Def Table"
vDBTablesList(12) = "All"

Set cnM = New ADODB.Connection
Set cnS = New ADODB.Connection

End Sub
Private Sub Class_Terminate()

Set oMServer = Nothing
Set oMDatabase = Nothing
Set oMTable = Nothing
Set oMCol = Nothing
Set oMProc = Nothing
Set oMView = Nothing
Set oMCheck = Nothing
Set oMIndex = Nothing
Set oMKey = Nothing
Set oMTrigger = Nothing

Set oSServer = Nothing
Set oMDatabase = Nothing
Set oSTable = Nothing
Set oSCol = Nothing
Set oSProc = Nothing
Set oSView = Nothing
Set oSCheck = Nothing
Set oSIndex = Nothing
Set oSKey = Nothing
Set oSTrigger = Nothing

Set cnS = Nothing
Set cnM = Nothing
Set FSO = Nothing
End Sub


Public Function ConnectMaster() As Boolean

On Error GoTo EH
ConnectMaster = False
ReturnStatus = ""

If oMServer.ConnectionID <> 0 Then
  oMServer.DisConnect
End If
NextLine:
oMServer.Connect sMServer, sMUser, sMPass
lMConnID = oMServer.ConnectionID

Set cnM = New ADODB.Connection
sMConnString = "Provider=SQLOLEDB.1;Initial Catalog=" & sMDBName & ";Data Source=" & sMServer
cnM.Open sMConnString, sMUser, sMPass

ConnectMaster = True
Exit Function
EH:
If Err.Number = -2147201022 Then
  GoTo NextLine
End If
  

ReturnStatus = "Error opening connection: " & Err.Number & " - " & Err.Description
Err.Clear

End Function
Public Function SetMasterDatabase() As Boolean

On Error GoTo EH
SetMasterDatabase = False
ReturnStatus = ""

If sMDBName <> "" Then
  Set oMDatabase = Nothing
  Set oMDatabase = oMServer.Databases(sMDBName)
Else
  ReturnStatus = "Master database name missing."
  Exit Function
End If

SetMasterDatabase = True
Exit Function
EH:
ReturnStatus = "Error setting master database: " & Err.Number & " - " & Err.Description
Err.Clear
End Function
Public Function SetSlaveDatabase() As Boolean

On Error GoTo EH
SetSlaveDatabase = False
ReturnStatus = ""

If sSDBName <> "" Then
  Set oSDatabase = Nothing
  Set oSDatabase = oSServer.Databases(sSDBName)
Else
  ReturnStatus = "Slave database name missing."
  Exit Function
End If

SetSlaveDatabase = True
Exit Function
EH:
ReturnStatus = "Error setting slave database: " & Err.Number & " - " & Err.Description
Err.Cflear
End Function

Public Function ConnectSlave() As Boolean

On Error GoTo EH
ConnectSlave = False
ReturnStatus = ""

If oSServer.ConnectionID <> 0 Then
  oSServer.DisConnect
End If
NextLine:
oSServer.Connect sSServer, sSUser, sSPass
lSConnID = oSServer.ConnectionID

Set cnS = New ADODB.Connection
sSConnString = "Provider=SQLOLEDB.1;Initial Catalog=" & sSDBName & ";Data Source=" & sSServer
cnS.Open sSConnString, sSUser, sSPass

ConnectSlave = True
Exit Function
EH:
If Err.Number = -2147201022 Then
  GoTo NextLine
End If

ReturnStatus = "Error opening connection: " & Err.Number & " - " & Err.Description
Err.Clear

End Function

Public Function GetDatabases( _
  oServer As SQLDMO.SQLServer, _
  sDatabases() As String) As Boolean

Dim iCount As Integer
Dim oDatabase As SQLDMO.Database


On Error GoTo EH
GetDatabases = False

Set oDatabase = New SQLDMO.Database

iCount = 0
For Each oDatabase In oServer.Databases
  ReDim Preserve sDatabases(iCount + 1)
  sDatabases(iCount) = oDatabase.Name
  iCount = iCount + 1
Next

Set oDatabase = Nothing
GetDatabases = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Public Function GetDBObjectDiffInfo( _
  eDBObjectType As DBObjects, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim eMyObjectType As DBObjects

GetDBObjectDiffInfo = False
CreateDiffsRecordset TYPE_DBOBJECTS, rsDiffs
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_TABLE) Then
  eMyObjectType = OBJ_TABLE
  GetDBObjectDiffInfo = GetTableDiffInfo(eMyObjectType, rsDiffs)
End If
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_VIEW) Then
  eMyObjectType = OBJ_VIEW
  GetDBObjectDiffInfo = GetViewDiffInfo(eMyObjectType, rsDiffs)
End If
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_STORED_PROC) Then
  eMyObjectType = OBJ_STORED_PROC
  GetDBObjectDiffInfo = GetSPDiffInfo(eMyObjectType, rsDiffs)
End If
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_CONSTRAINT) Then
  eMyObjectType = OBJ_CONSTRAINT
  GetDBObjectDiffInfo = GetTableObjectDiffInfo(eMyObjectType, rsDiffs)
End If
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_INDEX) Then
  eMyObjectType = OBJ_INDEX
  GetDBObjectDiffInfo = GetTableObjectDiffInfo(eMyObjectType, rsDiffs)
End If
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_KEY) Then
  eMyObjectType = OBJ_KEY
  GetDBObjectDiffInfo = GetTableObjectDiffInfo(eMyObjectType, rsDiffs)
End If
If (eDBObjectType = OBJ_ALL_DB) Or (eDBObjectType = OBJ_TRIGGER) Then
  eMyObjectType = OBJ_TRIGGER
  GetDBObjectDiffInfo = GetTableObjectDiffInfo(eMyObjectType, rsDiffs)
End If

End Function

Public Function GetTableFieldDiffInfo( _
  eTableType As DBTables, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim eMyTableType As DBTables
GetTableFieldDiffInfo = False
CreateDiffsRecordset TYPE_DBTABLES, rsDiffs
If (eTableType = OBJ_ALL_TABLE) Or (eTableType = OBJ_CODEFILE) Then
  eMyTableType = OBJ_CODEFILE
  GetTableFieldDiffInfo = GetCodeFileDiffInfo(eMyTableType, rsDiffs)
End If
If (eTableType = OBJ_ALL_TABLE) Or (eTableType = OBJ_USERDEFTBL) Then
  eMyTableType = OBJ_USERDEFTBL
  GetTableFieldDiffInfo = GetUserDefDiffInfo(eMyTableType, rsDiffs)
End If

End Function

Public Function FixDBObjectDiffInfo( _
  rsDiffs As ADODB.Recordset, _
  bAppendToFile As Boolean) As Boolean

Dim sKey As String
Dim sData() As String

On Error GoTo EH
ReturnStatus = ""
FixDBObjectDiffInfo = False

If bScriptFix = True Then
  If bAppendToFile Then
    Set TS = FSO.OpenTextFile(sScriptFile, ForAppending, False)
  Else
    Set TS = FSO.OpenTextFile(sScriptFile, ForWriting, True)
  End If
End If

rsDiffs.MoveFirst
Do Until rsDiffs.EOF
  
  If CStr(rsDiffs.Fields("Fix")) = "Yes" Then
    sKey = CStr(rsDiffs.Fields("FixKey"))
    sData = Split(sKey, TOKEN1)
    Select Case CInt(sData(0))
    
      Case 0, ALL_OBJ
        If UBound(sData) = 1 Then
          rsDiffs.Fields("Status") = AddTable(sData(1), IIf(CStr(rsDiffs.Fields("XferData")) = "Yes", True, False))
        ElseIf UBound(sData) = 2 Then
          rsDiffs.Fields("Status") = AddFieldToTable(sData(1), sData(2))
        ElseIf UBound(sData) = 4 Then
  ' BJR - Need to install SQL 7.0 SP2 to be able to do next line
'          rsDiffs.Fields("Status") = ChangeFieldDatatype(sData(1), sData(2), sData(3))
        End If
      Case 1, ALL_OBJ
        rsDiffs.Fields("Status") = AddView(sData(1))
      Case 2, ALL_OBJ
        rsDiffs.Fields("Status") = AddSP(sData(1))
      Case 3, 4, 5, 6, ALL_OBJ
        rsDiffs.Fields("Status") = AddTableObject(CInt(sData(0)), sData(1), sData(2))
    
    End Select
    rsDiffs.Update
  End If
  rsDiffs.MoveNext
Loop

If bScriptFix Then
  TS.Close
End If

FixDBObjectDiffInfo = True
Exit Function
EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear


End Function
Public Function FixTableFieldsDiffInfo( _
  rsDiffs As ADODB.Recordset, _
  bAppendToFile As Boolean) As Boolean

Dim sKey As String
Dim sData() As String

On Error GoTo EH
ReturnStatus = ""
FixTableFieldsDiffInfo = False

If bScriptFix = True Then
  If bAppendToFile Then
    Set TS = FSO.OpenTextFile(sScriptFile, ForAppending, False)
  Else
    Set TS = FSO.OpenTextFile(sScriptFile, ForWriting, True)
  End If
End If


rsDiffs.MoveFirst
Do Until rsDiffs.EOF
  
  If CStr(rsDiffs.Fields("Fix")) = "Yes" Then
    sKey = CStr(rsDiffs.Fields("FixKey"))
    sData = Split(sKey, TOKEN1)
    Select Case CInt(sData(0))
    
      Case 10, ALL_TABLES
        rsDiffs.Fields("Status") = AddCodeFileKey(sData(1))
      Case 11, ALL_TABLES
        rsDiffs.Fields("Status") = ChangeUserDefData(sData(1), sData(2), sData(3))
    
    End Select
    rsDiffs.Update
  End If
  rsDiffs.MoveNext
Loop

If bScriptFix Then
  TS.Close
End If

FixTableFieldsDiffInfo = True
Exit Function
EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear


End Function


Private Function GetTableDiffInfo( _
  eDBObjectType As DBObjects, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim iCount As Integer
Dim sMItem As String
Dim sSItem As String
Dim dcSTables As Dictionary
Dim dcSCols As Dictionary
Dim dcSTypes As Dictionary
Dim sKey As String
Dim sKey1 As String
Dim iDum As Integer

On Error GoTo EH
GetTableDiffInfo = False
ReturnStatus = ""

If Not ReadyToProc Then
  Exit Function
End If

Set dcSTables = New Dictionary
Set dcSCols = New Dictionary
Set dcSTypes = New Dictionary

iDum = 0

' Load the Slave info
For Each oSTable In oSDatabase.Tables
  sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oSTable.Name)
  If Not dcSTables.Exists(sKey) Then
    dcSTables.Add sKey, iDum
  End If
  For Each oSCol In oSTable.Columns
    sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oSTable.Name) & TOKEN1 & CStr(oSCol.Name)
    If Not dcSCols.Exists(sKey) Then
      dcSCols.Add sKey, iDum
    End If
    sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oSTable.Name) & TOKEN1 & CStr(oSCol.Name) & TOKEN1 & CStr(oSCol.Datatype) & TOKEN1 & CStr(oSCol.Length)
    If Not dcSTypes.Exists(sKey) Then
      dcSTypes.Add sKey, iDum
    End If
  Next
Next
  

' Loop through master tables, comparing existience, fields, and field types
For Each oMTable In oMDatabase.Tables
  
  If oMTable.TypeOf = SQLDMOObj_UserTable Or bUseSystemObjects Then
    
    ' Does this table exist in slave?
    sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oMTable.Name)
    If Not dcSTables.Exists(sKey) Then
      sMItem = TABLEMISS & oMTable.Name
      sSItem = TABLEADD & oMTable.Name
      AddElementsToLists rsDiffs, sMItem, sSItem, sKey, , "No"
    Else
      
      ' Does the fields in this table exist in slave?
      For Each oMCol In oMTable.Columns
        sKey = CStr(oMTable.Name) & TOKEN1 & CStr(oMCol.Name)
        sKey1 = CStr(eDBObjectType) & TOKEN1 & sKey
        If Not dcSCols.Exists(sKey1) Then
          sMItem = FIELDMISS & sKey
          sSItem = FIELDADD & sKey
          AddElementsToLists rsDiffs, sMItem, sSItem, sKey1
        Else
          ' Are the field types the same for both?
          sKey = CStr(oMTable.Name) & TOKEN1 & CStr(oMCol.Name) & TOKEN1 & CStr(oMCol.Datatype) & TOKEN1 & CStr(oMCol.Length)
          sKey1 = CStr(eDBObjectType) & TOKEN1 & sKey
          If Not dcSTypes.Exists(sKey1) Then
            Set oSCol = oSDatabase(oMTable.Name).Columns(oMCol.Name)
            sMItem = TYPEDIFF & CStr(oMTable.Name) & TOKEN1 & CStr(oMCol.Name) & TOKEN1 & CStr(oMCol.Datatype) & TOKEN1 & CStr(oMCol.Length)
            sSItem = TYPECHANGE & CStr(oMTable.Name) & TOKEN1 & CStr(oSCol.Name) & TOKEN1 & CStr(oSCol.Datatype) & TOKEN1 & CStr(oSCol.Length)
            AddElementsToLists rsDiffs, sMItem, sSItem, sKey1
          End If
        End If
      Next
    End If
  End If ' user defined table
Next


GetTableDiffInfo = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Private Function AddTable(sTableName As String, bXferData As Boolean) As String

Dim iCount As Integer
Dim oBCP As SQLDMO.BulkCopy
Dim oFile As File
Dim sDataXferErr As String

On Error GoTo EH
ReturnStatus = ""
AddTable = ERR_WHILE_FIXING

' Transfer structure
Set oMTable = oMDatabase.Tables(sTableName)
sScript = oMTable.Script(SQLDMOScript_Drops Or SQLDMOScript_Default)
If Trim(sScript) = "" Then
  AddTable = UNSCRIPTABLE
  Exit Function
End If

If bScriptFix Then
  TS.Write sScript
  AddTable = SCRIPTED
Else
  oSDatabase.ExecuteImmediate sScript
  oSDatabase.Tables.Refresh

  If bXferData Then
    On Error GoTo EH2
    If bUseBulkCopyForDataXfer Then
      ' Copy over data
      Set oBCP = New SQLDMO.BulkCopy
      sBCPFile = TMP_ROOT_DIR & FSO.GetTempName
      sBCPErrFile = sBCPFile & ".err"
      oBCP.DataFilePath = sBCPFile
      oBCP.ErrorFilePath = sBCPErrFile
      oBCP.UseBulkCopyOption = True
      oBCP.IncludeIdentityValues = True
      oBCP.DataFileType = SQLDMODataFile_TabDelimitedChar
      oMDatabase.Tables(sTableName).ExportData oBCP
      Set oFile = FSO.GetFile(sBCPErrFile)
      If oFile.Size = 0 Then
        oSDatabase.Tables(sTableName).ImportData oBCP
        FSO.DeleteFile sBCPFile, True
        FSO.DeleteFile sBCPErrFile, True
      End If
    Else
      XferTableDataUsingADO sTableName
    End If
  End If
  AddTable = FIXED_OK
End If ' just scripting

Exit Function
EH:
AddTable = ERR_WHILE_FIXING & Err.Description
Err.Clear
Exit Function

EH2:
  If bUseBulkCopyForDataXfer Then
    AddTable = BCP_ERROR & sBCPErrFile
  Else
    AddTable = "ADO Error: " & Err.Description
  End If
  oSDatabase.Tables.Remove sTableName
  Err.Clear
  Exit Function
End Function

Private Sub XferTableDataUsingADO(sTableName As String)
  
Dim sSQL As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim rsS As ADODB.Recordset
Dim rsM As ADODB.Recordset
Dim iCount As Integer
Dim lRecsAffected As Long
Dim bInTrans As Boolean
Dim iStartNum As Integer

On Error GoTo EH
bInTrans = False

Set rsS = New ADODB.Recordset
Set rsM = New ADODB.Recordset
iStartNum = 0

sSQL = "select * from " & sTableName
rsM.Open sSQL, cnM, adOpenForwardOnly, adLockReadOnly

cnS.BeginTrans
bInTrans = True

TryAgain:
sSQL1 = "Insert into " & sTableName & " ("
For iCount = iStartNum To rsM.Fields.Count - 1
  sSQL1 = sSQL1 & rsM.Fields(iCount).Name & ","
Next iCount
sSQL1 = Mid(sSQL1, 1, Len(sSQL1) - 1) & ") "

Do Until rsM.EOF
  sSQL2 = " Values ("
  For iCount = iStartNum To rsM.Fields.Count - 1
    If Not rsM.Fields(iCount).Attributes = adFldRowID Then
      If IsNull(rsM.Fields(iCount).Value) Then
        sSQL2 = sSQL2 & "null,"
      Else
        If (rsM.Fields(iCount).Type = adChar) Or _
           (rsM.Fields(iCount).Type = adDate) Or _
           (rsM.Fields(iCount).Type = adDBDate) Or _
           (rsM.Fields(iCount).Type = adDBTime) Or _
           (rsM.Fields(iCount).Type = adDBTimeStamp) Or _
           (rsM.Fields(iCount).Type = adLongVarChar) Or _
           (rsM.Fields(iCount).Type = adVarChar) Then
          sSQL2 = sSQL2 & FixString(rsM.Fields(iCount).Value) & ","
        Else
          sSQL2 = sSQL2 & rsM.Fields(iCount).Value & ","
        End If
      End If
    End If
  Next iCount
  sSQL2 = Mid(sSQL2, 1, Len(sSQL2) - 1) & ")"
  sSQL = sSQL1 & sSQL2
  cnS.Execute sSQL, lRecsAffected
  If lRecsAffected <> 1 Then
    Err.Raise vbObjectError + 10, "XferTablesDataUsingADO", _
      "Couldn't transfer data for " & sTableName & " using ADO."
    Exit Sub
  End If
  rsM.MoveNext
Loop

cnS.CommitTrans
bInTrans = False

Exit Sub
EH:
If Err.Number = -2147217900 Then
  rsM.MoveFirst
  iStartNum = 1
  Err.Clear
  On Error GoTo EH
  GoTo TryAgain
End If
If bInTrans Then
  cnS.RollbackTrans
End If
Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function AddFieldToTable( _
  sTableName As String, _
  sFieldName As String) As String

Dim oNewCol As SQLDMO.Column
Dim iCount As Integer

On Error GoTo EH
ReturnStatus = ""
AddFieldToTable = ERR_WHILE_FIXING

Set oMCol = oMDatabase.Tables(sTableName).Columns(sFieldName)
If bScriptFix Then
  If oMCol.Identity Then
    sScript = "ALTER Table " & sTableName & " ADD " & sFieldName & " INT IDENTITY " & _
      "(" & oMCol.IdentitySeed & "," & oMCol.IdentityIncrement & ")"
  Else
    sScript = "ALTER Table " & sTableName & " ADD " & sFieldName
    If (oMCol.Datatype Like "*char") Or (oMCol.Datatype Like "var*") Then
      sScript = sScript & " " & oMCol.Datatype & "(" & oMCol.Length & ") "
    ElseIf (oMCol.Datatype = "decimal") _
        Or (oMCol.Datatype = "numeric") _
        Then
      sScript = sScript & " " & oMCol.Datatype & "(" & oMCol.NumericPrecision & "," & oMCol.NumericScale & ") "
    ElseIf (oMCol.Datatype = "binary") _
        Or (oMCol.Datatype = "bit") _
        Or (oMCol.Datatype Like "*time*") _
        Or (oMCol.Datatype = "float") _
        Or (oMCol.Datatype = "image") _
        Or (oMCol.Datatype Like "*int") _
        Or (oMCol.Datatype Like "*money") _
        Or (oMCol.Datatype Like "*text") _
        Or (oMCol.Datatype = "real") _
        Then
      sScript = sScript & " " & oMCol.Datatype & " "
    End If
    If oMCol.AllowNulls Then
      sScript = sScript & " NULL "
    End If
    If oMCol.Default <> "" Then
      sScript = sScript & " DEFAULT " & oMCol.Default
    End If
  End If
  sScript = sScript & vbCrLf & "GO " & vbCrLf

Else
  oSDatabase.Tables(sTableName).BeginAlter
  Set oNewCol = New SQLDMO.Column
  oNewCol.Name = oMCol.Name
  oNewCol.Datatype = oMCol.Datatype
  oNewCol.Length = oMCol.Length
  oNewCol.AllowNulls = oMCol.AllowNulls
  oNewCol.Default = oMCol.Default
  oNewCol.Identity = oMCol.Identity
  oNewCol.IdentitySeed = oMCol.IdentitySeed
  oNewCol.IdentityIncrement = oMCol.IdentityIncrement
  oNewCol.IsComputed = oMCol.IsComputed
  oNewCol.NumericPrecision = oMCol.NumericPrecision
  oNewCol.NumericScale = oMCol.NumericScale
  oSDatabase.Tables(sTableName).Columns.Add oNewCol
  oSDatabase.Tables(sTableName).DoAlter

End If

If bScriptFix Then
  TS.Write sScript
  AddFieldToTable = SCRIPTED
Else
  AddFieldToTable = FIXED_OK
End If
Exit Function
EH:
If Err.Number = -2147201019 Then
  Resume Next
End If

AddFieldToTable = ERR_WHILE_FIXING & Err.Description
Err.Clear
End Function


Private Function ChangeFieldDatatype( _
  sTableName As String, _
  sFieldName As String, _
  sDatatype As String) As String

On Error GoTo EH
ReturnStatus = ""
ChangeFieldDatatype = ERR_WHILE_FIXING

Set oMCol = oMDatabase.Tables(sTableName).Columns(sFieldName)
Set oSCol = oSDatabase.Tables(sTableName).Columns(sFieldName)
oSDatabase.Tables(sTableName).BeginAlter
oSCol.Datatype = oMCol.Datatype
oSDatabase.Tables(sTableName).DoAlter

ChangeFieldDatatype = FIXED_OK
Exit Function
EH:
ChangeFieldDatatype = ERR_WHILE_FIXING & Err.Description
Err.Clear
End Function

Private Function GetViewDiffInfo( _
  eDBObjectType As DBObjects, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim iCount As Integer
Dim sMItem As String
Dim sSItem As String
Dim sMScript As String
Dim sSScript As String
Dim dcSViews As Dictionary
Dim sKey As String
Dim sKey1 As String
Dim iDum As Integer
Dim bSysObject As Boolean

On Error GoTo EH
GetViewDiffInfo = False
ReturnStatus = ""

If Not ReadyToProc Then
  Exit Function
End If

Set dcSViews = New Dictionary

iDum = 0

' Load the Slave info
For Each oSView In oSDatabase.Views
  sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oSView.Name)
  If Not dcSViews.Exists(sKey) Then
    dcSViews.Add sKey, iDum
  End If
Next
  

' Loop through master view, comparing to slave views
For Each oMView In oMDatabase.Views
  
  bSysObject = Not ((oMView.Owner = "INFORMATION_SCHEMA") Or (oMView.Name Like "sys*"))
  
  If bSysObject Or bUseSystemObjects Then
  
    sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oMView.Name)
    ' Does this view exist in slave?
    If Not dcSViews.Exists(sKey) Then
      sMItem = VIEWMISS & oMView.Name
      sSItem = VIEWADD & oMView.Name
      AddElementsToLists rsDiffs, sMItem, sSItem, sKey
    Else
      Set oSView = oSDatabase.Views(CStr(oMView.Name))
      sMScript = Replace(oMView.Script, vbCrLf, "")
      sSScript = Replace(oSView.Script, vbCrLf, "")
      sMScript = Replace(sMScript, " ", "")
      sSScript = Replace(sSScript, " ", "")
      sMScript = Replace(sMScript, vbTab, "")
      sSScript = Replace(sSScript, vbTab, "")
      If sMScript <> sSScript Then
        sMItem = VIEWDIFF & oMView.Name
        sSItem = VIEWCHANGE & oMView.Name
        AddElementsToLists rsDiffs, sMItem, sSItem, sKey
      End If
    End If
  End If
Next


GetViewDiffInfo = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Private Function AddView(sViewName As String) As String

On Error GoTo EH
ReturnStatus = ""
AddView = ERR_WHILE_FIXING

' Transfer structure
Set oMView = oMDatabase.Views(sViewName)
sScript = oMView.Script(SQLDMOScript_Drops Or SQLDMOScript_Default)
If Trim(sScript) = "" Then
  AddView = UNSCRIPTABLE
  Exit Function
End If
If bScriptFix Then
  TS.Write sScript
  AddView = SCRIPTED
Else
  oSDatabase.ExecuteImmediate sScript
  oSDatabase.Views.Refresh
  AddView = FIXED_OK
End If

Exit Function
EH:
AddView = ERR_WHILE_FIXING & Err.Description
Err.Clear
Exit Function

End Function


Private Function GetSPDiffInfo( _
  eDBObjectType As DBObjects, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim iCount As Integer
Dim sMItem As String
Dim sSItem As String
Dim sMScript As String
Dim sSScript As String
Dim dcStoredProcs As Dictionary
Dim sKey As String
Dim iDum As Integer

On Error GoTo EH
GetSPDiffInfo = False
ReturnStatus = ""

If Not ReadyToProc Then
  Exit Function
End If

Set dcStoredProcs = New Dictionary
iDum = 0

' Load the Slave info
For Each oSProc In oSDatabase.StoredProcedures
  sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oSProc.Name)
  If Not dcStoredProcs.Exists(sKey) Then
    dcStoredProcs.Add sKey, iDum
  End If
Next
  

For Each oMProc In oMDatabase.StoredProcedures
  
  If (Not (oMProc.Name Like "dt_*")) Or bUseSystemObjects Then
  
    sKey = CStr(eDBObjectType) & TOKEN1 & CStr(oMProc.Name)
    ' Does this SP exist in slave?
    If Not dcStoredProcs.Exists(sKey) Then
      sMItem = PROCMISS & oMProc.Name
      sSItem = PROCADD & oMProc.Name
      AddElementsToLists rsDiffs, sMItem, sSItem, sKey
    Else
      ' Is this view different in slave?
      Set oSProc = oSDatabase.StoredProcedures(CStr(oMProc.Name))
      sMScript = Replace(oMProc.Script, vbCrLf, "")
      sSScript = Replace(oSProc.Script, vbCrLf, "")
      sMScript = Replace(sMScript, " ", "")
      sSScript = Replace(sSScript, " ", "")
      sMScript = Replace(sMScript, vbTab, "")
      sSScript = Replace(sSScript, vbTab, "")
      If sMScript <> sSScript Then
        sMItem = PROCDIFF & oMProc.Name
        sSItem = PROCCHANGE & oMProc.Name
        AddElementsToLists rsDiffs, sMItem, sSItem, sKey
      End If
    End If
  End If
Next


GetSPDiffInfo = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Private Function AddSP(sSPName As String) As String

On Error GoTo EH
ReturnStatus = ""
AddSP = ERR_WHILE_FIXING

' Transfer structure
Set oMProc = oMDatabase.StoredProcedures(sSPName)
sScript = oMProc.Script(SQLDMOScript_Drops Or SQLDMOScript_Default)
If Trim(sScript) = "" Then
  AddSP = UNSCRIPTABLE
  Exit Function
End If
If bScriptFix Then
  TS.Write sScript
  AddSP = SCRIPTED
Else
  oSDatabase.ExecuteImmediate sScript
  oSDatabase.StoredProcedures.Refresh
  AddSP = FIXED_OK
End If

Exit Function
EH:
AddSP = ERR_WHILE_FIXING & Err.Description
Err.Clear
Exit Function

End Function

Private Function GetTableObjectDiffInfo( _
  eDBObjectType As DBObjects, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim iCount As Integer
Dim sMItem As String
Dim sSItem As String
Dim sMScript As String
Dim sSScript As String
Dim dcSTableObject As Dictionary
Dim sKey As String
Dim sKey1 As String
Dim iDum As Integer

On Error GoTo EH
GetTableObjectDiffInfo = False
ReturnStatus = ""
iDum = 0

If Not ReadyToProc Then
  Exit Function
End If

Set dcSTableObject = New Dictionary

Select Case eDBObjectType

  Case 3, 7
  
    ' Load the Slave info
    For Each oSTable In oSDatabase.Tables
      For Each oSCheck In oSTable.Checks
        sKey = CStr(eDBObjectType & TOKEN1 & CStr(oSTable.Name) & TOKEN1 & CStr(oSCheck.Name))
        If Not dcSTableObject.Exists(sKey) Then
          dcSTableObject.Add sKey, iDum
        End If
      Next
    Next
    For Each oMTable In oMDatabase.Tables
      For Each oMCheck In oMTable.Checks
      
        ' Does this index exist in slave?
        sKey = CStr(oMTable.Name) & TOKEN1 & CStr(oMCheck.Name)
        sKey1 = CStr(eDBObjectType & TOKEN1 & sKey)
        If Not dcSTableObject.Exists(sKey1) Then
          sMItem = CHECKMISS & sKey
          sSItem = CHECKADD & sKey
          AddElementsToLists rsDiffs, sMItem, sSItem, sKey1
        Else
          Set oSCheck = oSDatabase.Tables(oMTable.Name).Checks(oMCheck.Name)
          sMScript = Replace(Replace(Replace(oMCheck.Script, vbCrLf, ""), vbTab, ""), " ", "")
          sSScript = Replace(Replace(Replace(oSCheck.Script, vbCrLf, ""), vbTab, ""), " ", "")
          If sMScript <> sSScript Then
            sMItem = CHECKDIFF & sKey
            sSItem = CHECKCHANGE & sKey
            AddElementsToLists rsDiffs, sMItem, sSItem, sKey1, oMTable.Name
          End If
        End If
      Next
    Next
    
  Case 4, 7
    ' Load the Slave info
    For Each oSTable In oSDatabase.Tables
      For Each oSIndex In oSTable.Indexes
        If (Not (oSIndex.Name Like "_*")) Or bUseIndexesWithUnderscore Then
          sKey = CStr(eDBObjectType & TOKEN1 & CStr(oSTable.Name) & TOKEN1 & CStr(oSIndex.Name))
          If Not dcSTableObject.Exists(sKey) Then
            dcSTableObject.Add sKey, iDum
          End If
        End If
      Next
    Next
    For Each oMTable In oMDatabase.Tables
      For Each oMIndex In oMTable.Indexes
        If (Not (oMIndex.Name Like "_*")) Or bUseIndexesWithUnderscore Then
          ' Does this index exist in slave?
          sKey = CStr(oMTable.Name) & TOKEN1 & CStr(oMIndex.Name)
          sKey1 = CStr(eDBObjectType & TOKEN1 & sKey)
          If Not dcSTableObject.Exists(sKey1) Then
            sMItem = INDEXMISS & sKey
            sSItem = INDEXADD & sKey
            AddElementsToLists rsDiffs, sMItem, sSItem, sKey1, oMTable.Name
          Else
            ' Is this index different in slave?
            Set oSIndex = oSDatabase.Tables(oMTable.Name).Indexes(oMIndex.Name)
            sMScript = Replace(Replace(Replace(oMIndex.Script, vbCrLf, ""), vbTab, ""), " ", "")
            sSScript = Replace(Replace(Replace(oSIndex.Script, vbCrLf, ""), vbTab, ""), " ", "")
            If sMScript <> sSScript Then
              sMItem = INDEXDIFF & sKey
              sSItem = INDEXCHANGE & sKey
              AddElementsToLists rsDiffs, sMItem, sSItem, CStr(eDBObjectType & TOKEN1 & sKey1)
            End If
          End If
        End If
      Next
    Next
  
  Case 5, 7
  
    ' Load the Slave info
    For Each oSTable In oSDatabase.Tables
      For Each oSKey In oSTable.Keys
        sKey = eDBObjectType & TOKEN1 & CStr(oSTable.Name) & TOKEN1 & CStr(oSKey.Name)
        If Not dcSTableObject.Exists(sKey) Then
          dcSTableObject.Add sKey, iDum
        End If
      Next
    Next
    For Each oMTable In oMDatabase.Tables
      For Each oMKey In oMTable.Keys
      
        ' Does this key exist in slave?
        sKey = CStr(oMTable.Name) & TOKEN1 & CStr(oMKey.Name)
        sKey1 = CStr(eDBObjectType & TOKEN1 & sKey)
        If Not dcSTableObject.Exists(sKey1) Then
          sMItem = KEYMISS & sKey
          sSItem = KEYADD & sKey
          AddElementsToLists rsDiffs, sMItem, sSItem, sKey1, oMTable.Name
        Else
          ' Is this key different in slave?
          Set oSKey = oSDatabase.Tables(oMTable.Name).Keys(oMKey.Name)
          sMScript = Replace(Replace(Replace(oMKey.Script, vbCrLf, ""), vbTab, ""), " ", "")
          sSScript = Replace(Replace(Replace(oSKey.Script, vbCrLf, ""), vbTab, ""), " ", "")
          If sMScript <> sSScript Then
            sMItem = KEYDIFF & sKey
            sSItem = KEYCHANGE & sKey
            AddElementsToLists rsDiffs, sMItem, sSItem, sKey1
          End If
        End If
      Next
    Next
  
  Case 6, 7

    ' Load the Slave info
    For Each oSTable In oSDatabase.Tables
      For Each oSTrigger In oSTable.Triggers
        sKey = eDBObjectType & TOKEN1 & CStr(oSTable.Name) & TOKEN1 & CStr(oSTrigger.Name)
        If Not dcSTableObject.Exists(sKey) Then
          dcSTableObject.Add sKey, iDum
        End If
      Next
    Next
    For Each oMTable In oMDatabase.Tables
      For Each oMTrigger In oMTable.Triggers
      
        ' Does this trigger exist in slave?
        sKey = CStr(oMTable.Name) & TOKEN1 & CStr(oMTrigger.Name)
        sKey1 = CStr(eDBObjectType & TOKEN1 & sKey)
        If Not dcSTableObject.Exists(sKey1) Then
          sMItem = TRIGGERMISS & sKey
          sSItem = TRIGGERADD & sKey
          AddElementsToLists rsDiffs, sMItem, sSItem, sKey1, oMTable.Name
        Else
          ' Is this trigger different in slave?
          Set oSTrigger = oSDatabase.Tables(oMTable.Name).Triggers(oMTrigger.Name)
          sMScript = Replace(Replace(Replace(oMTrigger.Script, vbCrLf, ""), vbTab, ""), " ", "")
          sSScript = Replace(Replace(Replace(oSTrigger.Script, vbCrLf, ""), vbTab, ""), " ", "")
          If sMScript <> sSScript Then
            sMItem = TRIGGERDIFF & sKey
            sSItem = TRIGGERCHANGE & sKey
            AddElementsToLists rsDiffs, sMItem, sSItem, sKey1
          End If
        End If
      Next
    Next


End Select

GetTableObjectDiffInfo = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Private Function AddTableObject( _
  iObjectType As Integer, _
  sSTableName As String, _
  sSObjectName As String) As String

On Error GoTo EH
ReturnStatus = ""
AddTableObject = ERR_WHILE_FIXING

Select Case iObjectType
  Case 3, 7
    Set oMCheck = oMDatabase.Tables(sSTableName).Checks(sSObjectName)
    sScript = oMCheck.Script(SQLDMOScript_Drops) & " " & oMCheck.Script(SQLDMOScript_Default)
  Case 4, 7
    Set oMIndex = oMDatabase.Tables(sSTableName).Indexes(sSObjectName)
    sScript = oMIndex.Script(SQLDMOScript_Drops) & " " & oMIndex.Script(SQLDMOScript_Default)
  Case 5, 7
    Set oMKey = oMDatabase.Tables(sSTableName).Keys(sSObjectName)
    sScript = oMKey.Script(SQLDMOScript_Drops) & " " & oMKey.Script(SQLDMOScript_Default)
  Case 6, 7
    Set oMTrigger = oMDatabase.Tables(sSTableName).Triggers(sSObjectName)
    sScript = oMTrigger.Script(SQLDMOScript_Drops) & " " & oMTrigger.Script(SQLDMOScript_Default)
End Select


If Trim(sScript) = "" Then
  AddTableObject = UNSCRIPTABLE
  Exit Function
End If
If bScriptFix Then
  TS.Write sScript
  AddTableObject = SCRIPTED
Else
  oSDatabase.ExecuteImmediate sScript
  oSDatabase.StoredProcedures.Refresh
  AddTableObject = FIXED_OK
End If

Exit Function
EH:
AddTableObject = ERR_WHILE_FIXING & Err.Description
Err.Clear

End Function
Private Function GetCodeFileDiffInfo( _
  eTableType As DBTables, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim rsS As ADODB.Recordset
Dim rsM As ADODB.Recordset
Dim sSQL As String
Dim sMKeyVal As String
Dim sSKeyVal As String
Dim sSItem As String
Dim sMItem As String
Dim sKey As String

On Error GoTo EH
GetCodeFileDiffInfo = False
ReturnStatus = ""

Set rsS = New ADODB.Recordset
Set rsM = New ADODB.Recordset

' Get code file records for master and slave
sSQL = "select KeyVal from " & CODE_FILE_TABLE & " order by KeyVal"
rsM.Open sSQL, cnM, adOpenKeyset, adLockReadOnly
If rsM.EOF Then
  GetCodeFileDiffInfo = True
  Exit Function
End If
rsS.Open sSQL, cnS, adOpenKeyset, adLockReadOnly

' If no code file records in slave, copy all
If rsS.EOF Then
  Do Until rsM.EOF
    sMKeyVal = CStr(rsM.Fields("KeyVal"))
    sMItem = CODEFILEMISS & sMKeyVal
    sSItem = CODEFILEADD & sMKeyVal
    sKey = CStr(eTableType) & TOKEN1 & sMKeyVal
    AddElementsToLists rsDiffs, sMItem, sSItem, sKey
    rsM.MoveNext
  Loop
  GetCodeFileDiffInfo = True
  Exit Function
End If

' for each master, see if Key exists in slave
Do Until rsM.EOF
  rsS.MoveFirst
  sMKeyVal = CStr(rsM.Fields("KeyVal"))
  rsS.Find "KeyVal='" & sMKeyVal & "'"
  If rsS.EOF Then
    sMItem = CODEFILEMISS & sMKeyVal
    sSItem = CODEFILEADD & sMKeyVal
    sKey = CStr(eTableType) & TOKEN1 & sMKeyVal
    AddElementsToLists rsDiffs, sMItem, sSItem, sKey
  End If
  rsM.MoveNext
Loop


rsS.Close
rsM.Close
Set rsS = Nothing
Set rsM = Nothing


GetCodeFileDiffInfo = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear


End Function

Private Function GetUserDefDiffInfo( _
  eTableType As DBTables, _
  rsDiffs As ADODB.Recordset) As Boolean

Dim rsS As ADODB.Recordset
Dim rsM As ADODB.Recordset
Dim sSQL As String
Dim sFieldName As String
Dim sSItem As String
Dim sMItem As String
Dim sKey As String

On Error GoTo EH
GetUserDefDiffInfo = False
ReturnStatus = ""

Set rsS = New ADODB.Recordset
Set rsM = New ADODB.Recordset

' Get user def table records for master and slave
sSQL = "select * from " & USER_DEF_TABLE & " order by FieldName"
rsM.Open sSQL, cnM, adOpenKeyset, adLockReadOnly
If rsM.EOF Then
  GetUserDefDiffInfo = True
  Exit Function
End If

' If slave empty, don't check
rsS.Open sSQL, cnS, adOpenKeyset, adLockReadOnly
If rsS.EOF Then
  Err.Raise vbObjectError + 100, "GetUserDefDiffInfo", "No " & USER_DEF_TABLE & " records in slave DB."
  Exit Function
End If

' for each master, see if Key exists in slave
Do Until rsM.EOF
  rsS.MoveFirst
  sFieldName = CStr(rsM.Fields("FieldName"))
  rsS.Find "FieldName='" & sFieldName & "'"
  If Not rsS.EOF Then
    If rsS.Fields("FieldDesc") <> rsM.Fields("FieldDesc") Then
    ' Field Desc and Row must match
      sMItem = USERDEFDIFF & sFieldName & TOKEN1 & rsM.Fields("FieldDesc")
      sSItem = USERDEFCHANGE & sFieldName & TOKEN1 & rsS.Fields("FieldDesc")
      sKey = CStr(eTableType) & TOKEN1 & sFieldName & TOKEN1 & USERDEFFIELD1 & TOKEN1 & rsM.Fields("FieldDesc")
      AddElementsToLists rsDiffs, sMItem, sSItem, sKey
    End If
    If rsS.Fields("Row") <> rsM.Fields("Row") Then
    ' Field Desc and Row must match
      sMItem = USERDEFDIFF & sFieldName & TOKEN1 & rsM.Fields("Row")
      sSItem = USERDEFCHANGE & sFieldName & TOKEN1 & rsS.Fields("Row")
      sKey = CStr(eTableType) & TOKEN1 & sFieldName & TOKEN1 & USERDEFFIELD2 & TOKEN1 & rsM.Fields("Row")
      AddElementsToLists rsDiffs, sMItem, sSItem, sKey
    End If
  End If
  rsM.MoveNext
Loop


rsS.Close
rsM.Close
Set rsS = Nothing
Set rsM = Nothing


GetUserDefDiffInfo = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear


End Function

Private Function ChangeUserDefData( _
  sFieldName As String, _
  sFieldType As String, _
  sFieldData As String) As String

Dim sSQL As String
Dim lRecsAffected As Long
Dim sFieldDesc As String
Dim lRow As Long

On Error GoTo EH
ReturnStatus = ""
ChangeUserDefData = ERR_WHILE_FIXING

If sFieldType = USERDEFFIELD1 Then
  sFieldDesc = Trim(sFieldData)
  sSQL = "Update " & USER_DEF_TABLE & " set FieldDesc='" & sFieldDesc & "' " & _
        "Where FieldName='" & sFieldName & "'"
  If bScriptFix Then
    TS.Write sSQL & vbCrLf
    ChangeUserDefData = SCRIPTED
  Else
    cnS.Execute sSQL, lRecsAffected
    If lRecsAffected <> 1 Then
      Err.Raise vbObjectError + 10, "ChangeUserDefData", "Couldn't update UserDefTable FieldDesc to " & sFieldDesc & " where FieldName = " & sFieldName
      Exit Function
    End If
  End If
End If
  
If sFieldType = USERDEFFIELD2 Then
  lRow = CLng(sFieldData)
  sSQL = "Update " & USER_DEF_TABLE & " set Row=" & lRow & " " & _
        "Where FieldName='" & sFieldName & "'"
  If bScriptFix Then
    TS.Write sSQL & vbCrLf
    ChangeUserDefData = SCRIPTED
  Else
    cnS.Execute sSQL, lRecsAffected
    If lRecsAffected <> 1 Then
      Err.Raise vbObjectError + 10, "ChangeUserDefData", "Couldn't update UserDefTable Row to " & lRow & " where FieldName = " & sFieldName
      Exit Function
    End If
  End If
End If

If ChangeUserDefData <> SCRIPTED Then
  ChangeUserDefData = FIXED_OK
End If

Exit Function
EH:
ChangeUserDefData = ERR_WHILE_FIXING & Err.Description
Err.Clear

End Function

Private Function AddCodeFileKey(sKeyVal As String) As String

Dim sSQL As String
Dim rsS As ADODB.Recordset
Dim rsM As ADODB.Recordset
Dim sScript As String

On Error GoTo EH
AddCodeFileKey = ERR_WHILE_FIXING

Set rsS = New ADODB.Recordset
Set rsM = New ADODB.Recordset

rsM.Open "select * from " & CODE_FILE_TABLE & " where keyval='" & sKeyVal & "'", cnM, adOpenForwardOnly, adLockReadOnly

If bScriptFix Then
  sScript = "Insert into " & CODE_FILE_TABLE & " (KeyVal,CodeDesc,CodeValueInt," & _
    "CodeValueTxt,CodeValueTxt2,CodeValueBoolean,OrderFld,LongText," & _
    "AdminCanChange,FieldUsed) Values " & vbCrLf & "('" & _
    sKeyVal & "'," & _
    FixString(rsM.Fields("CodeDesc")) & "," & _
    IIf(IsNull(rsM.Fields("CodeValueInt")), "null", rsM.Fields("CodeValueInt")) & "," & _
    FixString(rsM.Fields("CodeValueTxt")) & "," & _
    FixString(rsM.Fields("CodeValueTxt2")) & "," & _
    IIf(IsNull(rsM.Fields("CodeValueBoolean")), "null", rsM.Fields("CodeValueBoolean")) & "," & _
    IIf(IsNull(rsM.Fields("OrderFld")), "null", rsM.Fields("OrderFld")) & "," & _
    FixString(rsM.Fields("LongText")) & "," & _
    IIf(IsNull(rsM.Fields("AdminCanChange")), "null", rsM.Fields("AdminCanChange")) & "," & _
    FixString(rsM.Fields("FieldUsed")) & ")" & vbCrLf
  TS.Write sScript
  AddCodeFileKey = SCRIPTED
  Exit Function
Else
  rsS.Open "select * from " & CODE_FILE_TABLE & " where 1=0", cnS, adOpenDynamic, adLockOptimistic
  rsS.AddNew
  rsS.Fields("KeyVal") = sKeyVal
  rsS.Fields("CodeDesc") = rsM.Fields("CodeDesc")
  rsS.Fields("CodeValueInt") = rsM.Fields("CodeValueInt")
  rsS.Fields("CodeValueTxt") = rsM.Fields("CodeValueTxt")
  rsS.Fields("CodeValueTxt2") = rsM.Fields("CodeValueTxt2")
  rsS.Fields("CodeValueBoolean") = rsM.Fields("CodeValueBoolean")
  rsS.Fields("OrderFld") = rsM.Fields("OrderFld")
  rsS.Fields("LongText") = rsM.Fields("LongText")
  rsS.Fields("AdminCanChange") = rsM.Fields("AdminCanChange")
  rsS.Fields("FieldUsed") = rsM.Fields("FieldUsed")
  rsS.Update
  rsS.Close
  Set rsS = Nothing
End If

AddCodeFileKey = FIXED_OK
Exit Function
EH:
ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Private Sub AddElementsToLists( _
  rsList As ADODB.Recordset, _
  sMItem As String, _
  sSItem As String, _
  sKey As String, _
  Optional sTableName As String, _
  Optional sXferData As String)
  
Dim sFilter As String

' If table name given, then check to see if table was added
' If so, don't add this item as it will be added with table

If sTableName <> "" Then
  sFilter = "Slave='" & TABLEADD & sTableName & "'"
  rsList.MoveFirst
  rsList.Find sFilter
  If Not rsList.EOF Then
    Exit Sub
  End If
End If

rsList.AddNew
rsList.Fields("Master") = sMItem
rsList.Fields("Slave") = sSItem
rsList.Fields("Fix") = "Yes"
rsList.Fields("XferData") = sXferData
rsList.Fields("FixKey") = sKey
  
End Sub

Private Function ReadyToProc() As Boolean

On Error GoTo EH
ReturnStatus = ""
ReadyToProc = False

If oMServer.ConnectionID <> lMConnID Then
  ReturnStatus = "Connections to Master server lost.  Reconnect and try again."
  Exit Function
End If
If oSServer.ConnectionID <> lSConnID Then
  ReturnStatus = "Connections to Slave server lost.  Reconnect and try again."
  Exit Function
End If


ReadyToProc = True
Exit Function
EH:
ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function


Private Sub CreateDiffsRecordset( _
  eDiffType As DiffType, _
  rsDiffs As ADODB.Recordset)

If rsDiffs.Fields.Count = 0 Then

  Set rsDiffs = Nothing
  Set rsDiffs = New ADODB.Recordset
  
  Select Case eDiffType
    Case 0, 1
  rsDiffs.Fields.Append "Master", adVarChar, 200, adFldUpdatable
  rsDiffs.Fields.Append "Fix", adVarChar, 5, adFldUpdatable
  rsDiffs.Fields.Append "XferData", adVarChar, 5, adFldUpdatable
  rsDiffs.Fields.Append "Slave", adVarChar, 200, adFldUpdatable
  rsDiffs.Fields.Append "Status", adVarChar, 200, adFldUpdatable
  rsDiffs.Fields.Append "FixKey", adVarChar, 200, adFldUpdatable
    
  End Select
  rsDiffs.Open
End If

End Sub

Public Function SaveDatabaseObjectScripts( _
  sKey As String, _
  sMasterFile As String, _
  sSlaveFile As String) As Boolean

Dim sMScript As String
Dim sSScript As String
Dim sData() As String
Dim sObjectName As String
Dim sSubObjectName As String
Dim iObjectType As Integer

On Error GoTo EH
ReturnStatus = ""
SaveDatabaseObjectScripts = False

sData = Split(sKey, TOKEN1)
iObjectType = CInt(sData(0))
sObjectName = sData(1)
If UBound(sData) > 1 Then
  sSubObjectName = sData(2)
End If

Select Case iObjectType

  Case 0
    Set oMTable = oMDatabase.Tables(sObjectName)
    Set oSTable = oSDatabase.Tables(sObjectName)
    sMScript = oMTable.Script
    sSScript = oSTable.Script
  Case 1
    Set oMView = oMDatabase.Views(sObjectName)
    Set oSView = oSDatabase.Views(sObjectName)
    sMScript = oMView.Script
    sSScript = oSView.Script
  Case 2
    Set oMProc = oMDatabase.StoredProcedures(sObjectName)
    Set oSProc = oSDatabase.StoredProcedures(sObjectName)
    sMScript = oMProc.Script
    sSScript = oSProc.Script
  Case 3
    Set oMCheck = oMDatabase.Tables(sObjectName).Checks(sSubObjectName)
    Set oSCheck = oSDatabase.Tables(sObjectName).Checks(sSubObjectName)
    sMScript = oMCheck.Script
    sSScript = oSCheck.Script
  Case 4
    Set oMIndex = oMDatabase.Tables(sObjectName).Indexes(sSubObjectName)
    Set oSIndex = oSDatabase.Tables(sObjectName).Indexes(sSubObjectName)
    sMScript = oMIndex.Script
    sSScript = oSIndex.Script
  Case 5
    Set oMKey = oMDatabase.Tables(sObjectName).Keys(sSubObjectName)
    Set oSKey = oSDatabase.Tables(sObjectName).Keys(sSubObjectName)
    sMScript = oMKey.Script
    sSScript = oSKey.Script
  Case 6
    Set oMTrigger = oMDatabase.Tables(sObjectName).Triggers(sSubObjectName)
    Set oSTrigger = oSDatabase.Tables(sObjectName).Triggers(sSubObjectName)
    sMScript = oMTrigger.Script
    sSScript = oSTrigger.Script

End Select

If (sSScript = "") And (sMScript = "") Then
  Exit Function
End If

sMasterFile = TMP_ROOT_DIR & "MasterObject.sql"
sSlaveFile = TMP_ROOT_DIR & "SlaveObject.sql"
Set TS = FSO.OpenTextFile(sMasterFile, ForWriting, True)
TS.Write sMScript
TS.Close
Set TS = FSO.OpenTextFile(sSlaveFile, ForWriting, True)
TS.Write sSScript
TS.Close

SaveDatabaseObjectScripts = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function
Public Function SaveIndexesToFiles( _
  sKey As String, _
  sMasterFile As String, _
  sSlaveFile As String) As Boolean

Dim sScript As String
Dim sData() As String

On Error GoTo EH
ReturnStatus = ""
SaveIndexesToFiles = False

sData = Split(sKey, TOKEN1)

Set oMIndex = oMDatabase.Tables(sData(0)).Indexes(sData(1))
Set oSIndex = oSDatabase.Tables(sData(0)).Indexes(sData(1))

If Not ((oMIndex Is Nothing) Or (oSIndex Is Nothing)) Then
  sMasterFile = TMP_ROOT_DIR & "MasterIndex.sql"
  sSlaveFile = TMP_ROOT_DIR & "SlaveIndex.sql"
  Set TS = FSO.OpenTextFile(sMasterFile, ForWriting, True)
  sScript = oMIndex.Script
  TS.Write sScript
  TS.Close
  Set TS = FSO.OpenTextFile(sSlaveFile, ForWriting, True)
  sScript = oSIndex.Script
  TS.Write sScript
  TS.Close
End If

SaveIndexesToFiles = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function
Public Function SaveStoredProceduresToFiles( _
  sProcName As String, _
  sMasterFile As String, _
  sSlaveFile As String) As Boolean

Dim sScript As String

On Error GoTo EH
ReturnStatus = ""
SaveStoredProceduresToFiles = False

Set oMProc = oMDatabase.StoredProcedures(sProcName)
Set oSProc = oSDatabase.StoredProcedures(sProcName)

If Not ((oMProc Is Nothing) Or (oSProc Is Nothing)) Then
  sMasterFile = TMP_ROOT_DIR & "MasterStoredProc.sql"
  sSlaveFile = TMP_ROOT_DIR & "SlaveStoredProc.sql"
  Set TS = FSO.OpenTextFile(sMasterFile, ForWriting, True)
  sScript = oMProc.Script
  TS.Write sScript
  TS.Close
  Set TS = FSO.OpenTextFile(sSlaveFile, ForWriting, True)
  sScript = oSProc.Script
  TS.Write sScript
  TS.Close
End If

SaveStoredProceduresToFiles = True
Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function

Public Function OpenADODBConnections() As Boolean

On Error GoTo EH
OpenADODBConnections = False
ReturnStatus = ""


OpenADODBConnections = True

Exit Function

EH:

ReturnStatus = "Error: " & Err.Description
Err.Clear

End Function

Public Function FixString(vInp As Variant) As String
FixString = ""

If IsNull(vInp) Then
  FixString = "null"
Else
  FixString = "'" & Replace(CStr(vInp), "'", "''") & "'"
End If

End Function

