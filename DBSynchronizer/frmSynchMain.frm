VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSynchMain 
   Caption         =   "DB Syncronizer"
   ClientHeight    =   7725
   ClientLeft      =   585
   ClientTop       =   765
   ClientWidth     =   10575
   Icon            =   "frmSynchMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdAllDiffs 
      Caption         =   "&Both Diffs"
      Height          =   495
      Left            =   5520
      TabIndex        =   32
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdTableDiffs 
      Caption         =   "&Table Diffs"
      Height          =   495
      Left            =   4680
      TabIndex        =   31
      Top             =   7200
      Width           =   735
   End
   Begin VB.ComboBox cboTableFields 
      Height          =   315
      Left            =   2880
      TabIndex        =   30
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Grid"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   7200
      Width           =   855
   End
   Begin MSComDlg.CommonDialog DLG 
      Left            =   6360
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboDBObject 
      Height          =   315
      Left            =   120
      TabIndex        =   26
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "&Fix"
      Height          =   495
      Left            =   8760
      TabIndex        =   25
      Top             =   7200
      Width           =   495
   End
   Begin VB.CommandButton cmdGetDiffs 
      Caption         =   "&Object Diffs"
      Height          =   495
      Left            =   2040
      TabIndex        =   22
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton cmdScript 
      Caption         =   "S&cript"
      Height          =   495
      Left            =   8040
      TabIndex        =   21
      Top             =   7200
      Width           =   615
   End
   Begin VB.Frame fraSlave 
      Caption         =   "Slave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5280
      TabIndex        =   14
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdSConnect 
         Caption         =   "Connect/Retrieve Databases"
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtSUserID 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtSPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox cboSDatabase 
         Height          =   315
         Left            =   3480
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtSServer 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "User ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Database:"
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      Top             =   7200
      Width           =   615
   End
   Begin VB.Frame fraMaster 
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      Begin VB.VScrollBar VScroll1 
         Height          =   30
         Left            =   4680
         TabIndex        =   23
         Top             =   4440
         Width           =   135
      End
      Begin VB.TextBox txtMServer 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cboMDatabase 
         Height          =   315
         Left            =   3480
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtMPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtMUserID 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdMConnect 
         Caption         =   "Connect/Retrieve Databases"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Database:"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "User ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDiff 
      Height          =   5055
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8916
      _Version        =   393216
      FocusRect       =   2
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label10 
      Caption         =   "Table/Fields"
      Height          =   255
      Left            =   2880
      TabIndex        =   29
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Database Objects"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuIgnoreSysObjects 
         Caption         =   "Ignore &System Objects"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIgnore_Indexes 
         Caption         =   "Ignore &Indexes Beginning With Underscore"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseBCP 
         Caption         =   "Use BCP To Transfer Data"
      End
   End
End
Attribute VB_Name = "frmSynchMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oDBSynch As clsDBSynch
Public rsList As ADODB.Recordset
Public bProcDBObjects As Boolean
Public bProcAll As Boolean

Private Sub cboMDatabase_LostFocus()

If Trim(cboMDatabase.Text) <> "" Then
  oDBSynch.sMDBName = cboMDatabase.Text
  If Not oDBSynch.SetMasterDatabase() Then
    MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Setting Database"
  End If
End If
End Sub


Private Sub cboSDatabase_LostFocus()

If Trim(cboSDatabase.Text) <> "" Then
  oDBSynch.sSDBName = cboSDatabase.Text
  If Not oDBSynch.SetSlaveDatabase() Then
    MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Setting Database"
  End If
End If
End Sub


Private Sub cmdAllDiffs_Click()

If Not CheckAndResetDB Then
  Exit Sub
End If


Set rsList = New ADODB.Recordset
If Not oDBSynch.GetDBObjectDiffInfo(CInt(cboDBObject.ItemData(cboDBObject.ListIndex)), rsList) Then
  frmSynchMain.MousePointer = vbDefault
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Getting Differences "
End If

If Not oDBSynch.GetTableFieldDiffInfo(CInt(cboTableFields.ItemData(cboTableFields.ListIndex)), rsList) Then
  frmSynchMain.MousePointer = vbDefault
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Getting Differences "
  Exit Sub
End If

Set flxDiff.DataSource = rsList
If (cboDBObject.Text = "Table/Fields") Or _
  (cboDBObject.Text = "All") Then
  SetGridProperties True
Else
  SetGridProperties False
End If


flxDiff.Refresh
frmSynchMain.MousePointer = vbDefault

cmdFix.Enabled = True
cmdScript.Enabled = True
bProcAll = True


End Sub

Private Sub cmdExit_Click()

Unload frmSynchMain
Set frmSynchMain = Nothing

End Sub

Private Sub cmdFix_Click()

frmSynchMain.MousePointer = vbHourglass

If bProcDBObjects Or bProcAll Then
  If Not oDBSynch.FixDBObjectDiffInfo(rsList, False) Then
    MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Fixing Differences"
  End If
End If
If (Not bProcDBObjects) Or bProcAll Then
  If Not oDBSynch.FixTableFieldsDiffInfo(rsList, bProcAll) Then
    MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Fixing Differences"
  End If
End If

rsList.MoveFirst
Set flxDiff.DataSource = Nothing
Set flxDiff.DataSource = rsList

bProcAll = False
frmSynchMain.MousePointer = vbDefault

End Sub


Private Sub cmdGetDiffs_Click()

If Not CheckAndResetDB Then
  Exit Sub
End If


Set rsList = New ADODB.Recordset
If Not oDBSynch.GetDBObjectDiffInfo(CInt(cboDBObject.ItemData(cboDBObject.ListIndex)), rsList) Then
  frmSynchMain.MousePointer = vbDefault
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Getting Differences "
  Exit Sub
End If

Set flxDiff.DataSource = rsList
If (cboDBObject.Text = "Table/Fields") Or _
  (cboDBObject.Text = "All") Then
  SetGridProperties True
Else
  SetGridProperties False
End If


flxDiff.Refresh
frmSynchMain.MousePointer = vbDefault

cmdFix.Enabled = True
cmdScript.Enabled = True
bProcDBObjects = True
bProcAll = False

End Sub
Private Sub cmdTableDiffs_Click()

If Not CheckAndResetDB Then
  Exit Sub
End If


Set rsList = New ADODB.Recordset
If Not oDBSynch.GetTableFieldDiffInfo(CInt(cboTableFields.ItemData(cboTableFields.ListIndex)), rsList) Then
  frmSynchMain.MousePointer = vbDefault
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Getting Differences "
  Exit Sub
End If

Set flxDiff.DataSource = rsList
If (cboDBObject.Text = "Table/Fields") Or _
  (cboDBObject.Text = "All") Then
  SetGridProperties True
Else
  SetGridProperties False
End If


flxDiff.Refresh
frmSynchMain.MousePointer = vbDefault

cmdFix.Enabled = True
cmdScript.Enabled = True
bProcDBObjects = False
bProcAll = False

End Sub

Private Sub cmdMConnect_Click()

Dim sDatabases() As String
Dim iCount As Integer

oDBSynch.sMUser = txtMUserID.Text
oDBSynch.sMPass = txtMPassword
oDBSynch.sMServer = txtMServer

frmSynchMain.MousePointer = vbHourglass

If Not oDBSynch.ConnectMaster Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error"
  frmSynchMain.MousePointer = vbDefault
  Exit Sub
End If

If Not oDBSynch.GetDatabases(oDBSynch.oMServer, sDatabases) Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error"
  frmSynchMain.MousePointer = vbDefault
  Exit Sub
End If

cboMDatabase.Clear
For iCount = 0 To UBound(sDatabases) - 1
  cboMDatabase.AddItem sDatabases(iCount)
Next iCount

frmSynchMain.MousePointer = vbDefault

End Sub

Private Sub cmdSConnect_Click()

Dim sDatabases() As String
Dim iCount As Integer

oDBSynch.sSUser = txtSUserID.Text
oDBSynch.sSPass = txtSPassword.Text
oDBSynch.sSServer = txtSServer.Text

frmSynchMain.MousePointer = vbHourglass

If Not oDBSynch.ConnectSlave Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error"
  frmSynchMain.MousePointer = vbDefault
  Exit Sub
End If

If Not oDBSynch.GetDatabases(oDBSynch.oSServer, sDatabases) Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error"
  frmSynchMain.MousePointer = vbDefault
  Exit Sub
End If

cboSDatabase.Clear
For iCount = 0 To UBound(sDatabases) - 1
  cboSDatabase.AddItem sDatabases(iCount)
Next iCount

frmSynchMain.MousePointer = vbDefault

End Sub

Private Sub cmdScript_Click()

Dim sScriptFile As String

On Error GoTo EH


With Me.DLG

    .CancelError = True
    .DefaultExt = "sql"
    .FileName = "DBSynch"
    .Filter = "*.sql"
    .ShowSave
    sScriptFile = .FileName

End With

If sScriptFile <> "" Then
  oDBSynch.bScriptFix = True
  oDBSynch.sScriptFile = sScriptFile
  frmSynchMain.MousePointer = vbHourglass
  cmdFix_Click
End If
oDBSynch.bScriptFix = False

Exit Sub
EH:
' This error number is set if the user selects "Cancel"
If Err.Number = 32755 Then
    Exit Sub
End If
End Sub


Private Sub flxDiff_DblClick()

Dim lRow As Long
Dim lCol As Long
Dim sMasterFile As String
Dim sSlaveFile As String
Dim sCmd As String

lCol = flxDiff.Col


Select Case lCol

  Case 2, 3
    If flxDiff.Text <> "" Then
      If flxDiff.Text = "Yes" Then
        flxDiff.Text = "No"
      Else
        flxDiff.Text = "Yes"
      End If
      flxDiff.Col = 1
      rsList.MoveFirst
      Do Until rsList.EOF
        If Trim(rsList.Fields("Master")) = Trim(flxDiff.Text) Then
          flxDiff.Col = lCol
          If lCol = 2 Then
            rsList.Fields("Fix") = flxDiff.Text
          Else
            rsList.Fields("XferData") = flxDiff.Text
          End If
          rsList.Update
          Exit Do
        End If
        rsList.MoveNext
      Loop
      flxDiff.Refresh
    End If
  
  Case 4
    flxDiff.Col = 6
    If oDBSynch.SaveDatabaseObjectScripts(flxDiff.Text, sMasterFile, sSlaveFile) Then
      sCmd = App.Path & "\windiff.exe " & sMasterFile & " " & sSlaveFile
      Shell sCmd, vbNormalFocus
    End If

  Case 5
    If Trim(flxDiff.Text) <> "" Then
      MsgBox flxDiff.Text, vbOKOnly
    End If

End Select


End Sub

Private Sub Form_Load()

Dim iCount As Integer
Set oDBSynch = New clsDBSynch

oDBSynch.bScriptFix = False
oDBSynch.bUseSystemObjects = False
If Me.mnuIgnoreSysObjects.Checked = False Then
  oDBSynch.bUseSystemObjects = True
End If
oDBSynch.bUseIndexesWithUnderscore = False
If Me.mnuIgnore_Indexes.Checked = False Then
  oDBSynch.bUseIndexesWithUnderscore = True
End If
oDBSynch.bUseBulkCopyForDataXfer = False
If Me.mnuUseBCP.Checked = True Then
  oDBSynch.bUseBulkCopyForDataXfer = True
End If

SetGridProperties False
For iCount = 0 To UBound(oDBSynch.vDBObjectsList) - 1
  cboDBObject.AddItem oDBSynch.vDBObjectsList(iCount)
  cboDBObject.ItemData(iCount) = iCount
Next iCount
cboDBObject.ListIndex = 0

For iCount = 0 To UBound(oDBSynch.vDBTablesList) - 10
  cboTableFields.AddItem oDBSynch.vDBTablesList(iCount + 10)
  cboTableFields.ItemData(iCount) = iCount + 10
Next iCount
cboTableFields.ListIndex = 0

cmdFix.Enabled = False
cmdScript.Enabled = False
bProcDBObjects = True
End Sub

Private Sub Form_Terminate()

Set oDBSynch = Nothing

End Sub

Function CheckAndResetDB() As Boolean

CheckAndResetDB = True
If (cboMDatabase.ListIndex = -1) Or (cboSDatabase.ListIndex = -1) Then
  MsgBox "You must select a master and slave database.", vbOKOnly, "Missing Database"
  CheckAndResetDB = False
  Exit Function
End If
frmSynchMain.MousePointer = vbHourglass
If Not oDBSynch.ConnectMaster Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error"
  frmSynchMain.MousePointer = vbDefault
  Exit Function
End If
If Not oDBSynch.ConnectSlave Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error"
  frmSynchMain.MousePointer = vbDefault
  Exit Function
End If
If Not oDBSynch.SetMasterDatabase() Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Setting Database"
End If
If Not oDBSynch.SetSlaveDatabase() Then
  MsgBox oDBSynch.ReturnStatus, vbOKOnly, "Error Setting Database"
End If


End Function

Private Sub SetGridProperties(bShowDataXferColumn As Boolean)

If bShowDataXferColumn Then
  flxDiff.ColWidth(0) = 0
  flxDiff.ColWidth(1) = 3500
  flxDiff.ColWidth(2) = 400
  flxDiff.ColWidth(3) = 800
  flxDiff.ColWidth(4) = 3500
  flxDiff.ColWidth(5) = 2000
  flxDiff.ColWidth(6) = 0
  flxDiff.Refresh
Else
  flxDiff.ColWidth(0) = 0
  flxDiff.ColWidth(1) = 3800
  flxDiff.ColWidth(2) = 400
  flxDiff.ColWidth(3) = 0
  flxDiff.ColWidth(4) = 3800
  flxDiff.ColWidth(5) = 2200
  flxDiff.ColWidth(6) = 0
End If

End Sub

Private Sub mnuExit_Click()

cmdExit_Click

End Sub

Private Sub mnuIgnore_Indexes_Click()

If mnuIgnore_Indexes.Checked = False Then
  mnuIgnore_Indexes.Checked = True
  oDBSynch.bUseIndexesWithUnderscore = False
Else
  mnuIgnore_Indexes.Checked = False
  oDBSynch.bUseIndexesWithUnderscore = True
End If

End Sub

Private Sub mnuIgnoreSysObjects_Click()

If mnuIgnoreSysObjects.Checked = False Then
  oDBSynch.bUseSystemObjects = False
Else
  mnuIgnoreSysObjects.Checked = False
  oDBSynch.bUseSystemObjects = True
End If


End Sub

Private Sub mnuUseBCP_Click()

If mnuUseBCP.Checked = False Then
  oDBSynch.bUseBulkCopyForDataXfer = False
Else
  oDBSynch.bUseBulkCopyForDataXfer = True
End If

End Sub
