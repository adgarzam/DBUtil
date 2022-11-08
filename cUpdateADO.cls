VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUpdateADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' #VBIDEUtils#************************************************************
' * Author           : Larry Rebich
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           : larry@buygold.net
' * Date             : 12/10/2003
' * Purpose          :
' * Project Name     : DBUpdateADO
' * Module Name      : cUpdateADO
' **********************************************************************
' * Comments         :
' *
' *
' * Example          :
' *
' * History          : Updated by Waty Thierry
' * 2003/05/19 Copyright © 2003, Larry Rebich, using the DELL7500
' * 2003/05/19 larry@larryrebich.com, www.larryrebich.com, 760-771-4730
' * 2003/12/11 Updated by Waty Thierry to add fields in right order
' *            and some other modifications, optimization...
' *            Better Error Handling
' *
' * See Also         :
' *
' *
' **********************************************************************

Option Explicit
DefLng A-Z
'

Private mConnDB         As New ADODB.Connection     'database to be updated
Private mConnModel      As New ADODB.Connection  'model database

Private tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX

Private mbAnyErrors     As Boolean  '2003/05/30
Private mbSkipUpdateRelationships As Boolean  '2003/05/30
Private mbSkipWriteToLog As Boolean  'don't write results log
Private msFileNameLog   As String   'fully qualified names
Private msFileNameErrorLog As String
Private msAppPath       As String

Private mcResultsTables As New cResults
Private mcResultsColumns As New cResults
Private mcResultsIndexes As New cResults
Private mcResultsRelationships As New cResults
Private mcResultsStoredProcedures As New cResults
Private mcResultsViews  As New cResults

Private mbUpdateOrUpdateTablesPerformed As Boolean  '2003/05/30 Added, must have done update before updating relationships
'

Private Sub Class_Initialize()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : Class_Initialize
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Me.sAppPath = AddBackslash(App.Path)
   gsAppPath = Me.sAppPath
End Sub

Public Property Let bAnyErrors(bAnyErrors As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : bAnyErrors
   ' * Parameters       :
   ' *                    bAnyErrors As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Property created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   mbAnyErrors = bAnyErrors
End Property

Public Property Get bAnyErrors() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : bAnyErrors
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   bAnyErrors = mbAnyErrors
End Property

Public Property Let bSkipUpdateRelationships(bSkipUpdateRelationships As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : bSkipUpdateRelationships
   ' * Parameters       :
   ' *                    bSkipUpdateRelationships As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Property bSkipUpdateRelationships created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   mbSkipUpdateRelationships = bSkipUpdateRelationships
End Property

Public Property Get bSkipUpdateRelationships() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : bSkipUpdateRelationships
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   bSkipUpdateRelationships = mbSkipUpdateRelationships
End Property

Public Property Let connDB(connDB As ADODB.Connection)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : connDB
   ' * Parameters       :
   ' *                    connDB As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set mConnDB = connDB
End Property

Public Property Get connDB() As ADODB.Connection
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : connDB
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set connDB = mConnDB
End Property

Public Property Let connModel(connModel As ADODB.Connection)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : connModel
   ' * Parameters       :
   ' *                    connModel As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set mConnModel = connModel
End Property

Public Property Get connModel() As ADODB.Connection
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : connModel
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set connModel = mConnModel
End Property

Public Property Get cResultsTables() As cResults
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : cResultsTables
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set cResultsTables = mcResultsTables
End Property

Public Property Get cResultsColumns() As cResults
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : cResultsColumns
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set cResultsColumns = mcResultsColumns
End Property

Public Property Get cResultsIndexes() As cResults
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : cResultsIndexes
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set cResultsIndexes = mcResultsIndexes
End Property

Public Property Get cResultsRelationships() As cResults
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : cResultsRelationships
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set cResultsRelationships = mcResultsRelationships
End Property

Public Property Get cResultsStoredProcedures() As cResults
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : cResultsStoredProcedures
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set cResultsStoredProcedures = mcResultsStoredProcedures
End Property

Public Property Get cResultsViews() As cResults
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : cResultsViews
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Set cResultsViews = mcResultsViews
End Property

Public Property Let sAppPath(sAppPath As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : sAppPath
   ' * Parameters       :
   ' *                    sAppPath As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   msAppPath = sAppPath
   gsAppPath = sAppPath
End Property

Public Property Get sAppPath() As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : sAppPath
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   sAppPath = msAppPath
End Property

Public Property Let sFileNameLog(sFileNameLog As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : sFileNameLog
   ' * Parameters       :
   ' *                    sFileNameLog As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   msFileNameLog = sFileNameLog
End Property

Public Property Get sFileNameLog() As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : sFileNameLog
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   sFileNameLog = msFileNameLog
End Property

Public Property Let sFileNameErrorLog(sFileNameErrorLog As String)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : sFileNameErrorLog
   ' * Parameters       :
   ' *                    sFileNameErrorLog As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   msFileNameErrorLog = sFileNameErrorLog
End Property

Public Property Get sFileNameErrorLog() As String
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : sFileNameErrorLog
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   sFileNameErrorLog = msFileNameErrorLog
End Property

Public Property Let bSkipWriteToLog(bSkipWriteToLog As Boolean)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : bSkipWriteToLog
   ' * Parameters       :
   ' *                    bSkipWriteToLog As Boolean
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   mbSkipWriteToLog = bSkipWriteToLog
End Property

Public Property Get bSkipWriteToLog() As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : bSkipWriteToLog
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   bSkipWriteToLog = mbSkipWriteToLog
End Property

Public Function Update(Optional oconnDB As ADODB.Connection, Optional oconnModel As ADODB.Connection, _
   Optional vntSkipUpdateRelationships As Variant) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : Update
   ' * Parameters       :
   ' *                    Optional oconnDB As ADODB.Connection
   ' *                    Optional oconnModel As ADODB.Connection
   ' *                    Optional vntSkipUpdateRelationships As Variant
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   If Not IsMissing(oconnDB) Then
      connDB = oconnDB
   End If
   If Not IsMissing(oconnModel) Then
      connModel = oconnModel
   End If
   If Not IsMissing(vntSkipUpdateRelationships) Then
      bSkipUpdateRelationships = vntSkipUpdateRelationships
   End If

   PassObjectDataToStructure
   On Error GoTo UpdateEH
   If UpdateDatabaseWithModelUsingADOX(mConnDB, mConnModel, tUpdateDatabaseWithModelUsingADOX) Then
      PassStructureResultsToObject tUpdateDatabaseWithModelUsingADOX
      mbUpdateOrUpdateTablesPerformed = True
      Update = True
   End If
   Exit Function
UpdateEH:
   Err.Raise Err.Number, "cUpdateADO:Update", Err.Description
End Function

Private Sub PassStructureResultsToObject(ByRef tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX)
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : PassStructureResultsToObject
   ' * Parameters       :
   ' *                    ByRef tUpdateDatabaseWithModelUsingADOX As udtUpdateDatabaseWithModelUsingADOX
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   '2003/05/19 Pass the results from the structure to the object
   If Me.sFileNameLog = "" Then        '2003/05/22
      Me.sFileNameLog = tUpdateDatabaseWithModelUsingADOX.sFileNameLog
   End If
   If Me.sFileNameErrorLog = "" Then   '2003/05/22
      Me.sFileNameErrorLog = tUpdateDatabaseWithModelUsingADOX.sFileNameErrorLog
   End If
   If Me.sAppPath = "" Then            '2003/05/22
      Me.sAppPath = tUpdateDatabaseWithModelUsingADOX.sAppPath
   End If
   With mcResultsTables
      .iAdded = tUpdateDatabaseWithModelUsingADOX.iAddedTables
   End With
   With mcResultsColumns
      .iAdded = tUpdateDatabaseWithModelUsingADOX.iAddedColumns
   End With
   With mcResultsIndexes
      .iAdded = tUpdateDatabaseWithModelUsingADOX.tIndexes.iAdded
      .iReplaced = tUpdateDatabaseWithModelUsingADOX.tIndexes.iReplaced
      .iSkipped = tUpdateDatabaseWithModelUsingADOX.tIndexes.iSkipped
   End With
   With mcResultsRelationships
      .iAdded = tUpdateDatabaseWithModelUsingADOX.tRelationships.iAdded
      .iReplaced = tUpdateDatabaseWithModelUsingADOX.tRelationships.iReplaced
      .iSkipped = tUpdateDatabaseWithModelUsingADOX.tRelationships.iSkipped
   End With
   With mcResultsStoredProcedures
      .iAdded = tUpdateDatabaseWithModelUsingADOX.tStoredProcedures.iAdded
      .iReplaced = tUpdateDatabaseWithModelUsingADOX.tStoredProcedures.iReplaced
      .iSkipped = tUpdateDatabaseWithModelUsingADOX.tStoredProcedures.iSkipped
   End With
   With mcResultsViews
      .iAdded = tUpdateDatabaseWithModelUsingADOX.tViews.iAdded
      .iReplaced = tUpdateDatabaseWithModelUsingADOX.tViews.iReplaced
      .iSkipped = tUpdateDatabaseWithModelUsingADOX.tViews.iSkipped
   End With
   '2003/05/30 Added
   Me.bSkipUpdateRelationships = tUpdateDatabaseWithModelUsingADOX.bSkipUpdateRelationships
   Me.bAnyErrors = tUpdateDatabaseWithModelUsingADOX.bAnyErrors
End Sub

Private Sub PassObjectDataToStructure()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : PassObjectDataToStructure
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ' 2003/05/19 Pass data from object to structure
   tUpdateDatabaseWithModelUsingADOX.bSkipWriteToLog = Me.bSkipWriteToLog
   If Me.sFileNameLog <> "" Then
      modUpdateDatabaseWithModelUsingADOX.gsFileNameLog = Me.sFileNameLog
   End If
   If Me.sFileNameErrorLog <> "" Then
      modUpdateDatabaseWithModelUsingADOX.gsFileNameErrorLog = Me.sFileNameErrorLog
   End If
   tUpdateDatabaseWithModelUsingADOX.bSkipUpdateRelationships = Me.bSkipUpdateRelationships
   
End Sub

Public Function UpdateRelationships(Optional oconnDB As ADODB.Connection, Optional oconnModel As ADODB.Connection) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : UpdateRelationships
   ' * Parameters       :
   ' *                    Optional oconnDB As ADODB.Connection
   ' *                    Optional oconnModel As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   If Not IsMissing(oconnDB) Then
      If Not (oconnDB Is Nothing) Then
         connDB = oconnDB
      End If
   End If
   If Not IsMissing(oconnModel) Then
      If Not (oconnModel Is Nothing) Then
         connModel = oconnModel
      End If
   End If

   If mbUpdateOrUpdateTablesPerformed Then
      PassObjectDataToStructure
      If UpdateRelationshipsWithModelUsingADOX(mConnDB, mConnModel, tUpdateDatabaseWithModelUsingADOX) Then
         PassStructureResultsToObject tUpdateDatabaseWithModelUsingADOX
         UpdateRelationships = True
      End If
   Else
      Err.Raise vbObjectError + 4, "cUpdateADO:UpdateRelationships", "Update or UpdateTables needs to be performed first"
   End If
   
End Function

Public Function UpdateTables(Optional oconnDB As ADODB.Connection, Optional oconnModel As ADODB.Connection) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : UpdateTables
   ' * Parameters       :
   ' *                    Optional oconnDB As ADODB.Connection
   ' *                    Optional oconnModel As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   If Not IsMissing(oconnDB) Then
      If Not (oconnDB Is Nothing) Then
         connDB = oconnDB
      End If
   End If
   If Not IsMissing(oconnModel) Then
      If Not (oconnModel Is Nothing) Then
         connModel = oconnModel
      End If
   End If

   PassObjectDataToStructure

   If UpdateTablesWithModelUsingADOX(mConnDB, mConnModel, tUpdateDatabaseWithModelUsingADOX) Then
      mbUpdateOrUpdateTablesPerformed = True
      PassStructureResultsToObject tUpdateDatabaseWithModelUsingADOX
      UpdateTables = True
   End If
End Function

Public Function UpdateIndexes(Optional oconnDB As ADODB.Connection, Optional oconnModel As ADODB.Connection) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : UpdateIndexes
   ' * Parameters       :
   ' *                    Optional oconnDB As ADODB.Connection
   ' *                    Optional oconnModel As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   If Not IsMissing(oconnDB) Then
      If Not (oconnDB Is Nothing) Then
         connDB = oconnDB
      End If
   End If
   If Not IsMissing(oconnModel) Then
      If Not (oconnModel Is Nothing) Then
         connModel = oconnModel
      End If
   End If

   If mbUpdateOrUpdateTablesPerformed Then
      PassObjectDataToStructure

      If UpdateIndexesWithModelUsingADOX(mConnDB, mConnModel, tUpdateDatabaseWithModelUsingADOX) Then
         PassStructureResultsToObject tUpdateDatabaseWithModelUsingADOX
         UpdateIndexes = True
      End If
   Else
      Err.Raise vbObjectError + 2, "cUpdateADO:UpdateIndexes", "Update or UpdateTables needs to be performed first"
   End If

End Function

Public Function UpdateStoredProceduresAndViews(Optional oconnDB As ADODB.Connection, Optional oconnModel As ADODB.Connection) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : UpdateStoredProceduresAndViews
   ' * Parameters       :
   ' *                    Optional oconnDB As ADODB.Connection
   ' *                    Optional oconnModel As ADODB.Connection
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/05/30 Function created by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   If Not IsMissing(oconnDB) Then
      If Not (oconnDB Is Nothing) Then
         connDB = oconnDB
      End If
   End If
   If Not IsMissing(oconnModel) Then
      If Not (oconnModel Is Nothing) Then
         connModel = oconnModel
      End If
   End If

   If mbUpdateOrUpdateTablesPerformed Then
      PassObjectDataToStructure

      If UpdateStoredProceduresAndViewsWithModelUsingADOX(mConnDB, mConnModel, tUpdateDatabaseWithModelUsingADOX) Then
         PassStructureResultsToObject tUpdateDatabaseWithModelUsingADOX
         UpdateStoredProceduresAndViews = True
      End If
   Else
      Err.Raise vbObjectError + 3, "cUpdateADO:UpdateStoredProceduresAndViews", "Update or UpdateTables needs to be performed first"
   End If
End Function

Private Sub Class_Terminate()
   ' #VBIDEUtils#************************************************************
   ' * Author           : Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : larry@buygold.net
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : Class_Terminate
   ' * Parameters       :
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   ReportErrorsADOX tUpdateDatabaseWithModelUsingADOX
End Sub

Public Function DropTable(sConnectionDBOut As ADODB.Connection, sTable As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Waty Thierry
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : DropTable
   ' * Parameters       :
   ' *                    sConnectionDBOut As ADODB.Connection
   ' *                    sTable As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim oConnection      As ADODB.Connection
   Dim sSQL             As String

   If Not sConnectionDBOut Is Nothing Then
      sSQL = "Drop Table [" & sTable & "]"
            
      Set oConnection = New ADODB.Parameter
      oConnection.Open sConnectionDBOut
      oConnection.Execute sSQL
      oConnection.Close
      Set oConnection = Nothing
      
      DropTable = True
      Exit Function
   End If

End Function

Public Function DeleteData(sConnectionDBOut As ADODB.Connection, sTable As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Waty Thierry
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : DeleteData
   ' * Parameters       :
   ' *                    sConnectionDBOut As ADODB.Connection
   ' *                    sTable As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim oConnection      As ADODB.Connection
   Dim sSQL             As String

   If Not sConnectionDBOut Is Nothing Then
      sSQL = "Delete From [" & sTable & "]"
            
      Set oConnection = New ADODB.Parameter
      oConnection.Open sConnectionDBOut
      oConnection.Execute sSQL
      oConnection.Close
      Set oConnection = Nothing
      
      DeleteData = True
      Exit Function
   End If

End Function

Public Function CopyData(sConnectionDBIn As String, sConnectionDBOut As String, sTable As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Waty Thierry
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : CopyData
   ' * Parameters       :
   ' *                    sConnectionDBIn As String
   ' *                    sConnectionDBOut As String
   ' *                    sTable As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim oRecordIn        As ADODB.Recordset
   Dim oRecordOut       As ADODB.Recordset
   Dim sSQL             As String
   Dim nI               As Long
   Dim sName            As String

   If Len(sConnectionDBIn) > 0 Then
      If Len(sConnectionDBOut) > 0 Then
         sSQL = "Select * From [" & sTable & "]"
         Set oRecordIn = New ADODB.Recordset
         oRecordIn.Open sSQL, sConnectionDBIn, adOpenForwardOnly, adLockReadOnly
         Set oRecordOut = New ADODB.Recordset
         oRecordOut.Open sSQL, sConnectionDBOut, adOpenDynamic, adLockOptimistic
         With oRecordIn
            Do While Not .EOF
               oRecordOut.AddNew    ' *** New row
               For nI = 0 To .Fields.Count - 1  ' *** Copy the data
                  sName = .Fields(nI).Name
                  oRecordOut.Fields(sName).Value = .Fields(sName).Value
               Next
               oRecordOut.Update    ' *** Update the row
               .MoveNext
            Loop
         End With
         oRecordOut.Close
         Set oRecordOut = Nothing
         
         oRecordIn.Close
         Set oRecordIn = Nothing
         
         CopyData = True
         Exit Function
      End If
   End If

End Function

Public Function DuplicateField(sConnection As String, sTable As String, sFieldSrc As String, sFieldDest As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Waty Thierry
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : cUpdateADO
   ' * Procedure Name   : DuplicateField
   ' * Parameters       :
   ' *                    sConnection As String
   ' *                    sTable As String
   ' *                    sFieldSrc As String
   ' *                    sFieldDest As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          :
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim oConnection      As ADODB.Connection
   Dim sSQL             As String

   If Len(sConnection) > 0 Then
      sSQL = "Update [" & sTable & "] Set [" & sFieldDest & "] = [" & sFieldSrc & "]"
            
      Set oConnection = New ADODB.Parameter
      oConnection.Open sConnection
      oConnection.Execute sSQL
      oConnection.Close
      Set oConnection = Nothing
      
      DuplicateField = True
      Exit Function
   End If

End Function
