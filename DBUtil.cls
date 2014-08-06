VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function testDbConnection(ByVal conn As String, ByVal strSQL As String) As String
'---Usage:  Function used to demostrate simple sql calls against oracle 11g
'---Author: Wesley T. Reisz (wes@wesleyreisz.com)
'---ModDate: July 22, 2014
'---Retval: String
    Dim intResult As Integer
    Dim strDatabase As String
    Dim strUserName As String
    Dim strPassword As String
    Dim results As String
    
    'On Error Resume Next
    Dim mRS As ADODB.Recordset: Set mRS = New ADODB.Recordset
    Dim dbConn As ADODB.Connection: Set dbConn = New ADODB.Connection
    
    dbConn.ConnectionTimeout = 40
    dbConn.ConnectionString = conn
    dbConn.Open
    
    If (dbConn.State = 1) And (Err = 0) Then
        mRS.Open strSQL, dbConn
        
        If mRS.State = 1 Then
            If Not (mRS.EOF) Then
               Do While Not mRS.EOF
                    For i = 0 To (mRS.Fields.Count - 1)
                        results = results & mRS(i) & " "
                    Next
                    results = results & vbCrLf
                    mRS.MoveNext
                Loop
            Else
                results = "No Results Found"
            End If
            mRS.Close
        End If
        testDbConnection = results
    Else
        intResult = MsgBox("Could not connect to the database.  Check your user name and password." & vbCrLf & Error(Err), 16, "Oracle Connection Demo")
    End If
    
    dbConn.Close
    Set mRS = Nothing
    Set dbConn = Nothing
End Function

Public Function insertUser(ByVal conn As String, _
                           ByVal mUserName As String, _
                           ByVal mCreatedBy As String, _
                           ByRef errMsg As String) As String
'---Usage:  Function used demonstrate adding a record to an oracle db via a stored proc
'---Author: Wesley T. Reisz (wes@wesleyreisz.com)
'---ModDate: July 22, 2014
'---Retval: nothing
On Error GoTo errList
    Dim dbConn As ADODB.Connection: Set dbConn = New ADODB.Connection
    Dim dbCommand As ADODB.command: Set dbCommand = New ADODB.command
   
    dbConn.ConnectionTimeout = 40
    dbConn.Open (conn)
    
    dbConn.CursorLocation = adUseClient
        
    Dim prUserName As New ADODB.Parameter
    Dim prCreatedBy As New ADODB.Parameter
    
    Dim sql As String
    
 '=========================================
 '==Set Properties & Populate Recordset====
 '=========================================
    '---Set Command Properties
    sqlQuestions = "call insertdbuser(?,?)"

    Set dbCommand.ActiveConnection = dbConn
    dbCommand.CommandText = sqlQuestions
    dbCommand.CommandType = adCmdText
    
    Set prUserName = dbCommand.CreateParameter("p_USERNAME", adVarChar, adParamInput, 20)
    prUserName.Value = mUserName
            
    Set prCreatedBy = dbCommand.CreateParameter("p_CREATEDBY", adVarChar, adParamInput, 20)
    prCreatedBy.Value = mCreatedBy
    
    dbCommand.Parameters.Append prUserName
    dbCommand.Parameters.Append prCreatedBy
    dbCommand.Execute
    
    If IsObject(dbConn) Then
        dbConn.Close
    End If
    If IsObject(dbCommand) Then
        Set dbCommand = Nothing
    End If
    
    insertUser = "success"
    Exit Function
errList:
    errMsg = Err.Description & " mUserId: " & mUserId
    Err.Raise 651, "Error Insert", errMsg
End Function

Public Function listUsers(ByVal conn As String, _
                          ByRef errMsg As String) As ADODB.Recordset
'---Usage:  ListUsers is used to return a list of users of the system
'---Author: Wesley T. Reisz (wes@wesleyreisz.com)
'---ModDate: July 28, 2014
'---Retval: Recordset
On Error GoTo errListUsers
    'open and configure connection
    Dim dbConn As ADODB.Connection: Set dbConn = New ADODB.Connection
    Dim dbCommand As ADODB.command: Set dbCommand = New ADODB.command
    Dim objRS As ADODB.Recordset: Set objRS = New ADODB.Recordset
    
    dbConn.ConnectionTimeout = 40
    dbConn.Open (conn)
    dbConn.CursorLocation = adUseClient
    Set dbCommand.ActiveConnection = dbConn
    
    objRS.CursorType = adOpenStatic
    objRS.LockType = adLockReadOnly
    
    Dim sqlQuestions As String
      
    ' Ref cursor parameters are automatically
    ' bound by the provider.
    sqlQuestions = "{ CALL GET_DBUSERS() }"
    
    dbCommand.CommandText = sqlQuestions
    dbCommand.CommandType = adCmdText
    ' Enable PLSQLRSet property
    dbCommand.Properties("PLSQLRSet") = True
    
    Set objRS = dbCommand.Execute
    
    'cleanup
    Set dbCommand.ActiveConnection = Nothing
    Set objRS.ActiveConnection = Nothing
    Set listUsers = objRS.Clone
    
    If IsObject(dbConn) Then
        dbConn.Close
    End If
    
    If IsObject(dbCommand) Then
        Set dbCommand = Nothing
    End If
    
    If IsObject(objRS) Then
        If objRS.State = 1 Then
            objRS.Close
        End If
        Set objRS = Nothing
    End If
    Exit Function
errListUsers:
    errMsg = Err.Description
    Err.Raise 651, "GET_DBUSERS", "Failed retrieving open cases from database. " & errMsg
End Function


Private Sub configureConnection(ByRef dbConn As ADODB.Connection, ByRef command As ADODB.command)
    Set dbConn = New ADODB.Connection
End Sub


