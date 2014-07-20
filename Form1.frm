VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1080
      Width           =   8175
   End
   Begin VB.TextBox txtSQL 
      Height          =   855
      Left            =   120
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "SELECT SYSDATE CURRENT_DATE FROM DUAL"
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "Execute"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const g_SEP As String = "--------------------------------------------" + vbCrLf

Private Sub btnExecute_Click()
    txtResults.Text = "executing: " + txtSQL.Text + vbCrLf
    'call db connect logic
    If (Len(txtSQL.Text) > 0) Then
        testDbConnection (txtSQL.Text)
    Else
        testDbConnection ("SELECT SYSDATE CURRENT_DATE FROM DUAL")
    End If
    txtResults.Text = txtResults.Text + vbCrLf + _
        "completed... " + vbCrLf + vbCrLf
End Sub

Private Sub testDbConnection(ByVal strSQL)
Dim intResult
Dim strDatabase
Dim strUserName
Dim strPassword
Dim dbDatabase

'On Error Resume Next

strDatabase = "xxx" 'From tnsnames.ora
strUserName = "xxx"
strPassword = "xxx"

Set mRS = CreateObject("ADODB.Recordset")
Set dbDatabase = CreateObject("ADODB.Connection")

dbDatabase.ConnectionTimeout = 40
dbDatabase.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=" & strDatabase & ";User ID=" & strUserName & ";Password=" & strPassword & ";"
dbDatabase.Open

If (dbDatabase.State = 1) And (Err = 0) Then
    mRS.Open strSQL, dbDatabase
    
    If mRS.State = 1 Then
        If Not (mRS.EOF) Then
           Do While Not mRS.EOF
                For i = 0 To (mRS.Fields.Count - 1)
                    txtResults.Text = txtResults.Text & mRS(i) & " "
                Next
                txtResults.Text = txtResults.Text & vbCrLf
                mRS.MoveNext
            Loop
        Else
            txtResults.Text = "No Results Found"
        End If
        mRS.Close
    End If
Else
    intResult = MsgBox("Could not connect to the database.  Check your user name and password." & vbCrLf & Error(Err), 16, "Oracle Connection Demo")
End If

dbDatabase.Close
Set mRS = Nothing
Set dbDatabase = Nothing
End Sub
