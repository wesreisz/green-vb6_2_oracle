VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSQL 
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Text            =   "Select * from TEST_TABLE"
      Top             =   240
      Width           =   5055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5415
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "Execute"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblResults 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExecute_Click()
    lblResults.Caption = "executing... "
    'call db connect logic
    testDbConnection
    lblResults.Caption = lblResults.Caption + "completed... "
End Sub

Private Sub testDbConnection()
Dim intResult
Dim strDatabase
Dim strUserName
Dim strPassword
Dim strSQL
Dim dbDatabase
Dim snpData

'On Error Resume Next

strDatabase = "xxx" 'From tnsnames.ora
strUserName = "xxx"
strPassword = "xxx"

Set snpData = CreateObject("ADODB.Recordset")
Set dbDatabase = CreateObject("ADODB.Connection")

dbDatabase.ConnectionTimeout = 40
dbDatabase.ConnectionString = "Provider=OraOLEDB.Oracle;Data Source=" & strDatabase & ";User ID=" & strUserName & ";Password=" & strPassword & ";"
dbDatabase.Open

If (dbDatabase.State = 1) And (Err = 0) Then
    strSQL = "SELECT" & vbCrLf
    strSQL = strSQL & "  SYSDATE CURRENT_DATE" & vbCrLf
    strSQL = strSQL & "FROM" & vbCrLf
    strSQL = strSQL & "  DUAL"

    snpData.Open strSQL, dbDatabase

    If snpData.State = 1 Then
        If Not (snpData.EOF) Then
            Do While Not (snpData.EOF)
            lblResults.Caption = lblResults.Caption + CStr(snpData("current_date")) & vbCrLf
            
            snpData.MoveNext
            Loop
        Else
            lblResults.Caption = "No Results Found"
            
        End If
        snpData.Close
    End If
Else
    intResult = MsgBox("Could not connect to the database.  Check your user name and password." & vbCrLf & Error(Err), 16, "Oracle Connection Demo")
End If

dbDatabase.Close
Set snpData = Nothing
Set dbDatabase = Nothing
End Sub
