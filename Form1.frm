VERSION 5.00
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
   Begin VB.CommandButton btnListAll 
      Caption         =   "List All"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   3720
      Width           =   7575
   End
   Begin VB.CommandButton btnStoredProcExecute 
      Caption         =   "Execute"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtCreatedBy 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Text            =   "wes"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Text            =   "sam"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtResults 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4320
      Width           =   8175
   End
   Begin VB.TextBox txtSQL 
      Height          =   855
      Left            =   240
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "SELECT SYSDATE CURRENT_DATE FROM DUAL"
      Top             =   600
      Width           =   5895
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "Execute"
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Stored Proc List Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label5 
      Caption         =   "Stored Proc Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "SQL Statement Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Created By"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SEP As String = "--------------------------------------------" + vbCrLf
Const DATABASE_NAME As String = "green" 'From tnsnames.ora
Const username As String = "green"
Const PASSWORD As String = "passwd"
Const conn As String = "Provider=OraOLEDB.Oracle;Data Source=" & DATABASE_NAME & ";User ID=" & username & ";Password=" & PASSWORD & ";"

Private Sub btnExecute_Click()
    Dim result As String
    MsgBox (conn)
    txtResults.Text = "executing: " + txtSQL.Text + vbCrLf
    'call db connect logic
    Dim x As New DBUtil
    If (Len(txtSQL.Text) > 0) Then
        result = x.testDbConnection(conn, txtSQL.Text)
    Else
        'if nothing was entered, just get the time from the oracle system tables.
        result = x.testDbConnection(conn, "SELECT SYSDATE CURRENT_DATE FROM DUAL")
    End If
    txtResults.Text = txtResults.Text + result + vbCrLf + _
        "completed... " + vbCrLf + vbCrLf
End Sub

Private Sub btnListAll_Click()
    Dim mRS As ADODB.Recordset
    Dim errMsg As String
    Dim db As New DBUtil
    Set mRS = db.listUsers(conn, errMsg)
    
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
    txtResults.Text = results
End Sub

Private Sub btnStoredProcExecute_Click()
    Dim result As String
    Dim username As String: username = txtUserName.Text
    Dim createdBy As String: createdBy = txtCreatedBy.Text
    If ((Len(username) <= 0) Or (Len(createdBy) <= 0)) Then
        Return
    End If
    
    Dim errMsg As String
    Dim db As New DBUtil
    result = db.insertUser(conn, username, createdBy, errMsg)
    If (result = "success") Then
        txtResults.Text = "Record added"
    Else
        txtResults.Text = errMsg
    End If
    
End Sub
