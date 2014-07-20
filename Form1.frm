VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Text            =   "Results here..."
      Top             =   1440
      Width           =   7935
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "Execute"
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExecute_Click()
    txtResults.Text = "Button was Clicked"
End Sub
