Attribute VB_Name = "Module1"
Option Explicit
'Result Variable for frmresult page
 Public result As String
'Voterlist variable for frmvoterlist page
 Public voterlist As String
'nomineelist variable for frmnominelist page
 Public nomineelist As String

'variables for storing Details of logged_in user and it's type
    Public usertype As String
    Public logged_userid As String
    Public logged_username As String

Dim rs As ADODB.Recordset
Dim conn As New ADODB.Connection


