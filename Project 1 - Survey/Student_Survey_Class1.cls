VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Student_Survey_Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub class_initialize()
    DoEvents
    Set con = New ADODB.Connection
    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" & App.Path & "\StudentSurveyDB.mdb;"
        .CursorLocation = adUseClient
        .Open
    End With
End Sub
