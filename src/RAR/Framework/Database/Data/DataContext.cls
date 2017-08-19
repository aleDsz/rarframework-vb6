VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private dbConnection       As ADODB.Connection
 Public IsConnected        As Boolean

Private Static Sub Class_Initialize()
    Set dbConnection = DataContextFactory.GetConnection
    
    If Not dbConnection Is Nothing Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Sub

Public Sub Begin()

    On Error GoTo Error_Begin

    If Not dbConnection = "" Then
        dbConnection.BeginTrans
        Exit Sub
    End If
    
Error_Begin:
    Call MessageBox.Show(Error(Err), "DataContext.Begin", vbOKOnly, vbError)
    Exit Sub

End Sub

Public Sub Commit()

    On Error GoTo Error_Commit

    If Not dbConnection = "" Then
        dbConnection.CommitTrans
        Exit Sub
    End If
    
Error_Commit:
    Call MessageBox.Show(Error(Err), "DataContext.Commit", vbOKOnly, vbError)
    Exit Sub

End Sub

Public Sub Rollback()

    On Error GoTo Error_Rollback

    If Not dbConnection = "" Then
        dbConnection.RollbackTrans
        Exit Sub
    End If
    
Error_Rollback:
    Call MessageBox.Show(Error(Err), "DataContext.Rollback", vbOKOnly, vbError)
    Exit Sub

End Sub

Public Sub DirtyRead()

    On Error GoTo Error_DirtyRead
    
    If Not dbConnection = "" Then
        dbConnection.IsolationLevel = adXactReadUncommitted
        Exit Sub
    End If
    
Error_DirtyRead:
    Call MessageBox.Show(Error(Err), "DataContext.DirtyRead", vbOKOnly, vbError)
    Exit Sub
    
End Sub

Public Sub CommitedRead()

    On Error GoTo Error_CommitedRead

    If Not dbConnection = "" Then
        dbConnection.IsolationLevel = adXactReadCommitted
        Exit Sub
    End If
    
Error_CommitedRead:
    Call MessageBox.Show(Error(Err), "DataContext.CommitedRead", vbOKOnly, vbError)
    Exit Sub
    
End Sub

Public Sub ExecuteQuery(ByVal sql As String)

    On Error GoTo Error_ExecuteQuery

    If Not dbConnection = "" Then
        dbConnection.Execute (sql)
        Exit Sub
    End If
    
Error_ExecuteQuery:
    Call MessageBox.Show(Error(Err), "DataContext.ExecuteQuery", vbOKOnly, vbError)
    Exit Sub
    
End Sub

Public Function ExecuteReader(ByVal sql As String) As ADODB.Recordset

    Dim Snap As New ADODB.Recordset
    
    On Error GoTo Error_ExecuteReader

    If Not dbConnection = "" Then
        Call Snap.Open(sql, dbConnection, adOpenUnspecified, adLockReadOnly)
        Set ExecuteReader = Snap
        Exit Function
    End If
    
Error_ExecuteReader:
    Call MessageBox.Show(Error(Err), "DataContext.ExecuteReader", vbOKOnly, vbError)
    Exit Function
    
End Function
