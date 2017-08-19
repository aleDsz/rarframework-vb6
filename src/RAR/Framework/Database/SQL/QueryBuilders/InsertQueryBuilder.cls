VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsertQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FieldList As New Collection
Private FromList As New Collection
Private ValueList As New Collection

Public Sub AddField(ByVal Field As String)
    FieldList.Add (Field)
End Sub

Public Sub AddFrom(ByVal From As String)
    FromList.Add (From)
End Sub

Public Sub AddValue(ByVal Value As String)
    ValueList.Add (Value)
End Sub

Public Function GetFieldClause() As String
    
    Dim Field As Variant
    
    Let GetFieldClause = Empty
    
    For Each Field In FieldList
        GetFieldClause = GetFieldClause & Field & ", "
    Next Field
    
    Let GetFieldClause = Trim(Mid(GetFieldClause, 1, Len(GetFieldClause) - 2))
    
End Function

Public Function GetFromClause() As String
    
    Dim From As Variant
    
    Let GetFromClause = Empty
    
    For Each From In FromList
        GetFromClause = GetFromClause & From & ", "
    Next From
    
    Let GetFromClause = Trim(Mid(GetFromClause, 1, Len(GetFromClause) - 2))
    
End Function

Public Function GetValueClause() As String
    
    Dim Value As Variant
    
    Let GetValueClause = Empty
    
    For Each Value In ValueList
        GetValueClause = GetValueClause & QueryBuilder.GetValue(Value) & ", "
    Next Value
    
    Let GetValueClause = Trim(Mid(GetValueClause, 1, Len(GetValueClause) - 2))
    
End Function

Public Function ToString() As String
    
    ToString = ""
    ToString = ToString & "INSERT INTO " & GetFromClause & vbCr & _
                          "(" & GetFieldClause & ")" & vbCr & _
                          " VALUES " & vbCr & _
                          "(" & GetValueClause & ")"
    
End Function


