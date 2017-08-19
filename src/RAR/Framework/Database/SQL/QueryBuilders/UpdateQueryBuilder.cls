VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpdateQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FieldList As New Collection
Private FromList As New Collection
Private ValueList As New Collection
Private WhereList As New Collection

Public Sub AddFrom(ByVal From As String)
    FromList.Add (From)
End Sub

Public Sub AddValue(ByVal Value As String)
    ValueList.Add (Value)
End Sub

Public Sub AddField(ByVal Field As String)
    FieldList.Add (Field)
End Sub

Public Sub AddWhere(ByVal Where As String)
    WhereList.Add (Where)
End Sub

Public Function GetFromClause() As String
    
    Dim From As Variant
    
    Let GetFromClause = Empty
    
    For Each From In FromList
        GetFromClause = GetFromClause & From & ", "
    Next From
    
    Let GetFromClause = Trim(Mid(GetFromClause, 1, Len(GetFromClause) - 2))
    
End Function

Public Function GetSetClause() As String
    
    Dim Value As Variant
    Dim I     As Integer
    
    Let GetSetClause = Empty
    
    For I = 1 To FieldList.Count
        GetSetClause = GetSetClause & FieldList.Item(I) & " = " & _
                       QueryBuilder.GetValue(ValueList.Item(I)) & "," & vbCr & "       "
    Next I
    
    Let GetSetClause = Trim(GetSetClause)
    Let GetSetClause = Trim(Mid(GetSetClause, 1, Len(GetSetClause) - 2))
    
End Function

Public Function GetWhereClause() As String
    
    Dim Where As Variant
    
    Let GetWhereClause = Empty
    
    For Each Where In WhereList
        GetWhereClause = GetWhereClause & Where & vbCr & "   AND "
    Next Where
    
    Let GetWhereClause = Trim(Mid(GetWhereClause, 1, Len(GetWhereClause) - 7))
    
End Function

Public Function ToString() As String
    
    ToString = ""
    ToString = ToString & "UPDATE " & GetFromClause & vbCr & _
                          "   SET " & GetSetClause & vbCr & _
                          " WHERE " & GetWhereClause
    
End Function

