VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private FieldList As New Collection
Private FromList As New Collection
Private WhereList As New Collection

Public Sub AddField(ByVal Field As String)
    FieldList.Add (Field)
End Sub

Public Sub AddFrom(ByVal From As String)
    FromList.Add (From)
End Sub

Public Sub AddWhere(ByVal Where As String)
    WhereList.Add (Where)
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

Public Function GetWhereClause() As String
    
    Dim Where As Variant
    
    Let GetWhereClause = Empty
    
    For Each Where In WhereList
        GetWhereClause = GetWhereClause & Where & vbCr & "   AND "
    Next Where
    
    Let GetWhereClause = Trim(Mid(GetWhereClause, 1, Len(GetWhereClause) - 7))
    
End Function

Public Function ToString() As String
    
    ToString = "SELECT " & GetFieldClause & vbCr & _
               "  FROM " & GetFromClause & vbCr
    
    If WhereList.Count > 0 Then
        ToString = ToString & " WHERE " & GetWhereClause
    End If
    
End Function

