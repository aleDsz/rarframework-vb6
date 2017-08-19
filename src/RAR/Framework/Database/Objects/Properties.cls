VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Property"
Attribute VB_Ext_KEY = "Member0" ,"Property"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private List As Collection

Public Sub Add(ByVal Prop As Property)
    Call List.Add(Prop)
End Sub

Public Function Item(Index As Integer) As Property
    Set Item = List.Item(Item)
End Function

Public Property Get Count() As Long
    Count = List.Count
End Property

Public Function GetCollection() As Collection
    Set GetCollection = List
End Function

Public Function GetPrimaryKeys() As Properties

    Dim PKs  As New Properties
    Dim Prop As Property
    
    For Each Prop In List
    
        If Prop.PrimaryKey = True And Not Trim(Prop.Value) = "" Then
            Call PKs.Add(Prop)
        End If
        
    Next Prop
    
    Set GetPrimaryKeys = PKs
    
End Function

Public Function GetNonPrimaryKeys() As Properties

    Dim NonPKs As New Properties
    Dim Prop   As Property
    
    For Each Prop In List
        If Prop.PrimaryKey = False Then
            Call NonPKs.Add(Prop)
        End If
        
    Next Prop
    
    Set GetNonPrimaryKeys = NonPKs
    
End Function

Private Sub Class_Initialize()
    Set List = New Collection
End Sub

Private Sub Class_Terminate()
    Set List = Nothing
End Sub
