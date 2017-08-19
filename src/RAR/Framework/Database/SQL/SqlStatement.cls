VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlStatement"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function GetQuotedValue(ByVal Prop As Property) As Variant
    
    On Error GoTo Error_GetQuotedValue
    
    If Prop.TypeName = "String" Then
        If InStr(1, Prop.Value, "%", vbTextCompare) Then
            Let GetQuotedValue = " LIKE '" & Prop.Value & "'"
        Else
            Let GetQuotedValue = " = '" & Prop.Value & "'"
        End If
    Else
        Let GetQuotedValue = " = " & Prop.Value
    End If
    
    Exit Function
    
Error_GetQuotedValue:
    Let GetQuotedValue = Empty
    Exit Function
    
End Function
