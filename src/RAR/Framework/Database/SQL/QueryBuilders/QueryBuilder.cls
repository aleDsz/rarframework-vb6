VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryBuilder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Static Function GetValue(Value As Variant) As Variant

    If Trim(Value) = "" Then
        GetValue = "null"
    Else
        If Not IsNumeric(Value) Then
            Let GetValue = "'" & Value & "'"
        Else
            Let GetValue = Value
        End If
    End If
    
End Function
