VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataFactory"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private MessageBox            As New MessageBox
Private InstanceOfDataContext As DataContext

Public Static Function GetInstanceOfDataContext() As DataContext
    
    On Error GoTo Error_GetInstanceOfDataContext
    
    If InstanceOfDataContext Is Nothing Then
        Set InstanceOfDataContext = New DataContext
    End If
    
    Set GetInstanceOfDataContext = InstanceOfDataContext
    
    Exit Function
    
Error_GetInstanceOfDataContext:
    Call MessageBox.Show(Error(Err), "DataFactory.GetInstanceOfDataContext", vbOKOnly, vbError)
    Exit Function
    
End Function

