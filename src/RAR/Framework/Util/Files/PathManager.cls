VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathManager"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Static Function GetCurrentPath() As String
    GetCurrentPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "")
End Function

Public Static Function GetConfig() As String
    GetConfig = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & App.EXEName & ".xml"
End Function

