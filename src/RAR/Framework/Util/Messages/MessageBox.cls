VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub Show(ByVal Msg As String, Optional Title As String, Optional Style As VbMsgBoxStyle, Optional Icon As VbMsgBoxStyle)
    
    If Style = Null Then
        Style = vbOKOnly
    End If
    
    If Icon = Null Then
        Icon = vbInformation
    End If
    
    If Title = "" Then
        Title = App.EXEName
    End If
    
    Call MsgBox(Msg, Style, Title, 0, Icon)
    Exit Sub
    
End Sub

