VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataContextFactory"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Connection  As ADODB.Connection
Private XML         As New DOMDocument
Private PathManager As New PathManager
Private IsConnected As Boolean

Private Static Sub LoadConfig()
    
    Dim ConnectionString As IXMLDOMNode
    Dim Timeout          As IXMLDOMNode
    
    On Error GoTo Error_LoadConfig
    
    If Not IsConnected Then
        XML.async = True
    
        If XML.Load(PathManager.GetConfig) Then
            Set ConnectionString = XML.documentElement.selectSingleNode("connectionString")
            Set Timeout = XML.documentElement.selectSingleNode("timeout")
            
            Connection.ConnectionString = ConnectionString.Text
            Connection.CommandTimeout = Timeout.Text
        Else
            Call MessageBox.Show("Arquivo de configuração não encontrado" & vbCr & _
                                 "Arquivo: " & PathManager.GetConfig, "DataContextFactory.LoadConfig", vbOKOnly, vbError)
        End If
    End If
    
    Exit Sub
    
Error_LoadConfig:
    Call MessageBox.Show(Error(Err), "DataContextFactory.LoadConfig", vbOKOnly, vbError)
    Exit Sub
    
End Sub

Public Static Function GetConnection() As ADODB.Connection
    
    On Error GoTo Error_GetConnection
    
    Call LoadConfig
    
    If Connection.ConnectionString = "" Then
        Call MessageBox.Show("Não foi possível realizar a conexão com o banco de dados!", "DataContextFactory.GetConnection", vbOKOnly, vbError)
        Set GetConnection = Nothing
        Exit Function
    End If
    
    Connection.Open
    IsConnected = True
    
    Set GetConnection = Connection
    
    Exit Function
    
Error_GetConnection:
    Call MessageBox.Show(Error(Err), "DataContextFactory.GetConnection", vbOKOnly, vbError)
    Exit Function
    
End Function

Private Static Sub Class_Initialize()
    Set Connection = New ADODB.Connection
End Sub

