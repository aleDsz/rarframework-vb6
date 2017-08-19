VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlStatementInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private SSQL   As String
Private Object As Object

Private Sub CreateSQL()

    Dim InsertQueryBuilder As New InsertQueryBuilder
    Dim ObjectContext      As New ObjectContext
    Dim Props              As Properties
    Dim Prop               As Property
    
    Set Props = ObjectContext.GetProperties(Object)
    
    If Not Props Is Nothing Then
    
        InsertQueryBuilder.AddFrom (ObjectContext.GetTable(Object))
        
        For Each Prop In Props.GetCollection
            InsertQueryBuilder.AddField (Prop.FieldName)
            InsertQueryBuilder.AddValue (Prop.Value)
        Next Prop
        
        Let SSQL = Trim(InsertQueryBuilder.ToString)
    Else
        Let SSQL = Empty
        Call MessageBox.Show("Não foram encontradas as propriedades para este objeto.", "SqlStatementInsert.CreateSQL", vbOKOnly, vbError)
    End If
    
End Sub

Public Function GetSQL(ByVal Obj As Object) As String
    Set Object = Obj
    Call CreateSQL
    Let GetSQL = SSQL
End Function
