VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlStatementUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SSQL   As String
Private Object As Object

Private Sub CreateSQL()

    Dim UpdateQueryBuilder As New UpdateQueryBuilder
    Dim ObjectContext      As New ObjectContext
    Dim Props              As Properties
    Dim ListPks            As Properties
    Dim ListNonPks         As Properties
    Dim Prop               As Property
    
    Set Props = ObjectContext.GetProperties(Object)
    
    If Not Props Is Nothing Then
    
        If Props.GetPrimaryKeys.Count = 0 Then
            Call MessageBox.Show("Não foram encontradas as primary keys", "SqlStatementUpdate.CreateSQL", vbOKOnly, vbError)
            Let SSQL = Empty
            Exit Sub
        End If
        
        Set ListPks = Props.GetPrimaryKeys
        Set ListNonPks = Props.GetNonPrimaryKeys

        UpdateQueryBuilder.AddFrom (ObjectContext.GetTable(Object))
        
        For Each Prop In ListPks.GetCollection
            UpdateQueryBuilder.AddWhere (Prop.FieldName & SqlStatement.GetQuotedValue(Prop))
        Next Prop
        
        For Each Prop In ListNonPks.GetCollection
            UpdateQueryBuilder.AddField (Prop.FieldName)
            UpdateQueryBuilder.AddValue (Prop.Value)
        Next Prop
        
        Let SSQL = Trim(UpdateQueryBuilder.ToString)
    Else
        Let SSQL = Empty
        Call MessageBox.Show("Não foram encontradas as propriedades para: " & ObjectContext.GetObjectName(Object), "SqlStatementUpdate.CreateSQL", vbOKOnly, vbError)
    End If
    
End Sub

Public Function GetSQL(ByVal Obj As Object) As String
    Set Object = Obj
    Call CreateSQL
    Let GetSQL = SSQL
End Function

