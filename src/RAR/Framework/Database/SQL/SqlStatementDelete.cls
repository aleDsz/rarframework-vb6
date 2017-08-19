VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlStatementDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SSQL   As String
Private Object As Object

Private Sub CreateSQL()

    Dim DeleteQueryBuilder As New DeleteQueryBuilder
    Dim ObjectContext      As New ObjectContext
    Dim Props              As Properties
    Dim Prop               As Property
    
    Set Props = ObjectContext.GetProperties(Object).GetPrimaryKeys
    
    If Not Props Is Nothing Then
    
        If Props.Count = 0 Then
            Call MessageBox.Show("Não foram encontradas as primary keys", "SqlStatementDelete.CreateSQL", vbOKOnly, vbError)
            Let SSQL = Empty
            Exit Sub
        End If

        DeleteQueryBuilder.AddFrom (ObjectContext.GetTable(Object))

        For Each Prop In Props.GetCollection
            If Not Trim(Prop.Value) = "" Then
                DeleteQueryBuilder.AddWhere (Prop.FieldName & SqlStatement.GetQuotedValue(Prop))
            End If
        Next Prop
        
        Let SSQL = Trim(DeleteQueryBuilder.ToString)
    Else
        Let SSQL = Empty
        Call MessageBox.Show("Não foram encontradas as propriedades para: " & ObjectContext.GetObjectName(Object), "SqlStatementDelete.CreateSQL", vbOKOnly, vbError)
    End If
    
End Sub

Public Function GetSQL(ByVal Obj As Object) As String
    Set Object = Obj
    Call CreateSQL
    Let GetSQL = SSQL
End Function


