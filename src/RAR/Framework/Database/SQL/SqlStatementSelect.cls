VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlStatementSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private SSQL   As String
Private Object As Object

Private Sub CreateSQL(ByVal IsList)

    Dim SelectQueryBuilder As New SelectQueryBuilder
    Dim ObjectContext      As New ObjectContext
    Dim Props              As Properties
    Dim Prop               As Property
    
    Set Props = ObjectContext.GetProperties(Object)
    
    If Not Props Is Nothing Then
    
        If Not IsList Then

            If Props.GetPrimaryKeys.Count = 0 Then
                Call MessageBox.Show("Não foram encontradas as primary keys", "SqlStatementSelect.CreateSQL", vbOKOnly, vbError)
                Let SSQL = Empty
                Exit Sub
            End If

        End If
    
        SelectQueryBuilder.AddFrom (ObjectContext.GetTable(Object))
        
        For Each Prop In Props.GetCollection

            SelectQueryBuilder.AddField (Prop.FieldName)
            
            If Not Trim(Prop.Value) = "" Then
                SelectQueryBuilder.AddWhere (Prop.FieldName & SqlStatement.GetQuotedValue(Prop))
            End If

        Next Prop
        
        Let SSQL = Trim(SelectQueryBuilder.ToString)
    Else
        Let SSQL = Empty
        Call MessageBox.Show("Não foram encontradas as propriedades para: " & ObjectContext.GetObjectName(Object), "SqlStatementSelect.CreateSQL", vbOKOnly, vbError)
    End If
    
End Sub

Public Function GetSQL(ByVal IsList As Boolean, ByVal Obj As Object) As String
    Set Object = Obj
    Call CreateSQL(IsList)
    Let GetSQL = SSQL
End Function
