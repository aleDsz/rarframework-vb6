VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private InterfaceInfo As TLI.InterfaceInfo
Private TLI           As TLI.TLIApplication

Public Function GetProperties(ByVal Object As Object) As Properties
    
    On Error GoTo Error_GetProperties

    Dim MemberInfo As TLI.MemberInfo
    Dim Props      As New Properties
    Dim Prop       As Property
    Dim I          As Integer
    
    Set TLI = New TLI.TLIApplication
    Set InterfaceInfo = TLI.InterfaceInfoFromObject(Object)
    
    For I = 1 To InterfaceInfo.Members.Count Step 2
        Set MemberInfo = InterfaceInfo.Members(I)
        
        Set Prop = New Property
        Set Prop = GetProperty(Object, MemberInfo.Name)
        
        Props.Add Prop
    Next I
    
    Set GetProperties = Props
    
    Exit Function
    
Error_GetProperties:
    Call MessageBox.Show(Error(Err), "ObjectContext.GetProperties", vbOKOnly, vbError)
    Set GetProperties = Nothing
    Exit Function
    
End Function

Public Function GetProperty(ByVal Object As Object, ByVal PropName As String) As Property
    
    On Error GoTo Error_GetProperty
    
    Dim MemberInfo   As TLI.MemberInfo
    Dim Prop         As Property
    Dim Attributes() As String
    
    Set InterfaceInfo = TLI.InterfaceInfoFromObject(Object)

    For Each MemberInfo In InterfaceInfo.Members
    
        If MemberInfo.Name = PropName Then
        
            Set Prop = New Property
            
            Let Attributes = Split(MemberInfo.HelpString, ", ")
            
            Prop.FieldName = Attributes(0)
            Prop.PrimaryKey = CBool(Attributes(1))
            Prop.PropName = MemberInfo.Name
            Prop.TypeName = TypeName(Prop.PropName)
            Prop.Value = GetPropValue(Object, MemberInfo)
            
            Set GetProperty = Prop
            
            Exit Function
            
        End If
        
    Next MemberInfo
    
    Exit Function
    
Error_GetProperty:
    Set GetProperty = Nothing
    Exit Function
    
End Function

Public Function GetPropValue(ByVal Object As Object, ByVal Prop As TLI.MemberInfo) As Variant

    On Error GoTo Error_GetPropValue
    
    Let GetPropValue = TLI.InvokeHook(Object, Prop.Name, INVOKE_PROPERTYGET)
    
    Exit Function
    
Error_GetPropValue:
    Let GetPropValue = Null
    Exit Function

End Function

Public Function GetTable(ByVal Object As Object) As Variant

    On Error GoTo Error_GetTable
    
    Let GetTable = InterfaceInfo.HelpString
    
    Exit Function
    
Error_GetTable:
    Let GetTable = Empty
    Exit Function

End Function

Public Function GetObjectName(ByVal Object As Object) As Variant

    On Error GoTo Error_GetObjectName
    
    Let GetObjectName = InterfaceInfo.Name
    
    Exit Function
    
Error_GetObjectName:
    Let GetObjectName = Empty
    Exit Function

End Function

Public Function GetObject(ByVal Obj As Object, ByVal DataReader As ADODB.Recordset) As Object
    
    Dim Props  As Properties
    Dim Prop   As Property
    Dim NewObj As Object
    
    On Error GoTo Error_GetObject
    
    Set Props = GetProperties(Obj)
    Set NewObj = CopyObject(Obj)
    
    If DataReader.EOF Then GoTo Error_GetObject

    Set Props = GetProperties(NewObj)
    
    For Each Prop In Props.GetCollection
        Call TLI.InvokeHookSub(NewObj, Prop.PropName, INVOKE_PROPERTYPUT, DataReader(Prop.FieldName).Value)
    Next Prop
    
    DataReader.Close
    
    Set GetObject = Obj
    Exit Function
    
Error_GetObject:
    Set GetObject = Nothing
    Exit Function
    
End Function
    
Public Function GetObjects(ByVal Obj As Object, ByVal DataReader As ADODB.Recordset) As Collection
    
    Dim Props    As Properties
    Dim Prop     As Property
    Dim Objects  As Collection
    Dim NewObj   As Object
    
    On Error GoTo Error_GetObjects
    
    Set Props = GetProperties(Obj)
    
    If DataReader.EOF Then GoTo Error_GetObjects
    
    Set Objects = New Collection
    
    Do While Not DataReader.EOF
    
        Set NewObj = CopyObject(Obj)
    
        For Each Prop In Props.GetCollection
            Call TLI.InvokeHookSub(NewObj, Prop.PropName, INVOKE_PROPERTYPUT, DataReader(Prop.FieldName).Value)
        Next Prop
        
        Call Objects.Add(NewObj, DataReader(Props.GetCollection.Item(2).FieldName).Value)
        
        DataReader.MoveNext
    Loop
    
    DataReader.Close
    
    Set GetObjects = Objects
    Exit Function
    
Error_GetObjects:
    Set GetObjects = Nothing
    Exit Function
    
End Function

Public Function CopyObject(ByVal Obj As Object) As Object

    Dim Props   As Properties
    Dim Prop    As Property
    Dim NewObj  As Object
    
    On Error GoTo Error_CopyObject
    
    Set Props = GetProperties(Obj)
    Set NewObj = Obj
    
    For Each Prop In Props.GetCollection
        Call TLI.InvokeHookSub(NewObj, Prop.PropName, INVOKE_PROPERTYPUT, Empty)
    Next Prop
    
    Set CopyObject = NewObj
    Exit Function
    
Error_CopyObject:
    Set CopyObject = Nothing
    Exit Function

End Function
