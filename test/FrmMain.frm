VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Testing Tools"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "Testar Select (All)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Testar Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Testar Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Testar Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Testar Select (ByKey)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Testar Conexão BD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DataContext     As DataContext
Private ObjectContext   As New ObjectContext
Private Usuario         As Usuario
Private Members         As TLI.Members
Private Member          As TLI.MemberInfo

Private Sub cmd_Click(Index As Integer)

    On Error GoTo Error_FrmMain

    Dim SqlStatementSelect As New SqlStatementSelect
    Dim SqlStatementInsert As New SqlStatementInsert
    Dim SqlStatementUpdate As New SqlStatementUpdate
    Dim SqlStatementDelete As New SqlStatementDelete
    Dim SSQL               As String
    Dim Snap               As ADODB.Recordset
    Dim ListObject         As Collection

    Select Case Index
        
        Case 0
            Set DataContext = DataFactory.GetInstanceOfDataContext
            Set Snap = New ADODB.Recordset
            Set Snap = DataContext.ExecuteReader("SELECT * FROM adm_estusu")
            
            Do While Not Snap.EOF
                Call MessageBox.Show("'" & Trim(Snap("usu_login")) & "'")
                Snap.MoveNext
            Loop
            
            Exit Sub
            
        Case 1
            Set Usuario = New Usuario
            Usuario.Login = "alexandre@aledsz.com.br"
            'Usuario.Estabelecimento = 1
            'Usuario.Nome = "Xandynho"
            'Usuario.TipoUsuario = 1
            'Usuario.Bloqueado = 0
            'Usuario.Senha = "123456"
            
            Let SSQL = SqlStatementSelect.GetSQL(False, Usuario)
            
            If SSQL <> "" Then
                Set DataContext = DataFactory.GetInstanceOfDataContext
                Set Snap = New ADODB.Recordset
                Set Snap = DataContext.ExecuteReader(SSQL)
                
                Set Usuario = ObjectContext.GetObject(Usuario, Snap)
                
                Call MessageBox.Show("CPF[1]: " & Usuario.CPF)
                
                Let SSQL = SqlStatementSelect.GetSQL(False, Usuario)
                Set Snap = New ADODB.Recordset
                Set Snap = DataContext.ExecuteReader(SSQL)
                
                Set Usuario = ObjectContext.GetObject(Usuario, Snap)
                
                Call MessageBox.Show("CPF[2]: " & Usuario.CPF)
            End If
            
            Exit Sub
            
        Case 2
            Set Usuario = New Usuario
            Usuario.Estabelecimento = 1
            Usuario.CPF = "43591017833"
            Usuario.RG = "54643101X"
            Usuario.Nome = "Xandynho"
            Usuario.Login = "alexandre@aledsz.com.br"
            Usuario.Senha = "123456"
            Usuario.TipoUsuario = 1
            Usuario.Bloqueado = 1
            
            Let SSQL = SqlStatementInsert.GetSQL(Usuario)
            
            If SSQL <> "" Then
                Set DataContext = DataFactory.GetInstanceOfDataContext
                Call DataContext.ExecuteQuery(SSQL)
                Call MessageBox.Show("Base de Dados Atualizada!")
            End If
            
        Case 3
            Set Usuario = New Usuario
            Usuario.Estabelecimento = 1
            Usuario.Nome = "Xandynho"
            Usuario.TipoUsuario = 1
            Usuario.Bloqueado = 0
            Usuario.Senha = "123456"
            
            Let SSQL = SqlStatementSelect.GetSQL(False, Usuario)
            
            If SSQL <> "" Then
                Set DataContext = DataFactory.GetInstanceOfDataContext
                Set Snap = New ADODB.Recordset
                Set Snap = DataContext.ExecuteReader(SSQL)
                
                If Not Snap.EOF Then
                    Usuario.Login = "alexandre@aledsz.com.br"
                    
                    Let SSQL = SqlStatementUpdate.GetSQL(Usuario)
                    Call DataContext.ExecuteQuery(SSQL)
                    Call MessageBox.Show("Base de Dados Atualizada!")
                End If
                
                Snap.Close
            End If
            
            Exit Sub
            
        Case 4
            Set Usuario = New Usuario
            Usuario.Estabelecimento = 1
            Usuario.Nome = "Xandynho"
            Usuario.TipoUsuario = 1
            Usuario.Bloqueado = 0
            Usuario.Senha = "123456"
            
            Let SSQL = SqlStatementSelect.GetSQL(False, Usuario)
            
            If SSQL <> "" Then
                Set DataContext = DataFactory.GetInstanceOfDataContext
                Set Snap = New ADODB.Recordset
                Set Snap = DataContext.ExecuteReader(SSQL)
                
                If Not Snap.EOF Then
                    Let SSQL = SqlStatementDelete.GetSQL(Usuario)
                    Call DataContext.ExecuteQuery(SSQL)
                    Call MessageBox.Show("Base de Dados Atualizada!")
                End If
                
                Snap.Close
            End If
            
            Exit Sub
            
        Case 5
            Set Usuario = New Usuario
            'Usuario.Login = "alexandre@aledsz.com.br"
            'Usuario.Estabelecimento = 1
            'Usuario.Nome = "Xandynho"
            'Usuario.TipoUsuario = 1
            'Usuario.Bloqueado = 0
            'Usuario.Senha = "123456"
            
            Let SSQL = SqlStatementSelect.GetSQL(True, Usuario)
            
            If SSQL <> "" Then
                Set DataContext = DataFactory.GetInstanceOfDataContext
                Set Snap = New ADODB.Recordset
                Set Snap = DataContext.ExecuteReader(SSQL)
                
                Set ListObject = ObjectContext.GetObjects(Usuario, Snap)
                
                For Each Usuario In ListObject
                    Call MessageBox.Show("Login: " & Usuario.Login)
                Next Usuario
                
                'Call MessageBox.Show("Quantidade de objetos obtidos: " & ListObject.Count)
            End If
            
            Exit Sub
            
    End Select
    
Error_FrmMain:
    Call MessageBox.Show(Error(Err), Me.Name, vbOKOnly, vbError)
    Exit Sub
    
End Sub
