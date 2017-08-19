VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "adm_estusu"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private pEst     As Variant
Private pLogin   As String
Private pNome    As String
Private pSenha   As String
Private pTpUsu   As Variant
Private pRG      As String
Private pCPF     As String
Private pBloq    As Variant

Public Property Get Estabelecimento() As Variant
Attribute Estabelecimento.VB_Description = "usu_est, True"
    Let Estabelecimento = pEst
End Property

Public Property Let Estabelecimento(sEst As Variant)
    Let pEst = sEst
End Property

Public Property Get Login() As String
Attribute Login.VB_Description = "usu_login, True"
    Let Login = pLogin
End Property

Public Property Let Login(sLogin As String)
    Let pLogin = sLogin
End Property

Public Property Get Nome() As String
Attribute Nome.VB_Description = "usu_nome, False"
    Let Nome = pNome
End Property

Public Property Let Nome(sNome As String)
    Let pNome = sNome
End Property

Public Property Get Senha() As String
Attribute Senha.VB_Description = "usu_senha, False"
    Let Senha = pSenha
End Property

Public Property Let Senha(sSenha As String)
    Let pSenha = sSenha
End Property

Public Property Get TipoUsuario() As Variant
Attribute TipoUsuario.VB_Description = "usu_tpusu, False"
    Let TipoUsuario = pTpUsu
End Property

Public Property Let TipoUsuario(sTpUsu As Variant)
    Let pTpUsu = sTpUsu
End Property

Public Property Get RG() As String
Attribute RG.VB_Description = "usu_rg, True"
    Let RG = pRG
End Property

Public Property Let RG(sRG As String)
    Let pRG = sRG
End Property

Public Property Get CPF() As String
Attribute CPF.VB_Description = "usu_cpf, True"
    Let CPF = pCPF
End Property

Public Property Let CPF(sCPF As String)
    Let pCPF = sCPF
End Property

Public Property Get Bloqueado() As Variant
Attribute Bloqueado.VB_Description = "usu_bloq, False"
    Let Bloqueado = pBloq
End Property

Public Property Let Bloqueado(sBloq As Variant)
    Let pBloq = sBloq
End Property

