# RAR Framework in VB6

## 1. Introdução

Após ter criado o mesmo framework, originalmente em [PHP](https://github.com/aleDsz/rarframework), percebi que eu teria a mesma necessidade de um ORM em outras linguagens. Assim como eu precisei quando comecei a utilizar o VB6 em ambiente profissional e, com a praticidade que eu tinha em PHP, resolvi adaptar para VB6.

## 2. Como Funciona

Através do pacote DBI, é possível realizar uma conexão com vários tipos de banco de dados. Além disso, por meio do `Generics`, é possível acessar o conteúdo de um objeto e obter todas as informações necessárias para criar uma instrução SQL.

Neste caso, uma classe deve seguir o seguinte modelo:

```vb
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
Attribute VB_Description = "nome_da_tabela" ' Nome da tabela
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private pCampo   As String

Public Property Get Campo() As Variant
Attribute Campo.VB_Description = "nome_do_campo, [True/False]" 'True quando for uma Primary Key, False quando não for
    Let Campo = pCampo
End Property

Public Property Let Campo(sCampo As String)
    Let pCampo = sCampo
End Property
```

## 3. Como Utilizar

Para que você possa utilizar todos as funcionalidades do framework no seu ambiente, você pode criar 1 (ou mais, dependendo da sua forma de trabalho) classe para acessar ao banco de dados de forma genérica.

```vb
' Estou sem o código ainda
```

**OBS.:** Você não precisa criar a classe de forma genérica, você pode criar uma classe de acesso a dados para cada entidade que você criar no modelo citado acima.

## 4. Como Contribuir

Para contribuir, você pode realizar um **fork** do nosso repositório e nos enviar um Pull Request.

## 5. Doação

Caso queria fazer uma doação para o projeto, você pode realizar [aqui](https://twitch.streamlabs.com/aleDsz)

## 6. Suporte

Caso você tenha algum problema ou uma sugestão, você pode nos contatar [aqui](https://github.com/aleDsz/rarframework-net/issues).

## 7. Licença

Cheque [aqui](LICENSE)
