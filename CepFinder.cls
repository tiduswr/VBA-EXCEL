VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CepFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public cep As String
Public logradouro As String
Public complemento As String
Public bairro As String
Public localidade As String
Public uf As String
Public ibge As Long
Public gia As String
Public ddd As Integer
Public siafi As Long

Public Function CepFind(ByVal cep As String)
    
    ''Requer a Referencia Microsoft XML
    Dim DOC As DOMDocument
    Dim node As IXMLDOMNode
    Dim nodes As IXMLDOMNodeList
    
    On Error GoTo erro
    clear
    Set DOC = New DOMDocument
    DOC.async = False
    DOC.Load "https://viacep.com.br/ws/" & cep & "/xml/"
    
    Set nodes = DOC.SelectNodes("/xmlcep/*")
    
    For Each node In nodes
        Select Case node.nodeName
            Case "cep"
                cep = node.Text
            Case "logradouro"
                logradouro = node.Text
            Case "complemento"
                complemento = node.Text
            Case "bairro"
                bairro = node.Text
            Case "localidade"
                localidade = node.Text
            Case "uf"
                uf = node.Text
            Case "ibge"
                ibge = CLng(node.Text)
            Case "gia"
                gia = node.Text
            Case "ddd"
                ddd = CInt(node.Text)
            Case "siafi"
                siafi = CLng(node.Text)
        End Select
    Next node
    
    Set DOC = Nothing
    Set nodes = Nothing
    CepFind = True
    
    Exit Function
erro:
    Set DOC = Nothing
    Set nodes = Nothing
    CepFind = False
End Function
Private Function clear()
    cep = ""
    logradouro = ""
    complemento = ""
    bairro = ""
    localidade = ""
    uf = ""
    ibge = 0
    gia = ""
    ddd = 0
    siafi = 0
End Function


