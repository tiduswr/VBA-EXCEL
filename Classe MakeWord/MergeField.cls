VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MergeField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Desenvolvido por Harllem Nascimento
'Em 30 de julho de 2021 foi concluida a versão final do projeto
'contato: harllemnascimento@gmail.com

'Features:
'Em 25 de Novembro de 2021 foi inserida a opção de argumento para verificar se a tabela tem Cabeçalho

'Atributos
Private Type dimensions
    altura As Integer
    largura As Integer
End Type
Private Type intervalo
    de As Integer
    ate As Integer
End Type
Private Type tableProperties
    source As String
    colunaCriterio As Integer
    color As String
    intervalo As intervalo
    tiposColunas() As String
    hasHeader As Boolean
End Type
Private Type properties
    nome As String
    tipo As String
    arrgumento As String
    dimensoes As dimensions
    tabela As tableProperties
End Type

'Variaveis
Dim propriedades As properties
Private fieldTag As String

'Construtor
Public Function construtor(ByVal fieldExpression As String) As Boolean
    On Error GoTo erro
        Dim auxStr As String
        propriedades.nome = returnStringDelimited("", "{", fieldExpression)
        
        If InStr(1, fieldExpression, "(", 0) > 0 And InStr(1, fieldExpression, "{", 0) > 0 Then
            Dim x() As String
            
            x = Split(fieldExpression, "(")
            
            x(0) = Replace(x(0), " ", "")
            
            fieldExpression = x(0) & "(" & x(1)
            
        Else
            fieldExpression = Replace(fieldExpression, " ", "")
        End If
        
        fieldTag = fieldExpression
        
        If isImagem() Then
            propriedades.tipo = returnStringDelimited("", "(", returnStringDelimited("{", "}", fieldExpression))
            auxStr = returnStringDelimited("(", ")", fieldExpression)
            
            If auxStr <> "null" Then
                Dim auxArr() As String: auxArr = Split(auxStr, "x")
            
                propriedades.dimensoes.altura = CInt(auxArr(0))
                propriedades.dimensoes.largura = CInt(auxArr(1))
                propriedades.arrgumento = "imagem"
            Else
                propriedades.dimensoes.altura = 0
                propriedades.dimensoes.largura = 0
                propriedades.arrgumento = "imagem"
            End If
            
            
        Else
            If isTable() Then
                ''nome_campo{table(dados;1;color;1-5;[texto, data, inteiro, decimal])}
                propriedades.tipo = returnStringDelimited("", "(", returnStringDelimited("{", "}", fieldExpression))
                auxStr = returnStringDelimited("(", ")", fieldExpression)
                auxArr = Split(auxStr, ";")
                
                If UBound(auxArr, 1) <> 4 Then
                    MsgBox "Informe todos os argumentos da tabela na coluna " & propriedades.nome & "!", _
                            vbCritical, "Erro!"
                    construtor = False
                    Exit Function
                End If
                
                propriedades.tabela.source = auxArr(0)
                propriedades.tabela.colunaCriterio = CInt(auxArr(1))
                propriedades.tabela.color = auxArr(2)
                If auxArr(2) = "sem_cabecalho" Then
                    propriedades.tabela.hasHeader = False
                Else
                    propriedades.tabela.hasHeader = True
                End If
                
                If InStr(1, auxArr(3), "-", 0) <> 0 Then
                    Dim auxIntervalo() As String: auxIntervalo = Split(auxArr(3), "-")
                    propriedades.tabela.intervalo.de = CInt(auxIntervalo(0))
                    propriedades.tabela.intervalo.ate = CInt(auxIntervalo(1))
                    
                    With propriedades.tabela.intervalo
                        If .de > .ate Then
                            MsgBox "A coluna inicial da tabela fonta da coluna " & propriedades.nome & " não pode ser maior que a coluna final!", _
                                    vbCritical, "Erro!"
                            construtor = False
                            Exit Function
                        End If
                    End With
                Else
                    If countColumns(propriedades.tabela.source) >= 1 Then
                        propriedades.tabela.intervalo.de = 1
                        propriedades.tabela.intervalo.ate = countColumns(propriedades.tabela.source)
                    Else
                        MsgBox "Não existem colunas na tabela fonte especificada na coluna " & propriedades.nome & "!", _
                                    vbCritical, "Erro!"
                    End If
                End If
                
                auxArr(4) = Replace(auxArr(4), "[", "")
                auxArr(4) = Replace(auxArr(4), "]", "")
                propriedades.tabela.tiposColunas = Split(auxArr(4), ",")
                propriedades.arrgumento = "tabela"
                
                If CInt(UBound(propriedades.tabela.tiposColunas, 1) + 1) <> countColumns(propriedades.tabela.source) Then
                    MsgBox "Erro, o total de colunas tem de ser igual ao total de tipos de dados especificadas na coluna " _
                                & propriedades.nome & "!", vbCritical, "Erro!"
                    construtor = False
                    Exit Function
                End If
                
            Else
                If InStr(1, returnStringDelimited("{", "}", fieldExpression), "(", 0) <> 0 Then
                    propriedades.tipo = returnStringDelimited("", "(", returnStringDelimited("{", "}", fieldExpression))
                    propriedades.arrgumento = returnStringDelimited("(", ")", fieldExpression)
                Else
                    propriedades.tipo = returnStringDelimited("{", "}", fieldExpression)
                    propriedades.arrgumento = "null"
                End If
            End If
            
            If propriedades.tipo <> "local" And propriedades.tipo <> "arquivo" Then
                propriedades.arrgumento = Replace(propriedades.arrgumento, " ", "")
            End If
            
        End If
        construtor = True
    Exit Function
erro:
    construtor = False
End Function

''Getters
Public Function getNome() As String
    getNome = propriedades.nome
End Function
Public Function getTipo() As String
    getTipo = propriedades.tipo
End Function
Public Function getImagemAltura() As Integer
    getImagemAltura = propriedades.dimensoes.altura
End Function
Public Function getImagemLargura() As Integer
    getImagemLargura = propriedades.dimensoes.largura
End Function
Public Function getTableSource() As String
    getTableSource = propriedades.tabela.source
End Function
Public Function getTableColunaCriterio() As Integer
    getTableColunaCriterio = propriedades.tabela.colunaCriterio
End Function
Public Function getTableColor() As String
    getTableColor = propriedades.tabela.color
End Function
Public Function getTableTiposColunas() As String()
    getTableTiposColunas = propriedades.tabela.tiposColunas
End Function
Public Function tableHasHeader() As Boolean
    tableHasHeader = propriedades.tabela.hasHeader
End Function
Public Function getArgumento() As String
    getArgumento = propriedades.arrgumento
End Function
Public Function getTableRangeInicial() As Integer
    getTableRangeInicial = propriedades.tabela.intervalo.de
End Function
Public Function getTableRangeFinal() As Integer
    getTableRangeFinal = propriedades.tabela.intervalo.ate
End Function

'Funções Auxiliares
Public Function isImagem() As Boolean
    isImagem = (returnStringDelimited("(", ")", fieldTag) <> "" And _
                    returnStringDelimited("", "(", returnStringDelimited("{", "}", fieldTag)) = "imagem")
End Function
Public Function isTable() As Boolean
    isTable = (returnStringDelimited("(", ")", fieldTag) <> "" And _
                    returnStringDelimited("", "(", returnStringDelimited("{", "}", fieldTag)) = "tabela")
End Function
Private Function returnStringDelimited(ByVal left As String, ByVal right As String, _
                                        ByVal stringSource As String) As String
    Dim posLeft As Integer
    Dim posRight As Integer: posRight = InStr(1, stringSource, right, 0)
    
    If left = "" Then posLeft = 0 Else posLeft = InStr(1, stringSource, left, 0)
    
    If posLeft <= 0 And posRight <= 0 Then
        returnStringDelimited = ""
        Exit Function
    End If
        
    posLeft = posLeft + 1
    posRight = posRight - posLeft
    
    returnStringDelimited = Mid(stringSource, posLeft, posRight)
End Function
Private Function countColumns(ByVal namePlan As String) As Integer
    
    On Error GoTo erro
    Dim cols As Integer: cols = 0
    Dim plan As Worksheet: Set plan = ThisWorkbook.Worksheets(namePlan)
    Dim cell, rng As Range: Set rng = plan.Rows(1)
    
    For Each cell In rng.Cells
        If cell.value <> "" Then
            cols = cols + 1
        End If
    Next cell
    countColumns = cols
    Exit Function
erro:
    countColumns = -1
End Function

