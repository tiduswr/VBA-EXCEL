Attribute VB_Name = "PARCELAS_PAGAS"
Option Explicit
Type search
    ano As String
    uf As String
    municipioIBGE As String
    agrupamento As String
    municipioNOME As String
End Type

Sub PuxaParcelas()
    Dim web As ChromeDriver
    Dim pesquisa As search
    Dim tempo As Long, uLin As Long
    Dim tbl As TableElement
    Dim arr()
    
    Set web = New ChromeDriver
    tempo = ActiveSheet.Range("B7").value
    uLin = ActiveSheet.Cells(Rows.Count, 4).End(xlUp).Row + 1
    
    'Acessa site
    If CStr(ActiveSheet.Range("B8").value) = "SIM" Then
        web.AddArgument ("--headless")
    End If
    web.Get ("https://aplicacoes.mds.gov.br/suaswebcons/restrito/execute.jsf?b=*dpotvmubsQbsdfmbtQbhbtNC&event=*fyjcjs")
    
    'pega dados para pesquisa
    pesquisa.ano = CStr(ActiveSheet.Range("B6").value)
    pesquisa.uf = CStr(ActiveSheet.Range("B4").value)
    pesquisa.municipioIBGE = Left(CStr(ActiveSheet.Range("B5").value), 6)
    pesquisa.agrupamento = "GRUPO"
    'Preenche formulario no site
    slctOptCB web, pesquisa.ano, "//*[@id=""form:ano""]", tempo
    slctOptCB web, pesquisa.uf, "//*[@id=""form:uf""]", tempo
    'Tempo de espera  para selecionar municipio E PESQUISAR
    On Error GoTo erro1
    web.Wait tempo * 2
    slctOptCB web, pesquisa.municipioIBGE, "//*[@id=""form:municipio""]", tempo
    slctOptCB web, pesquisa.agrupamento, "//*[@id=""form:agrupamento""]", tempo
    web.FindElementByXPath("//*[@id=""form:pesquisar""]").Click
    'Verifica tabela em 10 vezes o tempo definido de delay
    On Error GoTo erro2
    Do While HaveElement("//*[@id=""j_id176:datatableprincipal:tb""]", web) = False
        Dim tempoaux As Long
        tempoaux = tempoaux + tempo
        web.Wait tempo
        If tempoaux = tempo * 11 Then
            Exit Do
        End If
    Loop
    pesquisa.municipioNOME = CutInHalf(CStr(ActiveSheet.Range("B5").value), 2)
    Set tbl = web.FindElementByXPath("//*[@id=""j_id176:datatableprincipal:tb""]").AsTable
    arr() = tbl.Data
    
    Call toRangeParcelas(Range("D" & uLin), arr(), 1, 1, pesquisa.municipioNOME)
    
    web.Close
    Set web = Nothing
    tempo = 0
    Set tbl = Nothing
    uLin = 0
    MsgBox "Processo concluido!", vbInformation, "Consegui :)"
    Exit Sub
    
erro1:
    MsgBox "Não consegui entrar no sistema, verifique sua conexão ou se o site esta disponivel", vbCritical
    web.Close
    Set web = Nothing
    tempo = 0
    Set tbl = Nothing
    uLin = 0
    Exit Sub
    
erro2:
    MsgBox "Tabela de saldos não encontrada no site!, verifique sua internet ou se o mês esta disponivel no sistema ou se você deixou algum dado de pesquisa em branco!", vbCritical
    web.Close
    Set web = Nothing
    tempo = 0
    Set tbl = Nothing
    uLin = 0
End Sub
Function CutInHalf(ByVal text As String, ByVal pos As Integer)
    Dim arr() As String
    arr() = Split(text, " | ")
    CutInHalf = arr(pos - 1)
End Function

Private Sub toRangeParcelas(ByRef target As Range, ByRef arr() As Variant, ByVal startline As Integer, ByVal startcolumn As Integer, nomeMUNICIPIO As String)

    Dim ln As Long, i As Long, x As Long
    Dim col As Long, j As Long
    Dim W As Worksheet
    Dim tempHeader As String
    
    Set W = target.Worksheet
    col = UBound(arr, 2)
    ln = UBound(arr, 1)
    x = 1
    
    'Função exclusiva pra essa planilha
    For i = startline To ln
        On Error Resume Next
        'Verifica o tipo de Piso referente ao dinheiro
        If arr(i, 1) <> "" And arr(i, 6) = Empty Then
            tempHeader = arr(i, 1)
        End If
        For j = startcolumn To col
            If arr(i, 1) = "FUNDO MUNICIPAL" Then
                If j = 11 Or j = 10 Then
                'Move Saldo Liquido para coluna 9 e apaga coluna bugada de descontos/bloqueios
                    Select Case j
                        Case 11
                            W.Cells(target.Row + (x - startline), (target.Column + (j - startcolumn)) - 1).value = CDbl(arr(i, j))
                        Case 10
                            W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn) + 1).ClearContents
                    End Select
                Else
                    If j = 1 Or j = 4 Or j = 5 Or j = 8 Then
                        'formata coluna de datas, insere o nome do Municipio e formata coluna de saldo liquido
                        Select Case j
                            Case 1
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = nomeMUNICIPIO
                            Case 4
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = tempHeader
                            Case 5
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = CDate(arr(i, j))
                            Case 8
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn) + 1).value = CDbl(arr(i, j))
                        End Select
                    Else
                        'Imprime sem a coluna bugada de bloqueios/descontos
                        If j <> 9 Then
                            If j = 7 Then
                            'Formata a agencia e conta
                                Dim agencia_conta() As String
                                agencia_conta() = Split(Trim(arr(i, j)), "/")
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = agencia_conta(0)
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn) + 1).value = agencia_conta(1)
                            Else
                                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = Trim(arr(i, j))
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        If arr(i, 1) = "FUNDO MUNICIPAL" Then
            x = x + 1
        End If
    Next i
    
    Set W = Nothing
    col = 0
    ln = 0
    x = 0
    tempHeader = ""

End Sub
