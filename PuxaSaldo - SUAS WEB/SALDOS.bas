Attribute VB_Name = "SALDOS"
Option Explicit
Type SearchMunicipio
    ano As String
    mes As String
    esfereadm As String
    uf As String
    municipioIBGE As String
    municipioNOME As String
End Type
Sub PuxaSaldos()

Dim web As ChromeDriver
Dim pesquisa As SearchMunicipio
Dim tbl As TableElement
Dim arr()
Dim dateCol() As String
Dim tempo As Long
Dim tempoaux As Long
Dim uLin As Long

Set web = New ChromeDriver
uLin = ActiveSheet.Cells(Rows.Count, 6).End(xlUp).Row + 1
tempo = CLng(ActiveSheet.Range("B7").value)

'Pega informação da planilha

pesquisa.ano = CStr(ActiveSheet.Range("B4").value)
pesquisa.mes = formata_mes(CStr(ActiveSheet.Range("B5").value))
If CStr(ActiveSheet.Range("B6").value) = "MUNICIPAL" Then
    pesquisa.esfereadm = "M"
End If
pesquisa.uf = CStr(ActiveSheet.Range("D4").value)
'Acessa o site dos saldos mds
On Error GoTo erro1
'Verifica se esconde o navegador
If CStr(ActiveSheet.Range("D6").value) = "SIM" Then
    web.AddArgument ("--headless")
End If
web.Get "http://aplicacoes.mds.gov.br/suaswebcons/restrito/execute.jsf?b=*tbmepQbsdfmbtQbhbtNC&event=*fyjcjs"
web.Wait tempo
'Preenche campos

slctOptCB web, pesquisa.ano, "//*[@id=""form:ano""]", tempo
slctOptCB web, pesquisa.mes, "//*[@id=""form:mes""]", tempo
slctOptCB web, pesquisa.esfereadm, "//*[@id=""form:esferaAdministrativa""]", tempo
slctOptCB web, pesquisa.uf, "//*[@id=""form:uf""]", tempo

'Pesquisa municipio
web.Wait tempo * 2
pesquisa.municipioIBGE = Left(CStr(ActiveSheet.Range("D5").value), 6)
slctOptCB web, pesquisa.municipioIBGE, "//*[@id=""form:municipio""]", tempo
slctOptCB web, pesquisa.mes, "//*[@id=""form:mes""]", tempo
web.FindElementByXPath("//*[@id=""form:pesquisar""]").Click
'Verifica tabela em 10 vezes o tempo definido de delay
While HaveElement("//*[@id=""form:j_id173""]/table/tbody", web) = False
    tempoaux = tempoaux + tempo
    web.Wait tempo
    If tempoaux = tempo * 11 Then
        GoTo erro2
    End If
Wend
'Guarda no do municipio
pesquisa.municipioNOME = web.FindElementByXPath("//*[@id=""form:j_id141""]/center/fieldset/div/table/tbody/tr[2]/td[2]").text
'Seta tabela na memoria
If HaveElement("//*[@id=""form:j_id173""]/table/tbody", web) Then
    Set tbl = web.FindElementByXPath("//*[@id=""form:j_id173""]/table/tbody").AsTable
    arr() = tbl.Data
    'Usa formula personalizada e repassa dados para a planilha
    Application.ScreenUpdating = False
    Call toRANGE(Range("F" & uLin), arr, 1, 2, pesquisa.municipioNOME, pesquisa.mes & "/" & pesquisa.ano)
    Application.ScreenUpdating = True
Else
    GoTo erro2
End If

'Encerra programa
fim:
web.Close
Set web = Nothing
Set tbl = Nothing
tempo = 0
tempoaux = 0
MsgBox "Processo concluido!", vbInformation, "Consegui :)"
Exit Sub

erro1:
MsgBox "Não consegui entrar no sistema, verifique sua conexão ou se o site esta disponivel", vbCritical
web.Close
Set web = Nothing
Set tbl = Nothing
tempo = 0
tempoaux = 0
Application.ScreenUpdating = True
Exit Sub

erro2:
MsgBox "Tabela de saldos não encontrada no site!, verifique sua internet ou se o mês esta disponivel no sistema ou se você deixou algum dado de pesquisa em branco!", vbCritical
web.Close
Set web = Nothing
Set tbl = Nothing
tempo = 0
tempoaux = 0
Application.ScreenUpdating = True
End Sub
Sub slctOptCB(webdriver As ChromeDriver, value As Variant, xpath As String, tempo As Long)
Dim ELEMENT As SelectElement
Dim tempoaux As Long
Set ELEMENT = webdriver.FindElementByXPath(xpath).AsSelect
ELEMENT.SelectByValue (value)
Set ELEMENT = Nothing
webdriver.Wait tempo
End Sub
Function HaveElement(ByVal xpath As String, webdriver As ChromeDriver) As Boolean
Dim check As By
Set check = New By
HaveElement = webdriver.IsElementPresent(check.xpath(xpath))
Set check = Nothing
End Function
Private Sub toRANGE(ByRef target As Range, ByRef arr() As Variant, ByVal startline As Integer, ByVal startcolumn As Integer, nomeMUNICIPIO As String, ref As String)

Dim ln As Long, i As Long, x As Long
Dim col As Long, j As Long
Dim W As Worksheet

Set W = target.Worksheet
col = UBound(arr, 2)
ln = UBound(arr, 1)
x = 1

'Função exclusiva pra essa planilha
For i = startline To ln
    On Error Resume Next
    For j = startcolumn To col
        If arr(i, 4) <> 0 Then
            If j = 5 Then
                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = CDbl(Replace(Replace(arr(i, j), vbLf, ""), "R$", ""))
            Else
                W.Cells(target.Row + (x - startline), target.Column + (j - startcolumn)).value = Replace(Replace(arr(i, j), vbLf, ""), "R$", "")
            End If
        End If
    Next j
    If arr(i, 4) <> 0 Then
        'Adiciona codigo do ibge
        W.Cells(target.Row + (x - startline), target.Column + 4).value = nomeMUNICIPIO
        W.Cells(target.Row + (x - startline), target.Column + 5).value = ref
        x = x + 1
    End If
    
Next i

Set W = Nothing
col = 0
ln = 0

End Sub
Function formata_mes(ByVal val As String) As String
If Len(val) < 2 Then
    formata_mes = 0 & val
Else
    formata_mes = val
End If
End Function


