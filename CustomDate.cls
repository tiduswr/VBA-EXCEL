VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DateStructure
    dia As Integer
    mes As Integer
    ano As Integer
End Type
Private Type Calendario
    bissexto As Variant
    normal As Variant
End Type
Private Type CustomDateStructure
    dataObjeto As DateStructure
    diasCalend As Calendario
    isInitialized As Boolean
End Type

Private calendConstruido As Boolean
Private this As CustomDateStructure

Public Function construtor(ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer) As Boolean
    If Not calendConstruido Then
        With this.diasCalend
            .bissexto = Array(0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
            .normal = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
        End With
        calendConstruido = True
    End If
    
    If Not isValid(dia, mes, ano) Then
        construtor = False
        Exit Function
    End If
    
    With this.dataObjeto
        .dia = dia
        .mes = mes
        .ano = ano
    End With
    this.isInitialized = True
    construtor = True
End Function

Private Function isBissexto(ByVal ano As Integer) As Boolean
    isBissexto = (ano Mod 4 = 0) And ((ano Mod 100 <> 0) Or (ano Mod 400 = 0))
End Function

Private Function isValid(ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer) As Boolean
    isValid = False
    If Not (mes >= 1 And mes <= 12) Then Exit Function
    If isBissexto(ano) Then
        If Not (dia >= 1 And dia <= this.diasCalend.bissexto(mes)) Then Exit Function
    Else
        If Not (dia >= 1 And dia <= this.diasCalend.normal(mes)) Then Exit Function
    End If
    isValid = True
End Function

Public Function getDate() As Date
    If Not this.isInitialized Then Exit Function
    With this.dataObjeto
        getDate = CDate(.dia & "/" & .mes & "/" & .ano)
    End With
End Function

Public Function compareTo_Days(ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer) As Long
    If Not this.isInitialized Then Exit Function
    If Not isValid(dia, mes, ano) Then
        Exit Function
    End If
    
    Dim iDias, fDias, iAno As Long
    Dim dif_anos As Long: dif_anos = 0
    Dim i As Long
    Dim dbissexto As Boolean
    
    iDias = this.dataObjeto.dia
    dbissexto = isBissexto(this.dataObjeto.ano)
    For i = this.dataObjeto.mes - 1 To 1 Step -1
        If dbissexto Then
            iDias = iDias + this.diasCalend.bissexto(i)
        Else
            iDias = iDias + this.diasCalend.normal(i)
        End If
    Next i
    
    fDias = dia
    dbissexto = isBissexto(ano)
    For i = mes - 1 To 1 Step -1
        If dbissexto Then
            fDias = fDias + this.diasCalend.bissexto(i)
        Else
            fDias = fDias + this.diasCalend.normal(i)
        End If
    Next i
    
    iAno = this.dataObjeto.ano
    Do While iAno < ano
        If isBissexto(iAno) Then
            dif_anos = dif_anos + 366
        Else
            dif_anos = dif_anos + 365
        End If
        iAno = iAno + 1
    Loop
    
    compareTo_Days = dif_anos - iDias + fDias
    
End Function

Public Function compareTo_Years(ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer) As Integer
    If Not this.isInitialized Then Exit Function
    If Not isValid(dia, mes, ano) Then
        Exit Function
    End If
    
    Dim totdias As Long: totdias = compareTo_Days(dia, mes, ano)
    Dim iAno As Integer: iAno = this.dataObjeto.ano
    compareTo_Years = 0
    
    Do While iAno < ano
        If isBissexto(iAno) Then
            If totdias >= 366 Then
                totdias = totdias - 366
                compareTo_Years = compareTo_Years + 1
            End If
        Else
            If totdias >= 365 Then
                totdias = totdias - 365
                compareTo_Years = compareTo_Years + 1
            End If
        End If
        iAno = iAno + 1
    Loop
    
End Function

Public Function compareTo_Month(ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer, _
                                    Optional diasRestantes As Boolean = False) As Integer
    If Not this.isInitialized Then Exit Function
    If Not isValid(dia, mes, ano) Then
        Exit Function
    End If
    
    Dim totdias As Long: totdias = compareTo_Days(dia, mes, ano)
    Dim iAno As Integer: iAno = this.dataObjeto.ano
    Dim i, init, fim As Integer
    Dim diasDesconto As Integer
    init = this.dataObjeto.mes
    If iAno <> ano Then
        fim = 12
    Else
        fim = mes - 1
    End If
    
    Do While iAno <= ano
        For i = init To fim
            If isBissexto(iAno) Then
                diasDesconto = this.diasCalend.bissexto(i)
            Else
                diasDesconto = this.diasCalend.normal(i)
            End If
            
            If totdias >= diasDesconto Then
                totdias = totdias - diasDesconto
                compareTo_Month = compareTo_Month + 1
            End If
        Next i
        init = 1
        iAno = iAno + 1
        If ano = iAno Then fim = mes
    Loop
    
    If diasRestantes Then compareTo_Month = totdias
    
End Function

Public Function compareTo_CompleteDate(ByVal dia As Integer, ByVal mes As Integer, ByVal ano As Integer, _
                                        Optional formatString As String = "%y anos, %m meses e %d dias") As String
    If Not this.isInitialized Then Exit Function
    If Not isValid(dia, mes, ano) Then
        Exit Function
    End If
    
    Dim anos, meses, dias, aux As Integer
    
    anos = compareTo_Years(dia, mes, ano)
    aux = this.dataObjeto.ano
    this.dataObjeto.ano = ano
    meses = compareTo_Month(dia, mes, ano)
    dias = compareTo_Month(dia, mes, ano, True)
    this.dataObjeto.ano = aux
    
    compareTo_CompleteDate = Replace(formatString, "%y", CStr(anos))
    compareTo_CompleteDate = Replace(compareTo_CompleteDate, "%m", CStr(meses))
    compareTo_CompleteDate = Replace(compareTo_CompleteDate, "%d", CStr(dias))
    
End Function
