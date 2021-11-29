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

Public Function construtor(ByVal arg As Date) As Boolean
    Dim d As DateStructure: d = dateToStructureDate(arg)
    If Not calendConstruido Then
        With this.diasCalend
            .bissexto = Array(0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
            .normal = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
        End With
        calendConstruido = True
    End If
    
    If Not isValid(d.dia, d.mes, d.ano) Then
        construtor = False
        Exit Function
    End If
    With this.dataObjeto
        .dia = d.dia
        .mes = d.mes
        .ano = d.ano
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

Private Function dateToStructureDate(ByVal arg As Date) As DateStructure
    With dateToStructureDate
        .dia = Day(arg)
        .mes = Month(arg)
        .ano = Year(arg)
    End With
End Function

Public Function setDate(ByVal arg As Date) As Boolean
    Dim d As DateStructure: d = dateToStructureDate(arg)
    If Not isValid(d.dia, d.mes, d.ano) Then
        setDate = False
        Exit Function
    End If
    With this.dataObjeto
        .dia = d.dia
        .mes = d.mes
        .ano = d.ano
    End With
End Function

Public Function getDate() As Date
    If Not this.isInitialized Then Exit Function
    With this.dataObjeto
        getDate = CDate(.dia & "/" & .mes & "/" & .ano)
    End With
End Function

Public Function compareTo_Days(ByVal arg As Date) As Long
    Dim d As DateStructure: d = dateToStructureDate(arg)
    If Not this.isInitialized Then Exit Function
    If Not isValid(d.dia, d.mes, d.ano) Then
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
    
    fDias = d.dia
    dbissexto = isBissexto(d.ano)
    For i = d.mes - 1 To 1 Step -1
        If dbissexto Then
            fDias = fDias + this.diasCalend.bissexto(i)
        Else
            fDias = fDias + this.diasCalend.normal(i)
        End If
    Next i
    
    iAno = this.dataObjeto.ano
    Do While iAno < d.ano
        If isBissexto(iAno) Then
            dif_anos = dif_anos + 366
        Else
            dif_anos = dif_anos + 365
        End If
        iAno = iAno + 1
    Loop
    
    compareTo_Days = dif_anos - iDias + fDias
    
End Function

Public Function compareTo_Years(ByVal arg As Date) As Integer
    Dim d As DateStructure: d = dateToStructureDate(arg)
    If Not this.isInitialized Then Exit Function
    If Not isValid(d.dia, d.mes, d.ano) Then
        Exit Function
    End If
    
    Dim totdias As Long: totdias = compareTo_Days(arg)
    Dim iAno As Integer: iAno = this.dataObjeto.ano
    compareTo_Years = 0
    
    Do While iAno < d.ano
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

Public Function compareTo_Month(ByVal arg As Date, Optional diasRestantes As Boolean = False) As Integer
    Dim d As DateStructure: d = dateToStructureDate(arg)
    If Not this.isInitialized Then Exit Function
    If Not isValid(d.dia, d.mes, d.ano) Then
        Exit Function
    End If
    
    Dim totdias As Long: totdias = compareTo_Days(arg)
    Dim iAno As Integer: iAno = this.dataObjeto.ano
    Dim i, init, fim As Integer
    Dim diasdesconto As Integer
    init = this.dataObjeto.mes
    If iAno <> d.ano Then
        fim = 12
    Else
        fim = d.mes - 1
    End If
    
    Do While iAno <= d.ano
        For i = init To fim
            If isBissexto(iAno) Then
                diasdesconto = this.diasCalend.bissexto(i)
            Else
                diasdesconto = this.diasCalend.normal(i)
            End If
            
            If totdias >= diasdesconto Then
                totdias = totdias - diasdesconto
                compareTo_Month = compareTo_Month + 1
            End If
        Next i
        init = 1
        iAno = iAno + 1
        If d.ano = iAno Then fim = d.mes
    Loop
    
    If diasRestantes Then compareTo_Month = totdias
    
End Function

Public Function compareTo_CompleteDate(ByVal arg As Date, Optional formatString As String = "%y anos, %m meses e %d dias") As String
    Dim d As DateStructure: d = dateToStructureDate(arg)
    If Not this.isInitialized Then Exit Function
    If Not isValid(d.dia, d.mes, d.ano) Then
        Exit Function
    End If
    
    Dim anos, meses, dias, aux As Integer
    
    anos = compareTo_Years(arg)
    aux = this.dataObjeto.ano
    this.dataObjeto.ano = d.ano
    meses = compareTo_Month(arg)
    dias = compareTo_Month(arg, True)
    this.dataObjeto.ano = aux
    
    compareTo_CompleteDate = Replace(formatString, "%y", CStr(anos))
    compareTo_CompleteDate = Replace(compareTo_CompleteDate, "%m", CStr(meses))
    compareTo_CompleteDate = Replace(compareTo_CompleteDate, "%d", CStr(dias))
    
End Function

Public Function sumDays(ByVal qtd As Long) As Date
    If Not this.isInitialized Then Exit Function
    Dim diasSomar As Integer
    qtd = qtd + this.dataObjeto.dia
    
    While qtd > 0
        With this.dataObjeto
            If isBissexto(.ano) Then
                diasSomar = this.diasCalend.bissexto(.mes)
            Else
                diasSomar = this.diasCalend.normal(.mes)
            End If
            
            If qtd > diasSomar Then
                .dia = 1
                If .mes < 12 Then
                    .mes = .mes + 1
                Else
                    .mes = 1
                    .ano = .ano + 1
                End If
                qtd = qtd - diasSomar
            Else
                .dia = qtd
                qtd = 0
            End If
        End With
    Wend
    sumDays = getDate
End Function

Public Function sumMonths(ByVal qtd As Long) As Date
    Dim i, m, y As Long
    Dim anos, meses As Integer
    Dim totDays, diasSomar As Long
    
    meses = this.dataObjeto.mes
    anos = this.dataObjeto.ano
    
    For i = 1 To qtd
        If isBissexto(anos) Then
            diasSomar = this.diasCalend.bissexto(meses)
        Else
            diasSomar = this.diasCalend.normal(meses)
        End If
        If meses < 12 Then
            meses = meses + 1
        Else
            anos = anos + 1
            meses = 1
        End If
        totDays = totDays + diasSomar
    Next i
    sumMonths = sumDays(totDays)
End Function

Public Function sumYears(ByVal qtd As Integer) As Date
    sumYears = sumMonths(qtd * 12)
End Function

