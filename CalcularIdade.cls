VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalcularIdade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private idadeCompleta As String
Private Ano As Integer

Public Function construtor(ByVal dataNascimento As Date) As String
    ''O Código abaixo foi adaptado de: https://www.usandoaccess.com.br/dicas/calculo-de-idade-anos-meses-e-dias.asp
    Dim Anos, Meses, Dias As Byte
    Dim DataRef As Date
    Dim Resultado As Boolean
    
    If dataNascimento >= Now Or dataNascimento = 0 Then construtor = "#ERRO"
     
    'Ajusta ano bissexto
    dataNascimento = IIf(Format(dataNascimento, "mm/dd") = "02/29", dataNascimento - 1, dataNascimento)

    Anos = Int((Format(Date, "yyyymmdd") - Format(dataNascimento, "yyyymmdd")) / 10000)
     
    Resultado = (Format(dataNascimento, "mmdd") > Format(Date, "mmdd"))
    
    DataRef = DateSerial(Year(Date) + Resultado, Format(dataNascimento, "mm"), Format(dataNascimento, "dd"))
    
    Meses = DateDiff("m", DataRef, Date) + (Format(dataNascimento, "dd") > Format(Date, "dd"))
     
    Resultado = (Format(dataNascimento, "dd") > Format(Date, "dd"))
    DataRef = DateSerial(Year(Date), Format(Date, "mm") + Resultado, Format(dataNascimento, "dd"))
    DataRef = IIf(Format(dataNascimento, "dd") <> Format(DataRef, "dd"), DataRef - Format(DataRef, "dd"), DataRef)
    
    Dias = CDbl(Date) - CDbl(DataRef)
    
    Ano = CInt(Anos)
    idadeCompleta = IIf(Anos <= 1, IIf(Anos = 0, "", Anos & " ano "), Anos & " anos ") & _
                                  IIf(Meses <= 1, IIf(Meses = 0, "", Meses & " mes "), Meses & " meses ") & _
                                  IIf(Dias <= 1, IIf(Dias = 0, "", Dias & " dia "), Dias & " dias ")
End Function

Public Function getIdade() As Integer
    getIdade = Ano
End Function

Public Function getIdadeCompleta() As String
    getIdadeCompleta = idadeCompleta
End Function
