VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''Desenvolvido por HARLLEM NASCIMENTO
''Email: harllemnascimento@gmail.com
''Em 14/05/2021

Dim plan As Worksheet
Dim fromCol, toCol As Integer
Dim src() As Variant
Dim construido As Boolean
Dim srcTypes() As String ''int; long; double; String; Date
Public Function constructor(ByVal namePlan As String, ByVal numberFirstColumn As Integer, ByVal numberLastColumn As Integer, _
                                ByRef dataType() As String)
    Set plan = ThisWorkbook.Worksheets(namePlan)
    fromCol = numberFirstColumn
    toCol = numberLastColumn
    If UBound(dataType, 1) = toCol - fromCol + 1 Then
        srcTypes = dataType
        construido = True
    Else
        MsgBox "Erro ao tentar construir o Objeto, o Array que especifica o tipo de dados das colunas da tabela" _
            & "(dataType argument) esta com o tamanho diferente da tabela especificada!!", vbCritical, "Erro"
        construido = False
    End If
End Function

Private Function verifyConstruido(Optional preMessage As String = "") As Boolean
    If Not construido Then
        MsgBox preMessage & "Por favor, chame o método construtor antes!!", vbCritical, "Erro!!"
    End If
    verifyConstruido = construido
End Function

Public Function toExcelRange(ByVal plotRow As Integer, ByVal plotCol As Integer, Optional w As Worksheet = Nothing) As Boolean
    
    If Not verifyConstruido("Metodo toExcelRange() diz: ") Then
        toExcelRange = False
        Exit Function
    End If
    
    Dim Rows, Cols, i, j As Integer
    Dim callInClass As Boolean: callInClass = False
    
    If w Is Nothing Then
        callInClass = True
        w = plan
    End If
    
    Rows = UBound(src, 1)
    Cols = UBound(src, 2)
    
    If Not callInClass Then
        ThisWorkbook.Application.ScreenUpdating = False
    End If
    For i = plotRow To plotRow + Rows - 1
        For j = plotCol To plotCol + Cols - 1
            If (i - plotRow + 1) <> 1 Then
                Select Case (srcTypes(j - plotRow + 1))
                    Case "int":
                        w.Cells(i, j).Value = CInt(src(i - plotRow + 1, j - plotRow + 1))
                    Case "double":
                        w.Cells(i, j).Value = CDbl(src(i - plotRow + 1, j - plotRow + 1))
                    Case "String":
                        w.Cells(i, j).Value = CStr(src(i - plotRow + 1, j - plotRow + 1))
                    Case "Date":
                        w.Cells(i, j).Value = CDate(src(i - plotRow + 1, j - plotRow + 1))
                    Case "long":
                        w.Cells(i, j).Value = CLng(src(i - plotRow + 1, j - plotRow + 1))
                End Select
            Else
                w.Cells(i, j).Value = CStr(src(i - plotRow + 1, j - plotRow + 1))
            End If
        Next j
    Next i
    If Not callInClass Then
        ThisWorkbook.Application.ScreenUpdating = True
    End If
    toExcelRange = True
End Function

Public Function toExcelFile(ByVal fileName As String)
    
    If Not verifyConstruido("Metodo toExcelFile() diz: ") Then
        Exit Function
    End If
    
    Dim arq As Workbook: Set arq = Workbooks.Add
    
    Workbooks.Application.ScreenUpdating = False
    Workbooks.Application.DisplayAlerts = False
    
    Dim temp As String: temp = "plan " & arq.Worksheets.Count
    arq.Worksheets(1).Name = temp
    
    Dim p As Worksheet: Set p = arq.Worksheets(temp)
    
    toExcelRange 1, 1, p
    arq.SaveAs ThisWorkbook.Path & "\" & fileName
    
    arq.Close fileName:=ThisWorkbook.Path & "\" & fileName & ".xlsx"
    Workbooks.Application.DisplayAlerts = True
    Workbooks.Application.ScreenUpdating = True
End Function

Public Function makeFilter(ByVal criteriaVal As String, ByVal criteriaCol As Integer)
    
    If Not verifyConstruido("Metodo makeFilter() diz: ") Then
        Exit Function
    End If
    
    Dim i, j, aLn, aCl As Integer: aLn = 0
    
    'Monta array principal
    For i = 2 To plan.Cells(Rows.Count, fromCol).End(xlUp).Row
        If CStr(criteriaVal) = CStr(plan.Cells(i, criteriaCol).Value) Then
            aLn = aLn + 1
        End If
    Next i
    aCl = (toCol - fromCol) + 1
    Dim arr() As Variant
    ReDim arr(1 To aLn + 1, 1 To aCl) As Variant
    aLn = 1
    aCl = 1
    
    For i = 1 To plan.Cells(Rows.Count, 1).End(xlUp).Row
        If i = 1 Then
            'Coloca cabeçalho
            aCl = 1
            For j = fromCol To toCol
                arr(aLn, aCl) = CStr(plan.Cells(i, j).Value)
                aCl = aCl + 1
            Next j
             aLn = aLn + 1
        Else
            If CStr(criteriaVal) = CStr(plan.Cells(i, criteriaCol).Value) Then
                'Faz o filtro
                aCl = 1
                For j = fromCol To toCol
                    
                    arr(aLn, aCl) = CStr(plan.Cells(i, j).Value)
                    
                    aCl = aCl + 1
                Next j
                
                aLn = aLn + 1
            End If
        End If
    Next i
    src = arr
End Function

