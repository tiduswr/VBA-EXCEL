VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MakeWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Desenvolvido por Harllem Nascimento
'Em 19 de maio de 2021 foi concluida a vers�o final do projeto
'contato: harllemnascimento@gmail.com

'------------------------------------------
'Tipo usado para o tema das colunas do replaceWithTable():

'Public Type ColumnColorTheme
'    alignment As Integer
'    bold As Boolean
'    color As Long
'    backColor As Long
'End Type
'-----------------------------------------

'Atributos privados
Private Type app
    WordApp As Word.Application
    Doc As Word.Document
    DocOpen As Boolean
End Type
Private this As app


''Construtor
Public Function constructor(ByVal Hide As Boolean)
    ''Intancia objeto Word
    Set this.WordApp = CreateObject("Word.Application")
    ''Verifica a visibilidade passada no argumento hide
    If Hide Then
        this.WordApp.Visible = False
        this.WordApp.ScreenUpdating = False
        
    Else
        this.WordApp.Visible = True
    End If
    this.WordApp.DisplayAlerts = wdAlertsNone
    
    ''Altera o atributo que informa se o word foi instanciado
    this.DocOpen = True
    Exit Function
End Function

'Utilitarios
Private Function verifyPath(ByVal path As String) As Boolean
    If Dir(path) = vbNullString Then
        verifyPath = False
    Else
        verifyPath = True
    End If
End Function
Private Function fileIsOpen(strFicheiro As String)
    Dim intFicheiro As Long
    Dim intErro As Long
    
    On Error Resume Next
    intFicheiro = FreeFile()
    Open strFicheiro For Binary Lock Read As #intFicheiro
    Close intFicheiro
    intErro = Err
    On Error GoTo 0
    
    Select Case intErro
        Case 0: fileIsOpen = False
        Case 70: fileIsOpen = True
        Case Else: Error intErro
    End Select
End Function

''A��es da aplica��o
Public Function killApplication()
    Set this.WordApp = Nothing
    Shell "TASKKILL /F /IM Winword.exe", vbHide
    this.DocOpen = False
End Function
Public Function isOpen() As Boolean
    isOpen = this.DocOpen
End Function
Public Function saveDoc(ByVal path As String, ByVal nameArq As String) As Boolean
    If isOpen() Then
        this.Doc.SaveAs path & "\" & nameArq, wdFormatDocumentDefault
        saveDoc = True
    Else
        saveDoc = False
    End If
End Function
Public Function closeDoc() As Boolean
    If isOpen() Then
        this.Doc.Application.Quit wdDoNotSaveChanges
        closeDoc = True
    Else
        closeDoc = False
    End If
End Function
Private Function gotoFirstLine()
    this.Doc.Application.Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=1
End Function
Public Function openDoc(ByVal path As String) As Boolean
    If isOpen() And Not fileIsOpen(path) Then
        If verifyPath(path) Then
            openDoc = True
            Set this.Doc = this.WordApp.Documents.Open(path)
        Else
            openDoc = False
            Call closeDoc
        End If
    Else
        openDoc = False
        Call closeDoc
    End If
    this.Doc.Application.DisplayAlerts = wdAlertsNone
End Function
Public Function toPDF(ByVal path As String, ByVal nameArq As String) As Boolean
    If isOpen() Then
        this.Doc.SaveAs path & "\" & nameArq, FileFormat:=wdFormatPDF
        toPDF = True
    Else
        toPDF = False
    End If
End Function

''Replace functions
Function replaceWithPicture(ByVal ReplaceKey As String, ByVal path As String, ByVal Width As Integer, _
                            ByVal Height As Integer, Optional secao As String = "body") As Boolean
    
    Dim shape As Word.InlineShape
    Dim replaced As Boolean: replaced = False
    
    ''Verifica onde ser� feito o replace
    Select Case secao
        Case "body":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        Case "header":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        Case "footer":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    End Select
    
    If isOpen() Then
        If verifyPath(path) Then
            While Not replaced
                this.Doc.Application.Selection.Find.Text = ReplaceKey
                this.Doc.Application.Selection.Find.Execute
                If this.Doc.Application.Selection.Find.Found Then
                    ''Adiciona a imagem ao documento e seta na variavel
                    Set shape = this.Doc.Application.Selection.InlineShapes.AddPicture(path)
                    'Redimensiona a imagem
                    If Width <> 0 And Height <> 0 Then
                        shape.Width = Width
                        shape.Height = Height
                    End If
                Else
                    replaced = True
                End If
            Wend

            Set shape = Nothing
            replaceWithPicture = True
        Else
            replaceWithPicture = False
        End If
    Else
        replaceWithPicture = False
    End If
    
    this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    gotoFirstLine
End Function
Function replaceWithText(ByVal ReplaceKey As String, ByVal valtoReplace As Variant, Optional typeData As String = "String", _
                            Optional secao As String = "body") As Boolean
    
    Dim replaced As Boolean: replaced = False
    
    ''Verifica onde ser� feito o replace
    Select Case secao
        Case "body":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        Case "header":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        Case "footer":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    End Select

    If isOpen() Then
        While Not replaced
            this.Doc.Application.Selection.Find.Text = ReplaceKey
            this.Doc.Application.Selection.Find.Execute
            If this.Doc.Application.Selection.Find.Found Then
                Select Case typeData
                    Case "String":
                        this.Doc.Application.Selection.Range = CStr(valtoReplace)
                    Case "Date":
                        this.Doc.Application.Selection.Range = CDate(valtoReplace)
                    Case "long":
                        this.Doc.Application.Selection.Range = CLng(valtoReplace)
                    Case "int":
                        this.Doc.Application.Selection.Range = CInt(valtoReplace)
                    Case "double":
                        this.Doc.Application.Selection.Range = CDbl(valtoReplace)
                End Select
            Else
                replaced = True
            End If
        Wend
        replaceWithText = True
    Else
        replaceWithText = False
    End If
    
    this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    gotoFirstLine
End Function

Public Function replaceWithTable(ByRef arr() As Variant, ByVal ReplaceKey As String, _
                            ByRef columnsTheme() As ColumnColorTheme, _
                            Optional secao As String = "body", _
                            Optional HeaderTableColor As Long = vbBlue, _
                            Optional HeaderFontColor As Long = vbWhite, _
                            Optional HeaderAlignment As Long = wdAlignParagraphCenter, _
                            Optional semTemas As Boolean = False) As Boolean
    
    Dim lns, cols As Integer
    Dim i, j As Long
    Dim tbl As Word.Table
    Dim replaced As Boolean
    
    ''Verifica onde ser� feito o replace
    Select Case secao
        Case "body":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        Case "header":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        Case "footer":
            this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    End Select
    
    ''Verifica o array de temas, se estiver vazio bota o tema padr�o
    cols = UBound(arr, 2)
    If Not semTemas Then
        If Not isEmptyArray(columnsTheme) Then
            If UBound(columnsTheme, 1) < cols Then
                If UBound(columnsTheme, 1) = 0 Then
                    ReDim columnsTheme(1 To cols) As ColumnColorTheme
                    For i = 1 To cols
                        columnsTheme(i).backColor = vbWhite
                        columnsTheme(i).bold = False
                        columnsTheme(i).color = vbBlack
                        columnsTheme(i).alignment = wdAlignParagraphLeft
                    Next i
                Else
                    MsgBox "Erro, temas das colunas � menor que as colunas existentes no array de referencia!", vbCritical
                    replaceWithTable = False
                    closeDoc
                    Exit Function
                End If
            End If
        Else
            ReDim columnsTheme(1 To cols) As ColumnColorTheme
            For i = 1 To cols
                columnsTheme(i).backColor = vbWhite
                columnsTheme(i).bold = False
                columnsTheme(i).color = vbBlack
                columnsTheme(i).alignment = wdAlignParagraphLeft
            Next i
        End If
    End If
    
    If isOpen() Then
        While Not replaced
            this.Doc.Application.Selection.Find.Text = ReplaceKey
            this.Doc.Application.Selection.Find.Execute
            If this.Doc.Application.Selection.Find.Found Then
                'Cria a tabela
                lns = UBound(arr, 1)
                Set tbl = this.Doc.tables.Add(this.Doc.Application.Selection.Range, lns, cols)
                tbl.Borders.Enable = True
                
                For i = LBound(arr, 1) To lns
                    For j = LBound(arr, 2) To cols
                        tbl.cell(i, j).Range.Text = arr(i, j)
                        'Formata Header
                        If i = 1 Then
                            tbl.cell(i, j).Shading.BackgroundPatternColor = HeaderTableColor
                            tbl.cell(i, j).Range.bold = True
                            tbl.cell(i, j).Range.Font.TextColor = HeaderFontColor
                            tbl.cell(i, j).Range.ParagraphFormat.alignment = HeaderAlignment
                        Else
                            If Not semTemas Then
                                tbl.cell(i, j).Shading.BackgroundPatternColor = columnsTheme(j).backColor
                                tbl.cell(i, j).Range.bold = columnsTheme(j).bold
                                tbl.cell(i, j).Range.Font.TextColor = columnsTheme(j).color
                                tbl.cell(i, j).Range.ParagraphFormat.alignment = columnsTheme(j).alignment
                            End If
                        End If
                    Next j
                    ''Aplica os espa�amentos da tabela
                    With tbl.Rows(i).Range.ParagraphFormat
                        .spaceAfter = 0
                        .spaceBefore = 0
                        .lineSpacingRule = wdLineSpaceSingle
                    End With
                Next i
                
                ''organiza a tabela
                tbl.Rows.alignment = wdAlignRowCenter
                tbl.Columns.AutoFit
                
                Set tbl = Nothing
            Else
                replaced = True
            End If
        Wend
        replaceWithTable = True
    Else
        replaceWithTable = False
    End If
    
    this.Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    gotoFirstLine
End Function

Private Function isEmptyArray(ByRef arr() As ColumnColorTheme)
    On Error GoTo erro
    If UBound(arr, 1) < 1 Then
        isEmptyArray = True
        Exit Function
    End If
    isEmptyArray = False
    Exit Function
erro:
    isEmptyArray = True
End Function
