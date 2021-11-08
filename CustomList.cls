VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------CONTROLE DE LIST VIEW CUSTOMIZADA:
'Autor: Harllem Nascimento(TidusWR)
'E-Mail: harllemnascimento@gmail.com
'Data de finalização do projeto: 08/11/2021

'------------Definições:
' "r" significa "row" ou "linha";
' "c" significa "column" ou "coluna"

'------------Como criar o ListBox Customizado:
'- Para utilizar esse objeto crie uma listbox customizada contendo uma label de fundo para alteração de cor ao passar o
'mouse encima de outra label que é utilizada para o evento click. Essas duas labels tem que ter um prefixo qualquer em
'que o objeto ira caminhar através da string ("r" & linha), por exemplo:

    'Prefixo = "ListBoxClick_" ou "ListBoxBackground_"
    'Como o obeto caminha atraves dos objetos do listbox = "r1"
    'Nome completo de um dos objetos "ListBoxClick_r1" ou "ListBoxBackground_r1"

'- No meio dessas duas labels pode ter incontaveis outras labels contendo um prefixo da tabela que precedera do numero
'da linha e da coluna no padrão ("r" & linha & "c" & coluna). Por exemplo:

    'Prefixo = "ListBox_"
    'Como o objeto caminha atraves dos objetos do listbox = "r1c1"
    'Nome completo de um dos objetos = "ListBox_r1c1"

'Atenção! tenha certeza de renomear todos os itens que fazem parte dessa listbox customizada em um padrão unico,
'caso contrario o programa vai dar erro!

'------------Metodo de uso:
'- Use primeiro o construtor() para carregar todos os objetos do listbox na memória;

'------------Observações:
'Você pode utilizar outros métodos do ListBox que são publicos da forma que quiser, mas com o construtor preenchido
'ela ja vai executar todas as ações que foi programada, contanto que sejam passados os parametros corretamente

'Se o dataSource() começar na coluna 0, vai ser considerado a coluna 0 como Key
'No Construtor o argumento keyColumn diz respeito a coluna que será usada como padrão para busca e ordennação!
'------------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Type list_objects
    vPrefixo As String
    vPrefixoClick As String
    vPrefixoBackground As String
    vQtdLinhas As Long
    vQtdColunas As Long

    vKeys() As Variant
    vSource() As Variant
    vFilter() As Variant
    vBackup() As Variant
    vColWithPicture As Long
    vInitialized As Boolean
End Type

Private Type sourceinfo
    vQtdLin As Long
    vQtdCol As Long
    vFirstLoaded As Long
End Type

Private vForm As UserForm
Attribute vForm.VB_VarHelpID = -1
Private WithEvents vBackList As MSForms.label
Attribute vBackList.VB_VarHelpID = -1
Private WithEvents scrollList As MSForms.ScrollBar
Attribute scrollList.VB_VarHelpID = -1
Private WithEvents labelAZ As MSForms.label
Attribute labelAZ.VB_VarHelpID = -1
Private WithEvents labelZA As MSForms.label
Attribute labelZA.VB_VarHelpID = -1
Private WithEvents txtbFilter As MSForms.TextBox
Attribute txtbFilter.VB_VarHelpID = -1
Private picCol, keyCol As Long
Private rows() As CustomLabel
Private cols() As CustomLabel
Private sourceinfo As sourceinfo
Private objects As list_objects
Private pgSkp As Boolean
Private hasKeyCol As Boolean

'Confirgurações
Public Property Get objPrefixo() As String
    objPrefixo = objects.vPrefixo
End Property
Public Property Get objPrefixoClick() As String
    objPrefixoClick = objects.vPrefixoClick
End Property
Public Property Get objPrefixoBackground() As String
    objPrefixoBackground = objects.vPrefixoBackground
End Property
Public Property Get srcQtdLinhas() As Long
    srcQtdLinhas = sourceinfo.vQtdLin
End Property
Public Property Get srcQtdColunas() As Long
    srcQtdColunas = sourceinfo.vQtdCol
End Property
Public Property Get objQtdLinhas() As Long
    objQtdLinhas = objects.vQtdLinhas
End Property
Public Property Get objQtdColunas() As Long
    objQtdColunas = objects.vQtdColunas
End Property
Public Property Get keyColumn() As Long
    keyColumn = keyCol
End Property
Public Function setKeyColumn(ByVal newKey As Long) As Boolean
    If newKey <> picCol And newKey > 0 Then
        keyCol = newKey
        setKeyColumn = True
    Else
        setKeyColumn = False
    End If
End Function
Public Function getValuesByColumnID(ByVal value As Variant) As Variant()
    If Not IsEmpty(objects.vSource) And objects.vInitialized And hasKeyCol Then
        Dim i As Long
        Dim initRow As Long: initRow = 1
        Dim ret() As Variant
        Dim initCol As Long: initCol = LBound(objects.vSource, 2)
        
        For i = initRow To sourceinfo.vQtdLin
            If objects.vSource(i, initCol) = value Then
                ReDim ret(initCol To sourceinfo.vQtdCol) As Variant
                Dim j As Long
                
                For j = initCol To sourceinfo.vQtdCol
                    ret(j) = objects.vSource(i, j)
                Next j
                
                getValuesByColumnID = ret
                Exit Function
            End If
        Next i
    End If
    getValuesByColumnID = ret
End Function

'Métodos -> no Construtor o argumento keyColumn diz respeito a coluna que será usada como padrão para busca e ordennação!
Public Function construtor(ByVal keyColumn As Long, _
                            ByVal prefixo As String, _
                            ByVal prefixoClick As String, _
                            ByVal prefixoBackground As String, _
                            ByVal linhas As Long, _
                            ByVal colunas As Long, _
                            ByRef form As UserForm, _
                            ByRef backgroundLabelName As String, _
                            ByRef dataSource() As Variant, _
                            Optional ByRef scroll As MSForms.ScrollBar = Nothing, _
                            Optional ByRef lblAz As MSForms.label = Nothing, _
                            Optional ByRef lblZa As MSForms.label = Nothing, _
                            Optional ByRef textBoxFilter As MSForms.TextBox = Nothing, _
                            Optional pictureColumn As Long = 0, _
                            Optional backColor As Long = vbBlue, _
                            Optional captionColor As Long = vbWhite, _
                            Optional captionBold As Boolean = False, _
                            Optional ByVal pageSkipMode As Boolean = False)

    Dim i As Long, j As Long
    
    If LBound(dataSource, 1) = 0 Then
        MsgBox "A propriedade dataSource() tem que receber um array iniciando obrigatoriamente pelo numero 1!", vbCritical, "Erro!"
        Exit Function
    End If
    hasKeyCol = (LBound(dataSource, 2) = 0)
    
    objects.vPrefixo = prefixo
    objects.vPrefixoBackground = prefixoBackground
    objects.vPrefixoClick = prefixoClick
    objects.vQtdLinhas = linhas
    objects.vQtdColunas = colunas
    Set vForm = form
    pgSkp = pageSkipMode
    
    sourceinfo.vQtdLin = UBound(dataSource, 1)
    sourceinfo.vQtdCol = UBound(dataSource, 2)
    
    objects.vSource = dataSource
    objects.vBackup = objects.vSource
    objects.vInitialized = True
    If hasKeyCol Then ReDim objects.vKeys(1 To objects.vQtdLinhas + 1)
    
    ''Scoll da lista
    Set scrollList = scroll
    attScroll sourceinfo.vQtdLin
        
    ''Botoes de ordenação
    Set labelAZ = lblAz
    Set labelZA = lblZa
    picCol = pictureColumn
    keyCol = keyColumn
    
    ''Textbox de busca
    Set txtbFilter = textBoxFilter
    
    ''Configura as animações
    Set vBackList = vForm.Controls(backgroundLabelName)
    ReDim rows(1 To linhas) As CustomLabel
    ReDim cols(1 To linhas, 1 To colunas) As CustomLabel
    
    For i = 1 To linhas
        Set rows(i) = New CustomLabel
        rows(i).construtor vForm.Controls(prefixoClick & "r" & i), backColor, captionColor, captionBold, _
                            redirectAnimation:=vForm.Controls(prefixoBackground & "r" & i)
        For j = 1 To colunas
            Set cols(i, j) = New CustomLabel
            cols(i, j).construtor vForm.Controls(prefixoClick & "r" & i), backColor, captionColor, captionBold, _
                            redirectAnimation:=vForm.Controls(prefixo & "r" & i & "c" & j), _
                            clickSpecialEffectSunken:=False
        Next j
    Next i
    
    loadData_inRange 1
End Function

Private Function setValue(ByVal controlRef As String, ByVal value As String, ByVal DataType As String)
    If DataType = "string" Then vForm.Controls(objects.vPrefixo & controlRef).Caption = value
    If DataType = "picture" Then vForm.Controls(objects.vPrefixo & controlRef).Picture = LoadPicture(value)
End Function

Private Function setVisible(ByVal row As Long, ByVal isVisible As Boolean)
    Dim i As Long
    
    For i = 1 To objects.vQtdColunas
        vForm.Controls(objects.vPrefixo & "r" & row & "c" & i).Visible = isVisible
    Next i
    vForm.Controls(objects.vPrefixoClick & "r" & row).Visible = isVisible
    vForm.Controls(objects.vPrefixoBackground & "r" & row).Visible = isVisible
End Function

Public Function getKeyFromListRow(ByVal vRow As Long) As Variant
    If (vRow > 0 And vRow <= objects.vQtdLinhas) And hasKeyCol Then
        getKeyFromListRow = objects.vKeys(vRow)
    Else
        getKeyFromListRow = -1
    End If
End Function

Public Function update(ByVal BubSrtMetodo As String, ByVal BubSrtColumn As String, Optional ByVal colWithPic As Long = 0)
    
    If Not objects.vInitialized Then Exit Function
    
    Dim linCharge As Long
    Dim firstCharge As Long
    Dim colCharge As Long
    Dim i As Long, j As Long
    
    firstCharge = 1
    
    'Ativa todas as linhas do form
    For i = 1 To objects.vQtdLinhas
        setVisible i, True
    Next i
    
    'Verifica o maximo de linhas que podem ser carregadas
    If sourceinfo.vQtdLin >= objects.vQtdLinhas Then
        linCharge = objects.vQtdLinhas
    Else
        linCharge = sourceinfo.vQtdLin
        For i = linCharge + 1 To objects.vQtdLinhas
            setVisible i, False
        Next i
    End If
    'Verifica o maximo de colunas que podem ser carregadas
    colCharge = objects.vQtdColunas
    
    If linCharge = 0 Or colCharge = 0 Then Exit Function
    
    If BubSrtMetodo <> "" Then
        BubbleSort objects.vSource(), BubSrtColumn, BubSrtMetodo
    End If
    
    For i = 1 To linCharge
        For j = 1 To colCharge
            If colWithPic = 0 Then
                setValue "r" & i & "c" & j, objects.vSource(i + (firstCharge - 1), j), "string"
                If hasKeyCol Then objects.vKeys(i) = objects.vSource(i, 0)
            Else
                If j = colWithPic Then
                    setValue "r" & i & "c" & j, objects.vSource(i + (firstCharge - 1), j), "picture"
                    If hasKeyCol Then objects.vKeys(i) = objects.vSource(i, 0)
                Else
                    setValue "r" & i & "c" & j, objects.vSource(i + (firstCharge - 1), j), "string"
                    If hasKeyCol Then objects.vKeys(i) = objects.vSource(i, 0)
                End If
            End If
        Next j
    Next i
    
    sourceinfo.vFirstLoaded = CLng(firstCharge)
    objects.vColWithPicture = colWithPic
    attScroll sourceinfo.vQtdLin
End Function

Public Function loadData_inRange(ByVal fromLine As Long)
    If Not objects.vInitialized Then Exit Function
    Dim i As Long, j As Long, k As Long
    'Variavel k para andar nos controles de linhas do listbox
    k = 1
    If fromLine <= sourceinfo.vQtdLin And fromLine >= 1 Then
        For i = fromLine To fromLine + objects.vQtdLinhas - 1
            For j = 1 To sourceinfo.vQtdCol
                If i <= sourceinfo.vQtdLin Then
                    If objects.vColWithPicture = 0 Then
                        setValue "r" & k & "c" & j, objects.vSource(i, j), "string"
                        If hasKeyCol Then objects.vKeys(k) = objects.vSource(i, 0)
                    Else
                        If j = objects.vColWithPicture Then
                            setValue "r" & k & "c" & j, objects.vSource(i, j), "picture"
                            If hasKeyCol Then objects.vKeys(k) = objects.vSource(i, 0)
                        Else
                            setValue "r" & k & "c" & j, objects.vSource(i, j), "string"
                            If hasKeyCol Then objects.vKeys(k) = objects.vSource(i, 0)
                        End If
                    End If
                    setVisible k, True
                Else
                    setValue "r" & k & "c" & j, "", "string"
                    If hasKeyCol Then objects.vKeys(k) = -1
                    setVisible k, False
                End If
            Next j
            k = k + 1
        Next i
    End If
End Function

Private Function BubbleSort(ByRef MyArray() As Variant, ByVal ColumnToCompare As Long, ByVal metodo As String)
    'Sorts a one-dimensional VBA array from smallest to largest
    'using the bubble sort algorithm.
    Dim i As Long, j As Long, k As Long
    Dim Temp As Variant
    
    If metodo = "az" Then
        For i = LBound(MyArray, 1) To UBound(MyArray, 1)
            For j = i + 1 To UBound(MyArray, 1)
                If MyArray(i, ColumnToCompare) > MyArray(j, ColumnToCompare) Then
                    For k = LBound(MyArray, 2) To UBound(MyArray, 2)
                        Temp = MyArray(j, k)
                        MyArray(j, k) = MyArray(i, k)
                        MyArray(i, k) = Temp
                    Next k
                End If
            Next j
        Next i
    End If
    If metodo = "za" Then
        For i = LBound(MyArray, 1) To UBound(MyArray, 1)
            For j = i + 1 To UBound(MyArray, 1)
                If MyArray(i, ColumnToCompare) < MyArray(j, ColumnToCompare) Then
                    For k = LBound(MyArray, 2) To UBound(MyArray, 2)
                        Temp = MyArray(j, k)
                        MyArray(j, k) = MyArray(i, k)
                        MyArray(i, k) = Temp
                    Next k
                End If
            Next j
        Next i
    End If
End Function

Public Function filterBy(ByVal Key As String, ByVal ColumnToCompare As Long)
    
    If Not objects.vInitialized Then Exit Function
    Dim i As Long, j As Long, fRow As Long
    Dim matchs() As Long
    Dim Match As Long
    
    fRow = 1
    Match = 1
    
    Call ResetFilter
    
    If Key = "" Then
        attScroll sourceinfo.vQtdLin
        Exit Function
    End If
    
    For i = LBound(objects.vSource, 1) To UBound(objects.vSource, 1)
        If UCase(objects.vSource(i, ColumnToCompare)) Like CStr("*" & UCase(Key) & "*") Then
            ReDim Preserve matchs(1 To Match) As Long
            matchs(Match) = i
            Match = Match + 1
            fRow = fRow + 1
        End If
    Next i
    
    If fRow > 1 Then
        ReDim objects.vFilter(1 To fRow - 1, LBound(objects.vSource, 2) To sourceinfo.vQtdCol) As Variant
        For i = 1 To fRow - 1
            For j = LBound(objects.vFilter, 2) To UBound(objects.vFilter, 2)
                objects.vFilter(i, j) = objects.vSource(matchs(i), j)
            Next j
        Next i
        objects.vSource = objects.vFilter
        sourceinfo.vQtdLin = UBound(objects.vFilter, 1)
        sourceinfo.vQtdCol = UBound(objects.vFilter, 2)
    Else
        Erase objects.vSource
        sourceinfo.vQtdLin = 0
        sourceinfo.vQtdCol = 0
    End If
    attScroll sourceinfo.vQtdLin
End Function

Public Function ResetFilter()
    If Not IsEmpty(objects.vSource) And objects.vInitialized Then
        Erase objects.vFilter
        objects.vSource = objects.vBackup
        sourceinfo.vQtdLin = UBound(objects.vSource, 1)
        sourceinfo.vQtdCol = UBound(objects.vSource, 2)
    End If
End Function

Public Function getRowDataFromLabel(ByVal row As Long, ByVal col As Long)
    getRowDataFromLabel = vForm.Controls(objects.vPrefixo & "r" & row & "c" & col).Caption
End Function

Private Sub labelAZ_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not labelAZ Is Nothing Then
        update "az", keyCol, picCol
        If Not scrollList Is Nothing Then
            scrollList.value = 1
        End If
    End If
End Sub

Private Sub labelZA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not labelZA Is Nothing Then
        update "za", keyCol, picCol
        If Not scrollList Is Nothing Then
            scrollList.value = 1
        End If
    End If
End Sub

Private Sub txtbFilter_Change()
    If Not txtbFilter Is Nothing Then
        filterBy txtbFilter.value, keyCol
        update "", keyCol, picCol
    End If
End Sub

Private Sub vBackList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clearEffects
End Sub

Private Sub scrollList_Change()
    If Not scrollList Is Nothing Then
        If Not pgSkp Then
            loadData_inRange scrollList.value
        Else
            loadData_inRange ((scrollList.value - 1) * objects.vQtdLinhas) + 1
        End If
    End If
End Sub

Private Function clearEffects()
    Dim i, j As Long
    For i = 1 To objects.vQtdLinhas
        rows(i).restoreTheme
        For j = 1 To objects.vQtdColunas
            cols(i, j).restoreTheme
        Next j
    Next i
End Function

Private Function attScroll(ByVal sourceQtdLin As Long)
    If Not scrollList Is Nothing Then
        scrollList.Min = 1
        If sourceQtdLin > objects.vQtdLinhas Then
            scrollList.Enabled = True
            scrollList.Max = sourceQtdLin - objects.vQtdLinhas + 1
        Else
            scrollList.Max = 1
            scrollList.Enabled = False
        End If
        If pgSkp Then
            If scrollList.Max <> 1 Then
                scrollList.Max = WorksheetFunction.RoundUp((sourceQtdLin / objects.vQtdLinhas), 0)
            End If
        End If
    End If
End Function
