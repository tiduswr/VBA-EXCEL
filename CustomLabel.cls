VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type themeLabel
    backgroundColor As Long
    captionColor As Long
    captionBold As Boolean
    mousePassSpecialEffect As Integer
    backStyle As Integer
    borderStyle As Integer
    borderColor As Long
End Type

Private WithEvents animation As MSForms.label
Attribute animation.VB_VarHelpID = -1
Private theme As themeLabel
Private backupTheme As themeLabel
Private selected As Boolean
Private slctOpt As Boolean
Private controlName As String
Private myParentLabels As Collection
Private animatedOn As Boolean
Private redirect As MSForms.label
Private clickSunken As Boolean

Public Function construtor(ByRef label As MSForms.label, ByVal backgroundColor As Long, _
                            ByVal captionColor As Long, ByVal captionBold As Boolean, _
                            Optional backStyle As Integer = -1, Optional borderStyle As Integer = -1, _
                            Optional borderColor As Long = -1, Optional selectOption As Boolean = False, _
                            Optional animated As Boolean = True, Optional ByRef redirectAnimation As MSForms.label = Nothing, _
                            Optional ByVal clickSpecialEffectSunken As Boolean = True)
    ''Coloca os valores da propriedades
    Set animation = label
    slctOpt = selectOption
    selected = False
    Set myParentLabels = Nothing
    animatedOn = animated
    clickSunken = clickSpecialEffectSunken
    
    ''Verifica se vai redirecionar a animação
    Set redirect = redirectAnimation
    
    ''Pega o nome do controle
    Dim x As control: Set x = animation
    controlName = x.Name
    Set x = Nothing
    
    ''Faz backup do tema base
    If redirectAnimation Is Nothing Then
        backupTheme.captionBold = animation.Font.Bold
        backupTheme.captionColor = animation.ForeColor
        backupTheme.backgroundColor = animation.backColor
        backupTheme.mousePassSpecialEffect = animation.SpecialEffect
        backupTheme.backStyle = animation.backStyle
        backupTheme.borderStyle = animation.borderStyle
        backupTheme.borderColor = animation.borderColor
    Else
        backupTheme.captionBold = redirectAnimation.Font.Bold
        backupTheme.captionColor = redirectAnimation.ForeColor
        backupTheme.backgroundColor = redirectAnimation.backColor
        backupTheme.mousePassSpecialEffect = redirectAnimation.SpecialEffect
        backupTheme.backStyle = redirectAnimation.backStyle
        backupTheme.borderStyle = redirectAnimation.borderStyle
        backupTheme.borderColor = redirectAnimation.borderColor
    End If
    ''Aplica o tema de animação
    If backStyle <> -1 Then
        theme.backStyle = backStyle
    Else
        If redirectAnimation Is Nothing Then
            theme.backStyle = animation.backStyle
        Else
            theme.backStyle = redirectAnimation.backStyle
        End If
    End If
    If borderStyle <> -1 Then
        theme.borderStyle = borderStyle
    Else
        If redirectAnimation Is Nothing Then
            theme.borderStyle = animation.borderStyle
        Else
            theme.borderStyle = redirectAnimation.borderStyle
        End If
    End If
    If borderColor <> -1 Then
        If redirectAnimation Is Nothing Then
            theme.borderColor = animation.borderColor
        Else
            theme.borderColor = redirectAnimation.borderColor
        End If
    Else
        If redirectAnimation Is Nothing Then
            theme.borderColor = animation.borderColor
        Else
            theme.borderColor = redirectAnimation.borderColor
        End If
    End If
    theme.captionBold = captionBold
    theme.captionColor = captionColor
    theme.backgroundColor = backgroundColor
    
    ''Verifica se é uma label para aplicar somente o tema
    If Not animated Then aplyTheme
    
End Function

Public Function setParentsCollection(ByRef parentLabels As Collection)
    Set myParentLabels = parentLabels
End Function

Public Function getControlName()
    getControlName = controlName
End Function

Public Function isSelected() As Boolean
    isSelected = selected
End Function

Public Function setSelected(ByVal state As Boolean) As Boolean
    selected = state
End Function

''Animação de Clique
Private Sub animation_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not animatedOn Then Exit Sub
    
    ''Opção para habilitar o redirecionamento de animações
    Dim objec As MSForms.label
    If Not redirect Is Nothing Then
        Set objec = redirect
    Else
        Set objec = animation
    End If
    
    If clickSunken Then objec.SpecialEffect = fmSpecialEffectSunken
    If slctOpt Then
        If Not selected Then
            selected = True
            If Not myParentLabels Is Nothing Then
                Dim obj As CustomLabel
                
                For Each obj In myParentLabels
                    If obj.getControlName() <> getControlName() Then
                        If obj.isSelected() Then
                            obj.setSelected (False)
                            obj.restoreTheme
                        End If
                    End If
                Next obj
                
            End If
        Else
            selected = False
        End If
    End If
End Sub
Private Sub animation_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not animatedOn Then Exit Sub
    ''Opção para habilitar o redirecionamento de animações
    Dim objec As MSForms.label
    If Not redirect Is Nothing Then
        Set objec = redirect
    Else
        Set objec = animation
    End If
    
    objec.borderStyle = backupTheme.borderStyle
End Sub

''Animação de Passar o mouse
Private Sub animation_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Not animatedOn Then Exit Sub
    
    ''Opção para habilitar o redirecionamento de animações
    Dim objec As MSForms.label
    If Not redirect Is Nothing Then
        Set objec = redirect
    Else
        Set objec = animation
    End If
    
    aplyTheme
    objec.backStyle = fmBackStyleOpaque
End Sub
Private Function aplyTheme()
    ''Opção para habilitar o redirecionamento de animações
    Dim objec As MSForms.label
    If Not redirect Is Nothing Then
        Set objec = redirect
    Else
        Set objec = animation
    End If
    
    objec.SpecialEffect = fmSpecialEffectFlat
    objec.backStyle = fmBackStyleOpaque
    objec.ForeColor = theme.captionColor
    objec.Font.Bold = theme.captionBold
    objec.backColor = theme.backgroundColor
    objec.borderColor = theme.borderColor
    objec.borderStyle = theme.borderStyle
    objec.backStyle = theme.backStyle
End Function

''Restaurar o tema da Label
Public Function restoreTheme() As Boolean
    If Not animatedOn Then Exit Function

    ''Opção para habilitar o redirecionamento de animações
    Dim objec As MSForms.label
    If Not redirect Is Nothing Then
        Set objec = redirect
    Else
        Set objec = animation
    End If

    If Not selected Or Not slctOpt Then
        objec.SpecialEffect = backupTheme.mousePassSpecialEffect
        objec.backStyle = backupTheme.backStyle
        objec.borderStyle = backupTheme.borderStyle
        objec.borderColor = backupTheme.borderColor
        objec.ForeColor = backupTheme.captionColor
        objec.Font.Bold = backupTheme.captionBold
        objec.backColor = backupTheme.backgroundColor
        objec.backStyle = backupTheme.backStyle
        restoreTheme = True
    Else
        restoreTheme = False
    End If
End Function
