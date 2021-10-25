VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private rotulo As CustomLabel
Public txtboxValidate As TextBoxValidate
Private txtbox As MSForms.TextBox
Private cbbox As MSForms.ComboBox

Public Function construtor(ByRef vRotulo As MSForms.label, ByRef vLine As MSForms.label, _
                            ByRef control As MSForms.control, ByVal backgroundColor As Long, _
                            ByVal captionColor As Long, ByVal captionBold As Boolean, _
                            Optional backStyle As Integer = -1, Optional borderStyle As Integer = -1, _
                            Optional borderColor As Long = -1, Optional restChars As String = "")

    Set rotulo = New CustomLabel
    rotulo.construtor vRotulo, backgroundColor, captionColor, captionBold, backStyle, borderStyle, borderColor, _
                        False, False
    ''Aplica o tema de animação
    If backStyle <> -1 Then
        vLine.backStyle = backStyle
    End If
    If borderStyle <> -1 Then
        vLine.borderStyle = borderStyle
    End If
    If borderColor <> -1 Then
        vLine.borderColor = borderColor
    End If
    vLine.Font.Bold = captionBold
    vLine.ForeColor = captionColor
    vLine.backColor = backgroundColor
    
    If TypeName(control) = "TextBox" Then
        Set txtbox = control
        Set txtboxValidate = New TextBoxValidate
        txtboxValidate.construtor txtbox, restChars
        
        txtbox.borderStyle = fmBorderStyleNone
        txtbox.backStyle = fmBackStyleTransparent
    Else
        If TypeName(control) = "ComboBox" Then
            Set cbbox = control
            Set txtbox = Nothing
            
            cbbox.backStyle = fmBackStyleTransparent
            cbbox.borderStyle = fmBorderStyleNone
        End If
    End If
End Function
