VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents txtbox As MSForms.TextBox
Private myRestrictedChars() As String
Private qtdChars As Integer
Private tpIn As String

''typeInput -> none, data, restChars, number, cpf
Public Sub construtor(ByRef control As MSForms.TextBox, Optional restrictedChars As String = "", _
                        Optional typeInput As String = "none")
    Dim i As Integer
    Set txtbox = control
    tpIn = typeInput
    
    ''Regras de tipos de formatação
    If tpIn = "cpf" Then control.MaxLength = 14
    If tpIn = "data" Then control.MaxLength = 10
    
    If restrictedChars <> "" Then
        qtdChars = Len(restrictedChars)
        ReDim myRestrictedChars(1 To qtdChars) As String
        For i = 1 To qtdChars
            myRestrictedChars(i) = Mid(restrictedChars, i, 1)
        Next i
    Else
        qtdChars = 0
    End If
End Sub

Public Function setRestrictedChars(ByVal restrictedChars As String)
    Dim i As Integer
    
    If restrictedChars <> "" Then
        qtdChars = Len(restrictedChars)
        ReDim myRestrictedChars(1 To qtdChars) As String
        For i = 1 To qtdChars
            myRestrictedChars(i) = Mid(restrictedChars, i, 1)
        Next i
    Else
        qtdChars = 0
    End If
End Function

Public Function isADate() As Boolean
    isADate = IsDate(txtbox.Text)
End Function

Public Function isNumber() As Boolean
    isNumber = IsNumeric(txtbox.Text)
End Function

Public Function containsRestrictedChar() As Boolean
    Dim i As Integer
    If Not IsEmpty(myRestrictedChars) Then
        For i = 1 To qtdChars
            If InStr(1, txtbox.Text, myRestrictedChars(i)) <> 0 Then
                containsRestrictedChar = True
                Exit Function
            End If
        Next i
    End If
    containsRestrictedChar = False
End Function

''Funções de controles para os eventos KeyPress do TextBox
Private Function cpfKeyPress(ByRef l As MSForms.ReturnInteger, ByRef txtbox As MSForms.TextBox)
    Select Case l
        Case 8, 48 To 57
            If Len(txtbox) = 3 Then txtbox = txtbox + "."
            If Len(txtbox) = 7 Then txtbox = txtbox + "."
            If Len(txtbox) = 11 Then txtbox = txtbox + "-"
    Case Else
        l = 0
    End Select
End Function

Private Function dataKeyPress(ByRef l As MSForms.ReturnInteger, ByRef txtbox As MSForms.TextBox)
    Select Case l
        Case Asc("0") To Asc("9")
            If Len(txtbox) = 2 Or Len(txtbox) = 5 Then txtbox = txtbox + "/"
        Case Else
            l = 0
    End Select
End Function

Private Function numberKeyPress(ByRef l As MSForms.ReturnInteger)
    Select Case l
        Case Asc("0") To Asc("9")
            l = l
        Case Else
            l = 0
    End Select
End Function

Private Function restCharsKeyPress(ByRef l As MSForms.ReturnInteger)
    If containsRestrictedChar Then l = 0
End Function

''typeInput -> none, data, restChars, number, cpf
Private Sub txtbox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case tpIn
        Case "data"
            dataKeyPress KeyAscii, txtbox
        Case "restChars"
            restCharsKeyPress KeyAscii
        Case "cpf"
            cpfKeyPress KeyAscii, txtbox
        Case "number"
            numberKeyPress KeyAscii
        Case Else
            KeyAscii = KeyAscii
    End Select
End Sub
