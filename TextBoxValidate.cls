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

Private txtbox As MSForms.TextBox
Private myRestrictedChars() As String
Private qtdChars As Integer

Public Sub construtor(ByRef control As MSForms.TextBox, Optional restrictedChars As String = "")
    Set txtbox = control
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
    isADate = isDate(txtbox.Text)
End Function

Public Function isNumber() As Boolean
    isNumber = isNumeric(txtbox.Text)
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
End Function
