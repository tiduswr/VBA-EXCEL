VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomDatabaseAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------Facilitador de Manipulçaão de Banco de Dados ACCESS:
'Autor: Harllem Nascimento(TidusWR)
'E-Mail: harllemnascimento@gmail.com
'Data de finalização do projeto: 08/11/2021

Option Explicit
Private CurrentDb As New ADODB.Connection
Private DBIsOpen As Boolean
Private rs As New ADODB.Recordset

Public Function conect(ByVal pathStr As String, Optional provider As String = "Microsoft.ACE.OLEDB.12.0")
    Set CurrentDb = New ADODB.Connection
    With CurrentDb
        .provider = provider
        .ConnectionString = "Data Source=" & pathStr
        .Open
    End With
    DBIsOpen = True
End Function

Public Function disconect()
    If DBIsOpen Then
        CurrentDb.Close
        If rs <> Nothing Then rs.Close
        
        Set rs = Nothing
        Set CurrentDb = Nothing
    End If
End Function

Public Function getRecordSet() As ADODB.Recordset
    getRecordSet = rs
End Function

Public Function getRecordSetAsArray() As Variant()
    
    If rs = Nothing Then Exit Function
    
    Dim i, j As Integer
    Dim tblArray() As Variant: ReDim tblArray(0 To rs.RecordCount - 1, 0 To rs.Fields - 1) As Variant
    
    With rs
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            For j = 1 To .Fields.count
                 tblArray(i - 1, j - 1) = .Fields(j).value
            Next j
            .MoveNext
        Next i
    End With
    
    getRecordSetAsArray = tblArray
    
End Function

Public Function createRecordsetByQuery(ByVal query As String) As Boolean

    Set rs = New ADODB.Recordset
    rs.Open query, CurrentDb, adOpenKeyset, adLockOptimistic, adCmdText
    rs.MoveFirst
    
    createRecordsetByQuery = rs.EOF
    
End Function

Public Function getRSCount() As Integer

    If rs = Nothing Then Exit Function

    On Error GoTo handleError
    rs.MoveLast
    rs.MoveFirst
    If rs.EOF Then Exit Function
    
    getRSCount = rs.RecordCount
    Exit Function
    
handleError:
    getRSCount = rs.RecordCount
    Exit Function

End Function

Public Function getRSColumnCount() As Integer

    If rs = Nothing Then Exit Function

    On Error GoTo handleError
    rs.MoveLast
    rs.MoveFirst
    If rs.EOF Then Exit Function
    
    getRSColumnCount = rs.Fields.count
    Exit Function
    
handleError:
    getRSColumnCount = -1
    Exit Function

End Function

''As funções abaixo foram pegas e adptadas do seguinte repositório:
''
''@WyattCast44 -> On Github
''https://github.com/wyattcast44-side-project-graveyard/MS-Access-VBA-Libraries/tree/master/libraries/src/sql

Public Function insert(table As String, columns As String, ParamArray values() As Variant) As Variant

On Error GoTo handleError
    
    Dim insertQry, value, valueStr, columnCount, argumentCount As Variant

    valueStr = ""
    argumentCount = (UBound(values) - LBound(values) + 1)
    columnCount = (UBound(Split(columns, ",")) - LBound(Split(columns, ",")) + 1)
    
    If argumentCount <> columnCount Then GoTo handleError
    
    For Each value In values
        
        Select Case True
            Case VarType(value) = vbInteger
                valueStr = valueStr & value & ","
            Case VarType(value) = vbLong
                valueStr = valueStr & value & ","
            Case VarType(value) = vbDouble
                valueStr = valueStr & value & ","
            Case VarType(value) = vbSingle
                valueStr = valueStr & value & ","
            Case VarType(value) = vbDecimal
                valueStr = valueStr & value & ","
            Case VarType(value) = vbString
                valueStr = valueStr & "'" & value & "',"
            Case VarType(value) = vbDate
                valueStr = valueStr & "#" & value & "#,"
            Case VarType(value) = vbBoolean
                valueStr = valueStr & value & ","
            Case Else
                valueStr = valueStr & "'" & CStr(value) & "',"
        End Select
        
    Next
    
    valueStr = IIf(Right(valueStr, 1) = ",", Left(valueStr, (Len(valueStr) - 1)), valueStr)
    
    insertQry = "INSERT INTO " & table & "(" & CStr(columns) & ") VALUES(" & valueStr & ")"
    
    CurrentDb.Execute insertQry, dbFailOnError
    
    insert = True
    
    Exit Function
    
handleError:
    insert = False
    Exit Function

End Function

Public Function delete(table As String, identifier As Variant, columnName As String) As Boolean

On Error GoTo handleError

    Dim qry As String
    
    If columnName <> "" Then
        
        If IsNumeric(identifier) Then
            qry = "DELETE * FROM " & table & " WHERE(" & columnName & " = " & identifier & ")"
        Else
            qry = "DELETE * FROM " & table & " WHERE(" & columnName & " = '" & CStr(identifier) & "')"
        End If
        
    Else
        delete = False
        Exit Function
    End If
    
    CurrentDb.Execute qry, dbFailOnError
    delete = True
    Exit Function
    
handleError:
    delete = False
    Exit Function
    
End Function

Public Function doesTableContain(table As String, searchCriteria As String, searchColumn As String, _
    Optional fuzzyMatch As Boolean = True) As Boolean
    
    Dim selectQry As String
    Dim rs As Recordset
    
    If fuzzyMatch Then
        selectQry = "SELECT * FROM " & table & " WHERE " & searchColumn & " LIKE '*" & searchCriteria & "*'"
    Else
        selectQry = "SELECT * FROM " & table & " WHERE " & searchColumn & " = '" & searchCriteria & "'"
    End If
    
    Set rs = CurrentDb.OpenRecordset(selectQry)
    
    If getRSCount(rs) <> 0 Then
        doesTableContain = True
    Else
        doesTableContain = False
    End If
    
    Set rs = Nothing
    
End Function
