VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessDatabaseManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------Facilitador de Manipulçaão de Banco de Dados ACCESS:
'Autor: Harllem Nascimento(TidusWR)
'E-Mail: harllemnascimento@gmail.com
'Data de finalização do projeto: 25/11/2021

Option Explicit
Private CurrentDb As ADODB.Connection
Private DBIsOpen As Boolean
Private rs As ADODB.Recordset

Public Function connect(ByVal pathStr As String, Optional provider As String = "Microsoft.ACE.OLEDB.12.0")
    Set CurrentDb = New ADODB.Connection
    With CurrentDb
        .provider = provider
        .ConnectionString = "Data Source=" & pathStr
        .Open
    End With
    DBIsOpen = True
End Function

Public Function disconnect()
    If DBIsOpen Then
        CurrentDb.Close
        If rs <> Nothing Then rs.Close
        
        Set rs = Nothing
        Set CurrentDb = Nothing
        DBIsOpen = False
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
    
    If Not DBIsOpen Then
        createRecordsetByQuery = False
        Exit Function
    End If
    
    If rs <> Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open query, CurrentDb, adOpenKeyset, adLockOptimistic, adCmdText
    rs.MoveFirst
    
    createRecordsetByQuery = rs.EOF
    
End Function

Public Function getRSCount() As Integer
    
    If Not DBIsOpen Or rs = Nothing Then Exit Function

    On Error GoTo handleError
    rs.MoveLast
    rs.MoveFirst
    If rs.EOF Then Exit Function
    
    getRSCount = rs.RecordCount
    Exit Function
    
handleError:
    getRSCount = 0
    Exit Function

End Function

Public Function getRSColumnCount() As Integer

    If Not DBIsOpen Or rs = Nothing Then Exit Function

    On Error GoTo handleError
    rs.MoveLast
    rs.MoveFirst
    If rs.EOF Then Exit Function
    
    getRSColumnCount = rs.Fields.count
    Exit Function
    
handleError:
    getRSColumnCount = 0
    Exit Function

End Function

Private Function returnQueryString(ByVal table As String, ByVal action As String, ByVal columnName As String, _
                                    ByVal identifier As Variant, Optional fuzzyMatch As Boolean = False) As String
    
    Dim operator As String
    If fuzzyMatch Then
        operator = " LIKE "
    Else
        operator = " = "
    End If
    
    If IsNumeric(identifier) Then
        Select Case True
            Case VarType(identifier) = vbDouble
                returnQueryString = action & " * FROM " & table & " WHERE(" & columnName & operator & _
                        Replace(CStr(CDbl(identifier)), ",", ".") & ")"
            Case VarType(identifier) = vbDecimal
                returnQueryString = action & " * FROM " & table & " WHERE(" & columnName & operator & _
                        Replace(CStr(CDec(identifier)), ",", ".") & ")"
            Case Else
                returnQueryString = action & " * FROM " & table & " WHERE(" & columnName & operator & identifier & ")"
        End Select
    Else
        If IsDate(identifier) Then
            returnQueryString = action & " * FROM " & table & " WHERE(" & columnName & operator & "#" & CStr(identifier) & "#)"
        Else
            returnQueryString = action & " * FROM " & table & " WHERE(" & columnName & operator & "'" & CStr(identifier) & "')"
        End If
    End If
    
End Function

''As funções abaixo foram adptadas do seguinte repositório:
''
''@WyattCast44 -> On Github
''https://github.com/wyattcast44-side-project-graveyard/MS-Access-VBA-Libraries/tree/master/libraries/src/sql

Public Function insert(table As String, columns As String, ParamArray values() As Variant) As Variant

On Error GoTo handleError
    
    Dim insertQry, value, valueStr, columnCount, argumentCount As Variant
    Dim aux As String
    
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
                valueStr = valueStr & Replace(CStr(CDbl(value)), ",", ".") & ","
            Case VarType(value) = vbSingle
                valueStr = valueStr & value & ","
            Case VarType(value) = vbDecimal
                valueStr = valueStr & Replace(CStr(CDec(value)), ",", ".") & ","
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

Public Function delete(ByVal table As String, ByVal identifier As Variant, ByVal columnName As String) As Boolean

On Error GoTo handleError

    Dim qry As String
    
    If columnName <> "" Then
        qry = returnQueryString(table, "DELETE", columnName, identifier)
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

Public Function tableContain(table As String, searchCriteria As Variant, searchColumn As String, _
                                    Optional fuzzyMatch As Boolean = False) As Boolean
    
    On Error GoTo handleError
    
    Dim selectQry As String: selectQry = returnQueryString(table, "SELECT", searchColumn, searchCriteria, fuzzyMatch)
    Dim auxRs As ADODB.Recordset
    
    auxRs.Open selectQry, CurrentDb, adOpenKeyset, adLockOptimistic, adCmdText
    auxRs.MoveFirst
    
    tableContain = (Not auxRs.EOF) And (auxRs.RecordCount <> 0)
    
    If auxRs <> Nothing Then auxRs.Close
    Set auxRs = Nothing
    Exit Function
    
handleError:
    If auxRs <> Nothing Then auxRs.Close
    Set auxRs = Nothing
    tableContain = False
End Function
