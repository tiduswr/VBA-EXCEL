VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectoryUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_PATH As Integer = 256

Public Function CheckFileExists(ByVal pathString As String) As Boolean
    Dim strFileExists As String
    strFileExists = Dir(pathString)
 
    If strFileExists = "" Then
        CheckFileExists = False
    Else
        CheckFileExists = True
    End If
End Function
Public Function CheckFolderExists(ByVal pathString As String) As Boolean
    Dim strFolderExists As String
    strFolderExists = Dir(pathString, vbDirectory)
 
    If strFolderExists = "" Then
        CheckFolderExists = False
    Else
        CheckFolderExists = True
    End If
End Function
Public Function CreateFolder(ByVal pathString As String) As Boolean
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(pathString, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
    CreateFolder = CheckFolderExists(pathString)
End Function
Public Function GetFileExtension(ByVal pathString As String) As String
    Dim ver As Boolean: ver = False
    Dim aux As String
    Dim char As Integer: char = 0
    
    While Not ver
        char = char + 1
        aux = Right(pathString, char)
        
        If Left(aux, 1) = "." Then ver = True
    Wend
    If ver Then GetFileExtension = aux
End Function
Public Function DeleteFile(ByVal fileToDelete As String) As Boolean
   If CheckFileExists(fileToDelete) Then
      SetAttr fileToDelete, vbNormal 'Retira o somente leitura
      Kill fileToDelete
      DeleteFile = True
      Exit Function
   End If
   DeleteFile = False
End Function
Function IsDirLenghtValid(ByVal Destino As String) As Boolean
    If (MAX_PATH - Len(Destino)) > 0 Then
        DirLength = True
    Else
        DirLength = False
    End If
End Function
Public Function WIA_ResizeImage(sInitialImage As String, sResizedImage As String, _
                                           lMaximumWidth As Long, lMaximumHeight As Long) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : WIA_ResizeImage
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Resize an image based on Max width and Max height using WIA
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'
' Windows Image Acquisition (WIA)
'             https://msdn.microsoft.com/en-us/library/windows/desktop/ms630368(v=vs.85).aspx
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sInitialImage : Fully qualified path and filename of the original image to resize
' sResizedImage : Fully qualified path and filename of where to save the resized image
' lMaximumWidth : Maximum allowable image width
' lMaximumHeight: Maximum allowable image height
'
' Usage:
' ~~~~~~
' Call WIA_ResizeImage("C:\Users\Public\Pictures\Sample Pictures\Chrysanthemum.jpg", _
'                      "C:\Users\MyUser\Desktop\Chrysanthemum_small.jpg", _
'                      800, 600)
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2017-01-18              Initial Release
'---------------------------------------------------------------------------------------
    
    On Error GoTo Error_Handler
    Dim oWIA                  As Object 'WIA.ImageFile
    Dim oIP                   As Object 'ImageProcess
 
    Set oWIA = CreateObject("WIA.ImageFile")
    Set oIP = CreateObject("WIA.ImageProcess")
 
    oIP.Filters.Add oIP.FilterInfos("Scale").FilterID
    oIP.Filters(1).Properties("MaximumWidth") = lMaximumWidth
    oIP.Filters(1).Properties("MaximumHeight") = lMaximumHeight
 
    oWIA.LoadFile sInitialImage
    Set oWIA = oIP.Apply(oWIA)
    oWIA.SaveFile sResizedImage
    WIA_ResizeImage = True
 
Error_Handler_Exit:
    On Error Resume Next
    If Not oIP Is Nothing Then Set oIP = Nothing
    If Not oWIA Is Nothing Then Set oWIA = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: WIA_ResizeImage" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

