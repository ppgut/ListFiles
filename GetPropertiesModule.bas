Attribute VB_Name = "GetPropertiesModule"
Option Explicit

Type FolderAttributes
    Name As String
    Size As String
    DateCreated As Date
    DateAccessed As Date
    DateModified As Date
    Path As String
End Type

Type FileAttributes
    Name As String
    Size As String
    FileType As String
    DateModified As Date
    DateCreated As Date
    DateAccessed As Date
    Attributes As String
    Status As String
    Owner As String
    Author As String
    Title As String
    Subject As String
    Category As String
End Type

Sub GetProperties()

    If ActiveCell.Column > 24 Then Exit Sub
    
    Dim FileAttr As FileAttributes
    Dim FolderAttr As FolderAttributes
    Dim sProps As String
    Dim Path As String
    Dim PathCol As Long
    
    PathCol = ActiveSheet.Range("A1").Resize(1, 100).Find("Path").Column
    Path = Sheet1.Cells(ActiveCell.Row, PathCol).Value
    If Dir(Path, vbDirectory) = "" Then Exit Sub
    
    If ActiveCell.EntireRow.Resize(1, 1).End(xlToRight).Font.Color = rgbCrimson Then
        If Dir(Path) = "" Then Exit Sub
        
        FileAttr = GetFileAttributes(Path)
        
        With FileAttr
            sProps = sProps & "Name:" & vbTab & vbTab & .Name & vbNewLine
            sProps = sProps & "Size:" & vbTab & vbTab & .Size & vbNewLine
            sProps = sProps & "FileType:" & vbTab & vbTab & .FileType & vbNewLine
            sProps = sProps & "DateModified:" & vbTab & .DateModified & vbNewLine
            sProps = sProps & "DateCreated:" & vbTab & .DateCreated & vbNewLine
            sProps = sProps & "DateAccessed:" & vbTab & .DateAccessed & vbNewLine
            sProps = sProps & "Attributes:" & vbTab & .Attributes & vbNewLine
            sProps = sProps & "Status:" & vbTab & vbTab & .Status & vbNewLine
            sProps = sProps & "Owner:" & vbTab & vbTab & .Owner & vbNewLine
            sProps = sProps & "Author:" & vbTab & vbTab & .Author & vbNewLine
            sProps = sProps & "Title:" & vbTab & vbTab & .Title & vbNewLine
            sProps = sProps & "Subject:" & vbTab & vbTab & .Subject & vbNewLine
            sProps = sProps & "Category:" & vbTab & vbTab & .Category
        End With
    Else
        If Dir(Path, vbDirectory) = "" Then Exit Sub
        FolderAttr = GetFolderAttributes(Path)
        
        With FolderAttr
            sProps = sProps & "Name:" & vbTab & vbTab & .Name & vbNewLine
            sProps = sProps & "Size:" & vbTab & vbTab & .Size & " bytes" & vbNewLine
            sProps = sProps & "DateModified:" & vbTab & .DateModified & vbNewLine
            sProps = sProps & "DateCreated:" & vbTab & .DateCreated & vbNewLine
            sProps = sProps & "DateAccessed:" & vbTab & .DateAccessed & vbNewLine
            sProps = sProps & "Path:" & vbNewLine & .Path
        End With
    End If
    
    MsgBox sProps
    
End Sub

Function GetFolderAttributes(Path As String) As FolderAttributes
    
    Dim oFSO            As Object
    Dim oFolder         As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(Path)

    With oFolder
        GetFolderAttributes.Name = .Name
        GetFolderAttributes.Size = .Size
        GetFolderAttributes.DateCreated = .DateCreated
        GetFolderAttributes.DateAccessed = .DateLastAccessed
        GetFolderAttributes.DateModified = .DateLastModified
        GetFolderAttributes.Path = .Path
    End With

    Set oFolder = Nothing
    Set oFSO = Nothing
    
End Function

Function GetFileAttributes(Path As String) As FileAttributes

' Shell32 objects
Dim objShell
Dim objFolder As Object
Dim objFolderItem As Object

' Other objects
Dim strPath
Dim strFileName
Dim i As Integer

    ' If the file does not exist then quit out
    If Dir(Path) = "" Then Exit Function

    ' Parse the file name out from the folder path
    i = InStrRev(Path, "\")
    strFileName = Mid$(Path, i + 1)
    strPath = Left$(Path, i - 1)
    
    ' Set up the shell32 Shell object
    Set objShell = CreateObject("Shell.Application")

    ' Set the shell32 folder object
    Set objFolder = objShell.Namespace(strPath)

    ' If we can find the folder then ...
    If (Not objFolder Is Nothing) Then

        ' Set the shell32 file object
        Set objFolderItem = objFolder.ParseName(strFileName)

        ' If we can find the file then get the file attributes
        If (Not objFolderItem Is Nothing) Then

            GetFileAttributes.Name = objFolder.GetDetailsOf(objFolderItem, 0)
            GetFileAttributes.Size = objFolder.GetDetailsOf(objFolderItem, 1)
            GetFileAttributes.FileType = objFolder.GetDetailsOf(objFolderItem, 2)
            GetFileAttributes.DateModified = CDate(objFolder.GetDetailsOf(objFolderItem, 3))
            GetFileAttributes.DateCreated = CDate(objFolder.GetDetailsOf(objFolderItem, 4))
            GetFileAttributes.DateAccessed = CDate(objFolder.GetDetailsOf(objFolderItem, 5))
            GetFileAttributes.Attributes = objFolder.GetDetailsOf(objFolderItem, 6)
            GetFileAttributes.Status = objFolder.GetDetailsOf(objFolderItem, 7)
            GetFileAttributes.Owner = objFolder.GetDetailsOf(objFolderItem, 10)
            GetFileAttributes.Author = objFolder.GetDetailsOf(objFolderItem, 20)
            GetFileAttributes.Title = objFolder.GetDetailsOf(objFolderItem, 21)
            GetFileAttributes.Subject = objFolder.GetDetailsOf(objFolderItem, 22)
            GetFileAttributes.Category = objFolder.GetDetailsOf(objFolderItem, 23)

        End If

        Set objFolderItem = Nothing

    End If

    Set objFolder = Nothing
    Set objShell = Nothing

End Function
