Attribute VB_Name = "ListFilesModule"
Private iFilesCount As Integer
Private iFirstFunctionalColumn As Integer
Private bShowProcessBar As Boolean

Type TreeParams
   MaxRecurencyLevel As Long
   FilesCount As Long
End Type

Sub ListFiles()
    
    Dim diaFolder       As FileDialog
    Dim sPath           As String
    Dim oFSO            As Object
    Dim oFolder         As Object
    Dim shp             As Shape
    Dim tstart          As Single
    
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    If Not diaFolder.Show Then Exit Sub
    sPath = diaFolder.SelectedItems(1)

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder(sPath)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    iFilesCount = GetTreeParams(oFolder).FilesCount
    If iFilesCount > 300 Then bShowProcessBar = True
    
    iFirstFunctionalColumn = GetTreeParams(oFolder).MaxRecurencyLevel + 3
    If iFirstFunctionalColumn < 20 Then iFirstFunctionalColumn = 20
    
    Call ClearSheet
    
    Range("a1").Offset(0, iFirstFunctionalColumn - 2).Resize(1, 2).NumberFormat = ";;;"
    Range("a1").Offset(0, iFirstFunctionalColumn - 1).Value = sPath
    Range("a1").Offset(0, iFirstFunctionalColumn - 2).Value = iFirstFunctionalColumn
    
    With ThisWorkbook.Worksheets("Sheet1").Range("A1")
        .Offset(0, iFirstFunctionalColumn).Value = "Path"
        .Offset(0, iFirstFunctionalColumn + 1).Value = "Ext"
        .Offset(0, iFirstFunctionalColumn + 3).Value = "Rename"
        .Offset(0, iFirstFunctionalColumn + 3).EntireColumn.NumberFormat = "General"
        .Resize(1, iFirstFunctionalColumn + 4).Font.Bold = True
        .Value = CStr(oFolder.Name)
        If .Parent.cbxHyperlinks Then
            .Parent.Hyperlinks.Add _
                Anchor:=.Parent.Range("A1"), _
                Address:=oFolder.Path
            .Font.Underline = False
        End If
        .Font.Color = rgbDodgerBlue
    End With
    
    'tstart = Time * 100000
    If bShowProcessBar Then ProcessBar.Show 0
    
    Call ListFilesWithRecurence(oFolder, 1, 1)
    ThisWorkbook.Worksheets("Sheet1").Columns(iFirstFunctionalColumn + 2).AutoFilter
    ProcessBar.Hide
    '[ah2] = CSng(Time) * 100000 - tstart
        
    ActiveCell.Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    bShowProcessBar = False
    Set diaFolder = Nothing
    Set oFSO = Nothing
    Set oFolder = Nothing
    
End Sub

Private Sub ListFilesWithRecurence(ByVal oParentFolder As Object, ByRef i As Integer, ByVal j As Integer)

    Dim oFile As Object
    Dim oFolder As Object
    Dim ParentFolderLine As Integer
    Dim ParentFolderColumn As Integer
    Dim LastSubFolderLine As Integer
    Dim LastFileLine As Integer
    Dim LastLine As Integer
    
    ParentFolderLine = i - 1
    ParentFolderColumn = j - 1
    
    With ThisWorkbook.Worksheets("Sheet1").Range("A1")
        
        For Each oFolder In oParentFolder.SubFolders
        
            Call WriteFolderName(oFolder, .Offset(i, j))
            
            Call AddHorizontalConnector(.Offset(i, j))
            LastSubFolderLine = i
            i = i + 1
            
            If bShowProcessBar Then
                ProcessBar.Text.Caption = Format(i / iFilesCount, "0%") & " completed [Break to stop]"
                ProcessBar.Bar.Width = i / iFilesCount * 200
                If i Mod iFilesCount / 50 = 0 Then DoEvents 'to increase performance of procedure call DoEvents intermittently
            End If
            
            If .Parent.cbxSubfolders Then
                Call ListFilesWithRecurence(oFolder, i, j + 1)
            End If
        Next

        For Each oFile In oParentFolder.Files
            If oFile.Name <> "Thumbs.db" Then
            
                Call WriteFileName(oFile, .Offset(i, j))
                
                Call AddHorizontalConnector(.Offset(i, j))
                LastFileLine = i
                i = i + 1
                
                If bShowProcessBar Then
                    ProcessBar.Text.Caption = Format(i / iFilesCount, "0%") & " completed [Break to stop]"
                    ProcessBar.Bar.Width = i / iFilesCount * 200
                    If i Mod iFilesCount / 50 = 0 Then DoEvents 'to increase performance of procedure call DoEvents intermittently
                End If
                
            End If
        Next
        
        LastLine = Application.Max(LastSubFolderLine, LastFileLine)
        If LastLine > 0 Then
            Call AddVerticalConnector(.Offset(ParentFolderLine, ParentFolderColumn), .Offset(LastLine, ParentFolderColumn))
        End If
        
    End With
    
    Set oFile = Nothing
    Set oFolder = Nothing
    
End Sub

Function GetTreeParams(ByVal oParentFolder As Object, Optional ByVal RecLvl As Long = 0, Optional ByRef MaxRecLvl As Long = 0, Optional ByRef FilesCount As Long = 0) As TreeParams

    If RecLvl > MaxRecLvl Then MaxRecLvl = RecLvl
    
    Dim oFolder As Object

    For Each oFolder In oParentFolder.SubFolders
        Call GetTreeParams(oFolder, RecLvl + 1, MaxRecLvl, FilesCount)
    Next
    
    FilesCount = FilesCount + oParentFolder.Files.Count + 1
    
    GetTreeParams.MaxRecurencyLevel = MaxRecLvl
    GetTreeParams.FilesCount = FilesCount
    
    Set oFolder = Nothing
    
End Function

Sub WriteFileName(oFil As Object, ByVal NameRng As Range)

    iFirstFunctionalColumn = Range("A1").End(xlToRight).Value
    
    With NameRng
        .NumberFormat = "@"
        .Value = CStr(oFil.Name)
        If .Parent.cbxHyperlinks Then
            .Parent.Hyperlinks.Add _
                Anchor:=NameRng, _
                Address:=oFil.Path
        End If
        .Font.Color = rgbCrimson
        .Font.Underline = False
        
        .Offset(0, iFirstFunctionalColumn - .Column + 1).Value = oFil.Path
        .Offset(0, iFirstFunctionalColumn - .Column + 1).Font.Color = rgbGainsboro
        .Offset(0, iFirstFunctionalColumn - .Column + 2) = Extension(oFil.Name)
        .Offset(0, iFirstFunctionalColumn - .Column + 2).Font.Color = rgbForestGreen
        
    End With
    
End Sub

Sub WriteFolderName(oFol As Object, ByVal NameRng As Range)
    
    iFirstFunctionalColumn = Range("A1").End(xlToRight).Value
    
    With NameRng
        .NumberFormat = "@"
        .Value = CStr(oFol.Name)
        If .Parent.cbxHyperlinks Then
            .Parent.Hyperlinks.Add _
                Anchor:=NameRng, _
                Address:=oFol.Path
        End If
        .Font.Color = rgbDodgerBlue
        .Font.Underline = False
        
        .Offset(0, iFirstFunctionalColumn - .Column + 1).Value = oFol.Path
        .Offset(0, iFirstFunctionalColumn - .Column + 1).Font.Color = rgbGainsboro
        .Offset(0, iFirstFunctionalColumn - .Column + 2) = " "
    End With
    
End Sub

Public Function Extension(sFileName As String) As String

    Extension = " "

    If InStrRev(sFileName, ".") > 0 And Len(sFileName) > InStrRev(sFileName, ".") Then
        Extension = Right(sFileName, Len(sFileName) - InStrRev(sFileName, ".") + 1)
    End If

End Function

Public Sub ClearSheet()

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("Sheet1")
    Set ws2 = ThisWorkbook.Worksheets("Sheet2")
    Dim i As Long
    
    ws1.AutoFilterMode = False
    ws1.Range("A1").Copy
    ws2.Range("A1").PasteSpecial xlPasteFormats
    ws1.Cells.Clear
    ws2.Range("A1").Copy
    ws1.Range("A1").PasteSpecial xlPasteFormats
    ws2.Cells.Clear
    ws1.Range("A1").FormatConditions(1).ModifyAppliesToRange Range("A:AD")
    
    With ws1
        For Each shp In .Shapes
            If shp.Name <> "ListFilesButton" And _
                shp.Name <> "cbxSubfolders" And _
                shp.Name <> "cbxHyperlinks" And _
                shp.Name <> "RenameButton" And _
                shp.Name <> "cbxLines" Then shp.Delete
        Next
        Range(.Columns("BA"), .Columns("BA").End(xlToRight)).EntireColumn.Hidden = True
        
        .Columns("A:AZ").ColumnWidth = 5
        .Columns(iFirstFunctionalColumn + 4).ColumnWidth = 8
        Range(.Columns(1), .Columns(iFirstFunctionalColumn)).ColumnWidth = 2
        
        Range(.Columns(1), .Columns(iFirstFunctionalColumn)).Borders.Color = vbWhite
        
        '.Rows(1).Find("Rename").EntireColumn.NumberFormat = "General"
        
    End With
    
End Sub

Private Sub AddHorizontalConnector(rng As Range)

    If rng.Column < 2 Then Exit Sub
    ThisWorkbook.Worksheets("Sheet1").Shapes.AddConnector(msoConnectorStraight, _
    BeginX:=rng.Left - rng.Offset(0, -1).Width / 2, _
    BeginY:=rng.Top + rng.Height / 2, _
    EndX:=rng.Left, _
    EndY:=rng.Top + rng.Height / 2).Select
    
End Sub

Sub AddVerticalConnector(rngStart As Range, rngEnd As Range)

    If (rngEnd.Top + rngEnd.Height / 2) - (rngStart.Top + rngStart.Height) > 169000 Then 'metoda wysypuje sie przy zbyt dlugiej linii
        Dim rngEndInt As Range
        Set rngEndInt = rngStart.Offset(Application.RoundDown((rngEnd.Row - rngStart.Row) / 2, 0), 0)
        Call AddVerticalConnector(rngStart, rngEndInt)
        Call AddVerticalConnector(rngEndInt.Offset(-1, 0), rngEnd)
    Else
        ThisWorkbook.Worksheets("Sheet1").Shapes.AddConnector(msoConnectorStraight, _
        BeginX:=rngStart.Left + rngStart.Width / 2, _
        BeginY:=rngStart.Top + rngStart.Height, _
        EndX:=rngEnd.Left + rngEnd.Width / 2, _
        EndY:=rngEnd.Top + rngEnd.Height / 2).Select
    End If
    
End Sub

Sub Rename()
    
    Dim OldPath As String
    Dim NewPath As String
    Dim OldPathShort As String
    Dim NewPathShort As String
    Dim rng As Range
    Dim rng2 As Range
    Dim RenameCol As Range
    Set RenameCol = Worksheets("Sheet1").Rows(1).Find("Rename")
    Dim ErrorExist As Boolean
    
    Dim oFSO            As Object
    Dim oFile           As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Application.EnableEvents = False
    
    If Application.Subtotal(3, RenameCol.EntireColumn) < 2 Then
        MsgBox ("Enter new name to 'Rename' column")
        Exit Sub
    End If
    
    For Each rng In Range(RenameCol.Offset(1, 0), Cells(1000000, RenameCol.Column).End(xlUp))
        If rng.Value <> "" Then
            OldPath = rng.Offset(0, -3).Value
            NewPath = Left(OldPath, InStrRev(OldPath, "\")) & rng.Value
            If OldPath <> "" Then
                OldPathShort = "..." & Right(OldPath, Len(OldPath) - Len(Range("A1").End(xlToRight).Offset(0, 1).Value))
                NewPathShort = "..." & Right(NewPath, Len(NewPath) - Len(Range("A1").End(xlToRight).Offset(0, 1).Value))
                If MsgBox("Rename file" & vbNewLine & OldPathShort & vbNewLine & "to:" & vbNewLine & NewPathShort & "?", vbYesNo) = vbYes Then
                    On Error GoTo ErrHandler
                    Name OldPath As NewPath
                    On Error GoTo 0
                    If Not ErrorExist Then
                        If rng.EntireRow.Resize(1, 1).End(xlToRight).Font.Color = rgbCrimson Then
                            Set oFile = oFSO.getfile(NewPath)
                            Call WriteFileName(oFile, rng.EntireRow.Resize(1, 1).End(xlToRight))
                        Else
                            Set oFile = oFSO.getfolder(NewPath)
                            Call WriteFolderName(oFile, rng.EntireRow.Resize(1, 1).End(xlToRight))
                            For Each rng2 In Range(Cells(1, iFirstFunctionalColumn + 1), Cells(1, iFirstFunctionalColumn + 1).End(xlDown))
                                If rng2 <> NewPath Then
                                    rng2.Replace What:=OldPath, Replacement:=NewPath, LookAt:=xlPart
                                End If
                            Next rng2
                        End If
                        rng.Value = ""
                    End If
                    ErrorExist = False
                End If
            End If
        End If
    Next rng
    
    Application.EnableEvents = True

Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    ErrorExist = True
    Resume Next
End Sub
