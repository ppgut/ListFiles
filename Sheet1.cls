VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range

    Application.EnableEvents = False

    If Range(Cells(1, Target.Column).Address).Value = "Rename" Then
        For Each rng In Target
           If rng.Value <> "" And Extension(rng.Value) = " " And rng.Offset(0, -2).Value <> " " Then

                rng.Formula = rng.Formula & "&" & rng.Offset(0, -2).Address

            End If
        Next rng
    End If

    Application.EnableEvents = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Cells.Count > 50 Then
        Set Target = Target.Cells(1)
    End If

    Application.ScreenUpdating = False
    Dim TableRange As Range
    Dim Intersection As Range
    
    'Set TableRange = Range(Sheet1.Range("A1"), Sheet1.Range("A1").End(xlToRight).End(xlDown).Offset(0, 3))
    Set TableRange = Range(Sheet1.Range("A1"), Cells(Sheet1.Range("A1").End(xlToRight).End(xlDown).Row, Columns("AD:AD").Column))
    Set Intersection = Intersect(Target, TableRange)
    
    If Not Intersection Is Nothing Then
    
        Dim rng As Range
        Dim i As Integer
        Dim RowAlreadyConsidered As Boolean
        Dim NameValue As String
        Dim Arr() As Variant
        ReDim Arr(1 To 1) As Variant
        
        For Each rng In Target.Cells
            For i = LBound(Arr) To UBound(Arr)
                If rng.Row = Arr(i) Then
                    RowAlreadyConsidered = True
                    Exit For
                End If
            Next i
            If Not RowAlreadyConsidered Then
                If UBound(Arr) = 1 And Arr(1) = Empty Then
                    Arr(1) = rng.Row
                Else
                    ReDim Preserve Arr(1 To UBound(Arr) + 1)
                    Arr(UBound(Arr)) = rng.Row
                End If
            End If
            RowAlreadyConsidered = False
        Next rng
        
        NameValue = "={"
        For i = 1 To UBound(Arr)
            If i = 1 Then
                NameValue = NameValue & Arr(i)
            Else
                NameValue = NameValue & ";" & Arr(i)
            End If
        Next i
        NameValue = NameValue & "}"
        
        ThisWorkbook.Names("SelectedRows").RefersToR1C1 = NameValue
    
    Else
        ThisWorkbook.Names("SelectedRows").RefersToR1C1 = "{}"
    End If
    
    Application.ScreenUpdating = True
End Sub

