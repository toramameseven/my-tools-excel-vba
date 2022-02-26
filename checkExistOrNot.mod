Sub checkExistOrNot()
    
    'select range
    'output which column has which item or not.

    '' original data
    Dim dataColumns As Long
    Dim dataRows As Long
    Dim rngData As Range
    
    
    Set rngData = Selection
    dataColumns = rngData.columns.Count
    dataRows = rngData.Rows.Count
    
    If dataColumns < 2 Then
        MsgBox "Select 2 or more columns"
        Exit Sub
    End If
    
    
    Dim data As Variant
    data = rngData
    
    '' output range
    Dim rangeDiff As Range
    Set rangeDiff = Application.InputBox(prompt:="", title:="Select the topleft cell for Output.", Type:=8)
    If WorksheetFunction.CountA(rangeDiff.Resize(dataRows * dataColumns, dataColumns + 1)) <> 0 Then
        MsgBox "Output range is not Empty."
        Exit Sub
    End If

    
    '' dictionary
    Dim dic() As Object
    ReDim dic(1 To dataColumns)
    Dim c As Long
    
    For c = 1 To dataColumns
        Set dic(c) = CreateObject("Scripting.Dictionary")
    
    Next
    
    Dim dicDiff As Object
    Set dicDiff = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim v As String

    For c = 1 To dataColumns
        For i = 1 To dataRows
            v = data(i, c)
            If v <> "" Then
                If Not dic(c).Exists(v) Then
                    dic(c).Add v, v
                Else
                    MsgBox v & "is duplicate. column:" & c
                    Exit Sub
                End If
                
                If Not dicDiff.Exists(v) Then
                    dicDiff.Add v, v
                End If
            End If
        Next
    Next
    
    
    '' sort key
    Dim allKeys As Variant  '
    allKeys = dicDiff.Keys
    
    Dim diffArrayList As Object
    Set diffArrayList = CreateObject("System.Collections.ArrayList")
    
    For i = 0 To UBound(allKeys)
        diffArrayList.Add (allKeys(i))
    Next
    diffArrayList.Sort
    
    
    '' make result and output the result
    Dim outValue As Variant
    outValue = rangeDiff.Resize(UBound(allKeys) + 1, dataColumns + 1)
    
    For i = 0 To UBound(allKeys)
        outValue(i + 1, 1) = diffArrayList(i)
        For c = 1 To dataColumns
            outValue(i + 1, c + 1) = dic(c).Exists(diffArrayList(i))
        Next
    Next
    
    rangeDiff.Resize(UBound(allKeys) + 1, dataColumns + 1) = outValue
End Sub
