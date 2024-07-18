VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "reOrder_384_To_96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''DESCRIPTION'''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT: The input for this macro is a column containing genotype calls from the Quant Studio export files.
'   The entries in this column are from a 384-well plate that was run on the QS5.
'OUTPUT: The output for this macro is a column containing the entries from the INPUT column.
'   The entries in this column will be placed in the order they were collected in the field and placed in
'   the 96-well plates. The macro will automatically add missing data entries for those samples that were not of
'   sufficient quality and were therefore ommitted from the genotyping analysis. This ensures that the input and
'   output columns are of equal length and that genotype calls for the 384-well plate are assigned their appropriate
'   location in their corresponding 96-well plates.

Sub reOrder_384_To_96()

    Dim plate_384_array(1 To 16, 1 To 24) As Variant
    Dim rowDictionary As Object
    Dim results As Range
    Dim plateOneArray As New Collection
    Dim plateTwoArray As New Collection
    Dim plateThreeArray As New Collection
    Dim plateFourArray As New Collection
    
    Dim totalCounter As Integer
    Dim columnCounter As Integer
    Dim rowCounter As Integer
    Dim wellIterator As Range
    Dim listIterator As Range
    Dim missingData As String
    
    missingData = "-"
    totalCounter = 1
    columnCounter = 1
    rowCounter = 1
    Set listIterator = Application.InputBox("Please select the first cell in the column containing the genotype calls from a 384-well plate", Type:=8)
    Set wellIterator = Application.InputBox("Please select the first cell in the ""well"" column from the ""Results"" tab of the QS5 output file", Type:=8)
    Set results = Application.InputBox("Please select the first cell of the column where you would like to place the formated results", Type:=8)
    
    'Set rowDictionary = CreateObject("scripting.dictionary")
    'rowDictionary.Add 1, "A"
    'rowDictionary.Add 2, "B"
    'rowDictionary.Add 3, "C"
    'rowDictionary.Add 4, "D"
    'rowDictionary.Add 5, "E"
    'rowDictionary.Add 6, "F"
    'rowDictionary.Add 7, "G"
    'rowDictionary.Add 8, "H"
    'rowDictionary.Add 9, "I"
    'rowDictionary.Add 10, "J"
    'rowDictionary.Add 11, "K"
    'rowDictionary.Add 12, "L"
    'rowDictionary.Add 13, "M"
    'rowDictionary.Add 14, "N"
    'rowDictionary.Add 15, "O"
    'rowDictionary.Add 16, "P"
    
    For rowCounter = 1 To 16
        For columnCounter = 1 To 24
            If totalCounter <= 384 Then
                'plate_384_array(rowCounter, columnCounter) = listIterator.Value
                If columnCounter Mod 2 <> 0 And rowCounter Mod 2 <> 0 Then
                    If wellIterator.Value = totalCounter Then
                        plateOneArray.Add listIterator.Value
                        Set listIterator = listIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        Set wellIterator = wellIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        totalCounter = totalCounter + 1
                    Else
                        plateOneArray.Add missingData
                        totalCounter = totalCounter + 1
                    End If
                ElseIf columnCounter Mod 2 = 0 And rowCounter Mod 2 <> 0 Then
                    If wellIterator.Value = totalCounter Then
                        plateTwoArray.Add listIterator.Value
                        Set listIterator = listIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        Set wellIterator = wellIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        totalCounter = totalCounter + 1
                    Else
                        plateTwoArray.Add missingData
                        totalCounter = totalCounter + 1
                    End If
                ElseIf columnCounter Mod 2 = 0 And rowCounter Mod 2 = 0 Then
                    If wellIterator.Value = totalCounter Then
                        plateThreeArray.Add listIterator.Value
                        Set listIterator = listIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        Set wellIterator = wellIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        totalCounter = totalCounter + 1
                    Else
                        plateThreeArray.Add missingData
                        totalCounter = totalCounter + 1
                    End If
                ElseIf columnCounter Mod 2 <> 0 And rowCounter Mod 2 = 0 Then
                    If wellIterator.Value = totalCounter Then
                        plateFourArray.Add listIterator.Value
                        Set listIterator = listIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        Set wellIterator = wellIterator.Offset(RowOffset:=1, ColumnOffset:=0)
                        totalCounter = totalCounter + 1
                    Else
                        plateFourArray.Add missingData
                        totalCounter = totalCounter + 1
                    End If
                End If
            End If
        Next
    Next
    
    Debug.Print plateOneArray.count
    Debug.Print plateTwoArray.count
    Debug.Print plateThreeArray.count
    Debug.Print plateFourArray.count
    
    'Add each of the 4 plate collections to the results range column
    Dim i As Variant
    Dim j As Variant
    Dim twelveCounter As Variant
    Dim sortArray As Variant

    sortArray = 1
    twelveCounter = 1
    For j = 1 To 96
        results.Value = plateOneArray(sortArray)
        Set results = results.Offset(RowOffset:=1, ColumnOffset:=0)
        If (sortArray / 12) <= 7 Then
            sortArray = sortArray + 12
        Else
            twelveCounter = twelveCounter + 1
            sortArray = twelveCounter
        End If
    Next
    
    sortArray = 1
    twelveCounter = 1
    For j = 1 To 96
        results.Value = plateTwoArray(sortArray)
        Set results = results.Offset(RowOffset:=1, ColumnOffset:=0)
        If (sortArray / 12) <= 7 Then
            sortArray = sortArray + 12
        Else
            twelveCounter = twelveCounter + 1
            sortArray = twelveCounter
        End If
    Next
    
    sortArray = 1
    twelveCounter = 1
    For j = 1 To 96
        results.Value = plateThreeArray(sortArray)
        Set results = results.Offset(RowOffset:=1, ColumnOffset:=0)
        If (sortArray / 12) <= 7 Then
            sortArray = sortArray + 12
        Else
            twelveCounter = twelveCounter + 1
            sortArray = twelveCounter
        End If
    Next
    
    sortArray = 1
    twelveCounter = 1
    For j = 1 To 96
        results.Value = plateFourArray(sortArray)
        Set results = results.Offset(RowOffset:=1, ColumnOffset:=0)
        If (sortArray / 12) <= 7 Then
            sortArray = sortArray + 12
        Else
            twelveCounter = twelveCounter + 1
            sortArray = twelveCounter
        End If
    Next
        
End Sub



