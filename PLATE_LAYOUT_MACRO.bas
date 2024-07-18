Attribute VB_Name = "PLATE_LAYOUT_MACRO"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''GLOBAL VARIABLES DEFINED''''''''''''''''''''''''''''''''''''''''''''''''''''
'variables from the reference worksheet
Dim totalNo As Range
Dim cage As Range
Dim cageRow As Range
'variables for the current worksheet
Dim cell As Range
Dim plate As Range
'variables for labeling the plates
Dim columnLabels As Range
Dim rowLabels As Range
Dim plateLabels As Range
'place holders for Begin and End text
Dim beginPlateText As Range
Dim endPlateText As Range
'counter variables
Dim rowCount As Integer
Dim count As Integer
Dim totalCount As Integer
Dim plateCount As Integer
Dim rowArray As New Collection
Dim cellShading As Boolean

Sub PLATE_MACRO()
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''MAIN METHOD: PLACE CURSOR BELOW THIS LINE AND SELECT RUN TO RUN THIS MACRO'''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''MACRO VARIABLES DEFINED'''''''''''''''''''''''''''''''''''''''''''''''''''
    Set cell = Application.InputBox("Please select cell ""A1"" in the worksheet where you would like to place the plate grid", Type:=8)
    Set cage = Application.InputBox("Please select the first cell in the ""CAGE"" column on the ""PLATE_PLAN"" worksheet", Type:=8)
    Set cageRow = cage.Offset(0, 1)
    cell.Value = "PLATE"
    Set beginPlateText = cell.Offset(RowOffset:=1, ColumnOffset:=14)
    Set endPlateText = cell.Offset(RowOffset:=5, ColumnOffset:=14)
    
    With cell.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With cell.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Set cell = cell.Offset(0, 1)
    With cell.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Set cell = cell.Offset(0, -1)
    Set plate = cell
    Set plate = plate.Resize(8, 12)
    Set totalNo = Application.InputBox("Please select the first cell in the ""DNA_COUNT"" column on the ""PLATE_PLAN"" worksheet", Type:=8)
    
    rowCount = 1
    count = 1
    totalCount = 1
    plateCount = 1
    
    'offset plate range
    Set cell = cell.Offset(1, 2)
    Set plate = plate.Offset(1, 2)
    Set plateLabels = cell.Offset(0, -2)
    Set rowLabels = cell.Offset(0, -1)
    Set columnLabels = cell.Offset(-1, 0)
        
    rowArray.Add ("A")
    rowArray.Add ("B")
    rowArray.Add ("C")
    rowArray.Add ("D")
    rowArray.Add ("E")
    rowArray.Add ("F")
    rowArray.Add ("G")
    rowArray.Add ("H")
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''macro logic''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim c As Integer
    With columnLabels.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    For c = 1 To 12
        columnLabels.Value = c
        With columnLabels.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        If columnLabels.Value = 12 Then
            Exit For
        End If
        Set columnLabels = columnLabels.Offset(0, 1)
    Next c
    With columnLabels.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
start_loop_1:
    
    While totalNo.Value <> ""
        If count <= totalNo.Value Then
            If rowCount Mod 8 <> 0 Then
                cell.Value = count
                If totalCount = 1 Then
                    Call PLATE_BEGIN_TEXT(cage, cageRow, count, totalNo, beginPlateText)
                End If
                Call SHADE(cellShading, cell)
                If totalCount >= 92 Then
                    GoTo end_loop_1
                End If
                Set cell = cell.Offset(1, 0)
                rowCount = rowCount + 1
                count = count + 1
                totalCount = totalCount + 1
                If rowCount > 8 Then
                    rowCount = 1
                End If
            ElseIf rowCount Mod 8 = 0 Then
                Do
                    cell.Value = count
                    Call SHADE(cellShading, cell)
                    If totalCount >= 92 Then
                        Exit Do
                    End If
                    Set cell = cell.Offset(-7, 1)
                    rowCount = rowCount + 1
                    count = count + 1
                    totalCount = totalCount + 1
                    If rowCount > 8 Then
                        rowCount = 1
                    End If
                Loop While False
            End If
        Else
            count = 1
            Set totalNo = totalNo.Offset(1, 0)
            cellShading = Not cellShading
            Set cage = cage.Offset(1, 0)
            Set cageRow = cageRow.Offset(1, 0)
        End If
    Wend
    
    If totalNo.Value = "" Then
        GoTo end_macro
    End If
        
end_loop_1:
    
    'Row and Column Labels are added
    Call LABELS(rowArray, plateLabels, plateCount)
    Call OUTLINES(plate)
    Call PLATE_END_TEXT(cage, cageRow, count, totalNo, endPlateText)
    plateCount = plateCount + 1
    If count = totalNo.Value Then
        Set totalNo = totalNo.Offset(1, 0)
        count = 1
        cellShading = Not cellShading
        Set cage = cage.Offset(1, 0)
        Set cageRow = cageRow.Offset(1, 0)
    Else
        count = count + 1
    End If
    Set cell = cell.Offset(1, 0)
    'The EMPTY_CELLS subroutine is called here at the end of each plate becuase total count is greater than or equal to 92
    Call EMPTY_CELLS(cell)
    rowCount = 1
    totalCount = 1
    GoTo start_loop_1
    
end_macro:

    Call LABELS(rowArray, plateLabels, plateCount)
    Call OUTLINES(plate)
    Call SHADE(cellShading, cell)
    If totalCount = 92 Then
        'Call PLATE_END_TEXT(cage, cageRow, count, totalNo, endPlateText)
        Set cell = cell.Offset(1, 0)
        Call EMPTY_CELLS(cell)
    Else
        Set plateLabels = plateLabels.Offset(-1, 0)
        Set rowLabels = rowLabels.Offset(-1, 0)
        With plateLabels.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rowLabels.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Set cell = rowLabels.Offset(0, 12)
        Dim k As Integer
        For k = 1 To 4
            With cell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Set cell = cell.Offset(-1, 0)
        Next k
    End If
End Sub

'Labels the cage location where each plate STARTS by Cage, Row, and Plant number
Sub PLATE_BEGIN_TEXT(cage, cageRow, count, totalNo, beginPlateText)
    beginPlateText.Value = "Begin:"
    Set beginPlateText = beginPlateText.Offset(1, 1)
    beginPlateText.Value = "Cage: " & cage.Value
    Set beginPlateText = beginPlateText.Offset(1, 0)
    beginPlateText.Value = "Row" & ": " & cageRow.Value
    Set beginPlateText = beginPlateText.Offset(1, 0)
    beginPlateText.Value = "Plant" & ": " & count
    Set beginPlateText = beginPlateText.Offset(5, -1)
    'Increment the "cage" and "cageRow" variables to the next cell down in their respective columns.
    'Set cage = cage.Offset(1, 0)
    'Set cageRow = cageRow.Offset(1, 0)
End Sub

'Labels the Cage location where each plate STOPS by Cage, Row, and Plant number
Sub PLATE_END_TEXT(cage, cageRow, count, totalNo, endPlateText)
    endPlateText.Value = "End:"
    Set endPlateText = endPlateText.Offset(1, 1)
    endPlateText.Value = "Cage: " & cage.Value
    Set endPlateText = endPlateText.Offset(1, 0)
    endPlateText.Value = "Row: " & cageRow.Value
    Set endPlateText = endPlateText.Offset(1, 0)
    endPlateText = "Plant: " & count
    Set endPlateText = endPlateText.Offset(5, -1)
    'Increment the "cage" and "cageRow" variables to the next cell down in their respective columns.
    'Set cage = cage.Offset(1, 0)
    'Set cageRow = cageRow.Offset(1, 0)
End Sub

'Alternate shading is applied for the boundaries of each row in a plate
Sub SHADE(cellShading As Boolean, cell As Range)
    If cellShading = True Then
        With cell.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
    Else
        With cell.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .Color = xlNone
            .PatternTintAndShade = 0
        End With
    End If
End Sub

'Plate outlines are drawn around each plate once each of the cells are completed
Sub OUTLINES(plate As Range)
    With plate.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plate.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plate.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plate.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Set plate = plate.Offset(8, 0)
End Sub

'Row Labels are added to each plate
Sub LABELS(rowArray As Collection, plateLabels As Range, plateCount As Integer)
    
    Dim r As Variant
    With rowLabels.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    For Each r In rowArray
        rowLabels.Value = r
        With rowLabels.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rowLabels.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Set rowLabels = rowLabels.Offset(1, 0)
    Next r
    Dim p As Integer
    For p = 1 To 8
        plateLabels.Value = plateCount
        With plateLabels.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        Set plateLabels = plateLabels.Offset(1, 0)
    Next p
End Sub

'Fills the last 4 cells of each plate black to indicate that they are not used for sample processing
Sub EMPTY_CELLS(cell As Range)
    Dim i As Integer
    For i = 1 To 4
        With cell.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Set cell = cell.Offset(1, 0)
    Next i
    Set cell = cell.Offset(0, -11)
End Sub







