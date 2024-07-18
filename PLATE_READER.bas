Attribute VB_Name = "PLATE_READER"
Option Explicit

Dim wellCount As Range
Dim plateNum As Range
Dim wellPosition As Range
Dim startGrid As Range

Sub reRun_PlateLayout()

    Set wellCount = Application.InputBox("Please select the first cell in the ""Count"" column", Type:=8)
    Set plateNum = Application.InputBox("Please select the first cell in the ""Plate"" column", Type:=8)
    Set wellPosition = Application.InputBox("Please select the first cell in the ""Well"" column", Type:=8)
    Set startGrid = Application.InputBox("Please select the top left cell where you would like the plate grid to be positioned", Type:=8)
    
    While wellPosition.Value <> ""
        For i = 1 To 12
            For j = 1 To 8
                
                
                
            Next j
        Next i
    Wend
    
End Sub
