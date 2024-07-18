VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "varInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''VARIABLE DECLARATIONS SUB ROUTINE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



'DECLARATION of "CAGE_PLAN" variable
Dim DNA_PRTY As Range, CUMM_NO As Range, NO_1ST_PLATE As Range, NO_2ND_PLATE As Range, NO_3RD_PLATE As Range, _
    PLATE_LOCATION As Range, PLATE_ARRAY As Range, START_CELL As Range
    
'DECLARATION of "PLATE_PLAN" variables
Dim PLATE_LOCATION_GRID As Range, platePosition As Integer, numFirstPlate As Integer, numSecondPlate As Integer, _
    numThirdPlate As Integer, shading As Integer, plateCounter As Integer

'DECLARATION of "PLATE_PLAN" plate representation variables
Dim plateRowLabels(7) As String, plateColumnLabels(11) As Integer, cellColor As Integer, total As Integer
    




