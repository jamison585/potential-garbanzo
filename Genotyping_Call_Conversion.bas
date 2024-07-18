Attribute VB_Name = "Genotyping_Call_Conversion"
Option Explicit

Sub genotypeConversion()
    
    Dim callInput As Range
    Dim callOutput As Range
    Dim marker As String
    Dim Marker1_a1_Homo As String
    Dim Marker1_a2_Homo As String
    Dim Marker1_Het As String
    Dim Marker2_a1_Homo As String
    Dim Marker2_a2_Homo As String
    Dim Marker2_Het As String
    
    Set callInput = Application.InputBox("Please select the first cell in the ""Call"" column in the ""Results"" tab of the QS5 output excel file.", Type:=8)
    Set callOutput = Application.InputBox("Please select the first cell of the column where you would like to store the results of the genotype conversion.", Type:=8)
    
    marker = "Marker1"
    
    'Nuclear Marker
    Marker1_a1_Homo = "F"
    Marker1_a2_Homo = "S"
    Marker1_Het = "H"
    
    'Cytoplasmic Marker
    Marker2_a1_Homo = "S"
    Marker2_a2_Homo = "N"
    
    While callInput <> 0
        If marker = "Marker1" Then
            If callInput.Value = "Homozygous Allele 1/Allele 1" Then
                callOutput.Value = Marker1_a1_Homo
            ElseIf callInput.Value = "Homozygous Allele 2/Allele 2" Then
                callOutput.Value = Marker1_a2_Homo
            ElseIf callInput.Value = "Heterozygous Allele 1/Allele 2" Then
                callOutput.Value = Marker1_Het
            Else
                callOutput.Value = "-"
            End If
        ElseIf marker = "Marker2" Then
            If callInput.Value = "Homozygous Allele 1/Allele 1" Then
                callOutput.Value = Marker2_a1_Homo
            ElseIf callInput.Value = "Homozygous Allele 2/Allele 2" Then
                callOutput.Value = Marker2_a2_Homo
            Else
                callOutput.Value = "-"
            End If
        End If
        Set callInput = callInput.Offset(1, 0)
        Set callOutput = callOutput.Offset(1, 0)
    Wend
End Sub
