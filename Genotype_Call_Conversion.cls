VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Genotype_Call_Conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub Genotype_Call_Conversion()
    
    Dim callInput As Range
    Dim callOutput As Range
    Dim marker As String
    Dim AcPms_a1_Homo As String
    Dim AcPms_a2_Homo As String
    Dim AcPms_Het As String
    Dim CMS_a1_Homo As String
    Dim CMS_a2_Homo As String
    Dim CMS_Het As String
    
    Set callInput = Application.InputBox("Please select the first cell in the ""Call"" column in the ""Results"" tab of the QS5 output excel file.", Type:=8)
    Set callOutput = Application.InputBox("Please select the first cell of the column where you would like to store the results of the genotype conversion.", Type:=8)
    
    marker = "CMS"
    
    AcPms_a1_Homo = "F"
    AcPms_a2_Homo = "S"
    AcPms_Het = "H"
    CMS_a1_Homo = "S"
    CMS_a2_Homo = "N"
    
    While callInput <> 0
        If marker = "AcPms" Then
            If callInput.Value = "Homozygous Allele 1/Allele 1" Then
                callOutput.Value = AcPms_a1_Homo
            ElseIf callInput.Value = "Homozygous Allele 2/Allele 2" Then
                callOutput.Value = AcPms_a2_Homo
            ElseIf callInput.Value = "Heterozygous Allele 1/Allele 2" Then
                callOutput.Value = AcPms_Het
            Else
                callOutput.Value = "-"
            End If
        ElseIf marker = "CMS" Then
            If callInput.Value = "Homozygous Allele 1/Allele 1" Then
                callOutput.Value = CMS_a1_Homo
            ElseIf callInput.Value = "Homozygous Allele 2/Allele 2" Then
                callOutput.Value = CMS_a2_Homo
            Else
                callOutput.Value = "-"
            End If
        End If
        Set callInput = callInput.Offset(1, 0)
        Set callOutput = callOutput.Offset(1, 0)
    Wend
End Sub

