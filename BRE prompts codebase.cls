VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateStandardDeviation()
    ' Declare variables
    Dim numberOfResults As String
    Dim characteristicStrength As Double
    Dim standardDeviation As Double

    ' Read user selections
    numberOfResults = ThisWorkbook.Sheets("BRE prompts").Range("B16").Value
    characteristicStrength = ThisWorkbook.Sheets("BRE prompts").Range("B17").Value

    ' Perform calculations based on user selections
    If characteristicStrength < 20 Then
        If numberOfResults = "<20" Then
            ' Use formula for A; s for less than 20 results
            standardDeviation = characteristicStrength * 0.4
        Else
            ' Use formula for A; minimum s for 20 or more 20 results
            standardDeviation = characteristicStrength * 0.2
        End If
    Else
        If numberOfResults = "<20" Then
            ' Use formula for B; s for less than 20 results
            standardDeviation = 8
        Else
            ' Use formula for B; minimum s for 20 or more 20 results
            standardDeviation = 4
        End If
    End If

    ' Display the result in the sheet
    ThisWorkbook.Sheets("BRE prompts").Range("B18").Value = standardDeviation
End Sub

Sub GetCompressiveStrength()
    ' Declare variables
    Dim cementStrengthClass As String
    Dim typeOfAggregate As String
    Dim ageDays As Integer
    Dim compressiveStrength As Double
    
    ' Read user selections
    cementStrengthClass = ThisWorkbook.Sheets("BRE prompts").Range("B9").Value
    typeOfAggregate = ThisWorkbook.Sheets("BRE prompts").Range("B10").Value
    ageDays = ThisWorkbook.Sheets("BRE prompts").Range("B11").Value
    
    ' Access the table in the "BRE Mix Design Charts" worksheet
    Dim mixDesignSheet As Worksheet
    Set mixDesignSheet = ThisWorkbook.Sheets("BRE Mix Design Charts")
    
    ' Define variables to store row and column indexes
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    ' Determine row index based on Cement Strength Class and Type of Aggregate
    If cementStrengthClass = "42.5" Then
        If typeOfAggregate = "Uncrushed" Then
            rowIndex = 35
        Else
            rowIndex = 36
        End If
    Else
        If typeOfAggregate = "Uncrushed" Then
            rowIndex = 37
        Else
            rowIndex = 38
        End If
    End If
    
    ' Determine column index based on Age (days)
    Select Case ageDays
        Case 3
            colIndex = 4
        Case 7
            colIndex = 5
        Case 28
            colIndex = 6
        Case 91
            colIndex = 7
    End Select
    
    ' Read the compressive strength from the table
    compressiveStrength = mixDesignSheet.Cells(rowIndex, colIndex).Value
    
    ' Display the result in the sheet
    ThisWorkbook.Sheets("BRE prompts").Range("B12").Value = compressiveStrength
End Sub


Sub GetFreeWaterContent()
    ' Declare variables
    Dim MaxAggregateSize As String
    Dim typeOfAggregate As String
    Dim slumpCondition As String
    Dim freeWaterContent As Double
    
    ' Read user selections
    MaxAggregateSize = ThisWorkbook.Sheets("BRE prompts").Range("B2").Value
    typeOfAggregate = ThisWorkbook.Sheets("BRE prompts").Range("B3").Value
    slumpCondition = ThisWorkbook.Sheets("BRE prompts").Range("B4").Value
    
    ' Access the table in the "BRE Mix Design Charts" worksheet
    Dim mixDesignSheet As Worksheet
    Set mixDesignSheet = ThisWorkbook.Sheets("BRE Mix Design Charts")
    
    ' Define variables to store row and column indexes
    Dim rowIndex As Integer
    Dim colIndex As Integer
    
    ' Determine row index based on Maximum Aggregate Size and Type of Aggregate
    Select Case MaxAggregateSize
        Case "10mm"
            If typeOfAggregate = "Uncrushed" Then
                rowIndex = 6
            Else
                rowIndex = 7
            End If
        Case "20mm"
            If typeOfAggregate = "Uncrushed" Then
                rowIndex = 8
            Else
                rowIndex = 9
            End If
        Case "40mm"
            If typeOfAggregate = "Uncrushed" Then
                rowIndex = 10
            Else
                rowIndex = 11
            End If
    End Select
    
    ' Determine column index based on Slump condition
    Select Case slumpCondition
        Case "0-10mm"
            colIndex = 16
        Case "10-30mm"
            colIndex = 17
        Case "30-60mm"
            colIndex = 18
        Case "60-180mm"
            colIndex = 19
    End Select
    
    ' Read the free water content from the table
    freeWaterContent = mixDesignSheet.Cells(rowIndex, colIndex).Value
    
    ' Display the result in the sheet
    ThisWorkbook.Sheets("BRE prompts").Range("B5").Value = freeWaterContent
End Sub

