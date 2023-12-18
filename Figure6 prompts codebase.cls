VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateProportion()
    ' Declare variables to store user selections
    Dim MaxAggregateSize As String
    Dim SlumpSelection As String
    Dim Passing600um As Double
    Dim InputValue As Double
    
    ' Read user selections from cells B1 to B4
    MaxAggregateSize = ThisWorkbook.Sheets("Figure6 prompts").Range("B1").Value
    SlumpSelection = ThisWorkbook.Sheets("Figure6 prompts").Range("B2").Value
    Passing600um = CDbl(ThisWorkbook.Sheets("Figure6 prompts").Range("B3").Value)
    InputValue = CDbl(ThisWorkbook.Sheets("Figure6 prompts").Range("B4").Value)
    
    ' Navigate to the correct sheet based on MaxAggregateSize
    Dim SelectedSheet As Worksheet
    Set SelectedSheet = ThisWorkbook.Sheets("Figure 6, " & MaxAggregateSize)
    
    ' Construct the table name based on user selection
    Dim TableName As String
    If SlumpSelection = "0-10mm" Then
        TableName = "F0_" & MaxAggregateSize
    ElseIf SlumpSelection = "10-30mm" Then
        TableName = "F1_" & MaxAggregateSize
    ElseIf SlumpSelection = "30-60mm" Then
        TableName = "F2_" & MaxAggregateSize
    ElseIf SlumpSelection = "60-180mm" Then
        TableName = "F3_" & MaxAggregateSize
    End If
    
    ' Access the table
    Dim SelectedTable As ListObject
    On Error Resume Next
    Set SelectedTable = SelectedSheet.ListObjects(TableName)
    On Error GoTo 0

    If SelectedTable Is Nothing Then
        ' Handle the case when the selected table doesn't exist
        ThisWorkbook.Sheets("Figure6 prompts").Range("B6").Value = "Table not found"
    Else
        ' Identify the user's choice for the row name in the table
        Dim AvailableOptions As Variant
        Dim SelectedOption As Double
        Dim PercentageString As String
        Dim equation As String

        ' Define available options and convert them to numerical equivalents
        AvailableOptions = Array(0.15, 0.4, 0.6, 0.8, 1#)

        ' Compare Passing600um value with the AvailableOptions
        SelectedOption = AvailableOptions(0) ' initialize with the first option

        ' Find the option closest to Passing600um value
        For Each Row In SelectedTable.ListRows
            If Abs(Passing600um - CDbl(Row.Range(1, 1).Value)) < Abs(Passing600um - SelectedOption) Then
                SelectedOption = CDbl(Row.Range(1, 1).Value)
                equation = Row.Range(1, 2).Value ' Assuming the equation is in the second column
            End If
        Next Row

        ' Choose the equation based on the selected option
        If SelectedOption = 0.15 Then
            equation = SelectedTable.ListRows(1).Range(1, 2).Value
        ElseIf SelectedOption = 0.4 Then
            equation = SelectedTable.ListRows(2).Range(1, 2).Value
        ElseIf SelectedOption = 0.6 Then
            equation = SelectedTable.ListRows(3).Range(1, 2).Value
        ElseIf SelectedOption = 0.8 Then
            equation = SelectedTable.ListRows(4).Range(1, 2).Value
        ElseIf SelectedOption = 1# Then
            equation = SelectedTable.ListRows(5).Range(1, 2).Value
        Else
            ThisWorkbook.Sheets("Figure6 prompts").Range("B6").Value = "No row found"
            Exit Sub
        End If

        ' Convert the numerical value of SelectedOption to percentage and then into a string
        PercentageString = Format(SelectedOption * 100, "0%")

        ' Identify the user's choice for the row name in the table
        Dim selectedRow As Range
        On Error Resume Next
        Set selectedRow = SelectedTable.ListColumns("Percentage").DataBodyRange.Find(What:=PercentageString, LookIn:=xlValues)
        On Error GoTo 0

        If selectedRow Is Nothing Then
            ' Handle the case when the selected row doesn't exist
            ThisWorkbook.Sheets("Figure6 prompts").Range("B6").Value = "No row found"
            Exit Sub
        Else
            ' Calculate Proportion of fine aggregate using the straight-line equation
            Dim FreeWaterCementRatio As Double
            Dim ProportionOfFineAggregate As Double

            ' Assume the Free-water/cement ratio is in cell B5
            FreeWaterCementRatio = CDbl(ThisWorkbook.Sheets("Figure6 prompts").Range("B5").Value)

            ' Use the straight-line equation: y = mx + c
            ' Calculate Proportion of fine aggregate using the straight-line equation
            ProportionOfFineAggregate = EvaluateTrendlineEquation(equation, FreeWaterCementRatio)

            ' Display the result in cell B5
            ThisWorkbook.Sheets("Figure6 prompts").Range("B5").Value = ProportionOfFineAggregate
        End If
    End If
End Sub

Function EvaluateTrendlineEquation(equation As String, xValue As Double) As Double
    ' Assuming the equation is in the form of "y = mx + c"
    ' Extract coefficients (m and c) from the equation
    Dim equationParts() As String
    equationParts = Split(equation, " ")

    ' Extract coefficients
    Dim coefficientM As Double
    Dim coefficientC As Double
    coefficientM = CDbl(equationParts(2))
    coefficientC = CDbl(equationParts(4))

    ' Calculate y = mx + c
    EvaluateTrendlineEquation = coefficientM * xValue + coefficientC
End Function

