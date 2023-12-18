VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GetAverageEValue()
    ' Declare variables to store user selections
    Dim CEMType As String
    Dim SampleType As String
    Dim AverageECell As Range
    Dim AverageE As Double
    
    ' Read user selections from cells B1 to B2
    CEMType = ThisWorkbook.Sheets("e, CEM prompts").Range("B1").Value
    SampleType = ThisWorkbook.Sheets("e, CEM prompts").Range("B2").Value
    
    ' Navigate to the correct sheet based on CEMType
    Dim SelectedSheet As Worksheet
    Set SelectedSheet = ThisWorkbook.Sheets("e, CEM " & CEMType)
    
    ' Construct the table name based on user selection
    Dim TableName As String
    TableName = "Sample" & SampleType & "_" & CEMType
    
    ' Access the table
    Dim SelectedTable As ListObject
    On Error Resume Next
    Set SelectedTable = SelectedSheet.ListObjects(TableName)
    On Error GoTo 0

    If SelectedTable Is Nothing Then
        ' Handle the case when the selected table doesn't exist
        ThisWorkbook.Sheets("e, CEM prompts").Range("B3").Value = "Table not found"
    Else
        ' Get the header cell of the "e" column plus six rows
        On Error Resume Next
        Set AverageECell = SelectedTable.ListColumns("e").DataBodyRange.Cells(1, 1).Offset(7, 0)
        On Error GoTo 0

        If AverageECell Is Nothing Then
            ' Handle the case when the header cell plus six rows is not found
            ThisWorkbook.Sheets("e, CEM prompts").Range("B3").Value = "Average E not found"
        Else
            ' Read the value from the identified cell
            AverageE = CDbl(AverageECell.Value)
            
            ' Display the result in cell B3
            ThisWorkbook.Sheets("e, CEM prompts").Range("B3").Value = AverageE
        End If
    End If
End Sub
