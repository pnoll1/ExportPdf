Attribute VB_Name = "Module4"
Sub ExportPdf()
project = ThisWorkbook.Worksheets("BASE").Cells(6, 3)
customer = ThisWorkbook.Worksheets("BASE").Cells(8, 3)
contact = ThisWorkbook.Worksheets("BASE").Cells(9, 3)

'Fill out max pick
'define range to search
'Dim x As Integer
'For x = 11 To 51
    'If IsNumeric(Cells(1, x).Value) Then
    'endRow = x
    'Else
    'Exit For
    'End If
'Next

'find cell with max percentage
'Dim searchRange As Range
'endCell = "O" & endRow
'Range("N29").Value = endCell
'Set searchRange = Range("O11:endCell")
'AddressOfMax = WorksheetFunction.Index(searchRange, WorksheetFunction.Match(WorksheetFunction.Max(searchRange), searchRange, 0)).Address
'Range("N30").Value = AddressOfMax
'set values of load etc equal to mobile max % cells
'Range("Load").Value = Range
'Range("Capacity").Value =
'Range("Percentage").Value = Range(AddressOfMax).Value
'export pdf

If ActiveSheet.Name = "ERECT" Then
    folder = "4 ERECT"
    pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\" & folder & "\PDF\"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\PTC assembly sequence spread sheet -" & project
ElseIf ActiveSheet.Name = "DISMAN" Then
    folder = "6 Dismantle"
    pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\" & folder & "\PDF\"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\PTC assembly sequence spread sheet -" & project
ElseIf ActiveSheet.Name = "ERECT Timeline" Then
    folder = "4 ERECT"
    pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\" & folder & "\PDF\"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\PTC Timeline -" & project
ElseIf ActiveSheet.Name = "DISMAN Timeline" Then
    folder = "6 Dismantle"
    pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\" & folder & "\PDF\"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\PTC Timeline -" & project
ElseIf ActiveSheet.Name = "BASE Timeline" Then
    folder = "3 Base Set"
    pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\" & folder & "\PDF\"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\PTC Timeline -" & project
ElseIf ActiveSheet.Name = "BASE" Then
    folder = "3 Base Set"
    pdfpath = "S:\Sicklesteel Cranes\Engineering\Clients\" & customer & "\" & project & "\" & folder & "\PDF\"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfpath & "\PTC assembly sequence spread sheet -" & project
End If
    

End Sub
