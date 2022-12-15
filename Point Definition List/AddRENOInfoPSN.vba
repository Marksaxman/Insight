Public Sub AddRENOInfoPSN()

' This macro sets up RENO information based on selected RENO point list report.
' It assumes CreateList macro has already been run.
'
' Setup
'
'

'Initialize Parameters
Dim fs, f
sheetoffset = -1 '1 will be added before use
writebook = ActiveWorkbook.Name
startcell = "A2"
writesheet = ActiveSheet.Name

'Setup and confirm data file
readfile = Application _
.GetOpenFilename("Export files (*.csv),*.csv")

If readfile = "" Then Exit Sub
'Open file
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFile(readfile)
filedate = Format(f.DateLastModified, "ddMMMyy")
filedate = UCase(filedate)

'Setup Worksheet Header
'HeaderRange = "AH1:AO1"
'Set writehere = Workbooks(writebook).Sheets(writesheet)
'With writehere
'    .Range("AH1").Cells.Value = "RENO RTN"
'    .Range("AI1").Cells.Value = "RENO Failed"
'    .Range("AJ1").Cells.Value = "RENO PRI1"
'    .Range("AK1").Cells.Value = "RENO PRI2"
'    .Range("AL1").Cells.Value = "RENO PRI3"
'    .Range("AM1").Cells.Value = "RENO PRI4"
'    .Range("AN1").Cells.Value = "RENO PRI5"
'    .Range("AO1").Cells.Value = "RENO PRI6"
'End With
'Format the header range
'Range(HeaderRange).Select
'Selection.Font.Bold = True
'Range(HeaderRange).WrapText = True
'Range(HeaderRange).HorizontalAlignment = xlLeft
'Range(HeaderRange).ColumnWidth = 10

'Prepare for reading file
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile(readfile)

'Process file
thisline = f.readline
Do While f.atendofstream <> True

    'Check for new point
    If ((Mid(thisline, 2, 1) >= "0") And (Mid(thisline, 2, 1) <= "Z")) Then
        
        'initialize all tags which may not exist in a point definition
        RENORTN = ""
        RENOFailed = ""
        RENOPri1 = ""
        RENOPri2 = ""
        RENOPri3 = ""
        RENOPri4 = ""
        RENOPri5 = ""
        RENOPri6 = ""
        
        'Extract point system name
        Tempvar = InStr(thisline, ",")
        PointSystemName = Mid(thisline, 2, (Tempvar - 3))
        
        'Get RENO Group configuration on point configuration
        Tempvar2 = InStr((Tempvar + 2), thisline, ",")
        Tempvar3 = InStr((Tempvar2 + 1), thisline, ",")
        GroupName = Mid(thisline, (Tempvar2 + 2), (Tempvar3 - Tempvar2 - 3))
        If GroupName = "NORMAL" Then
            RENORTN = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "FAILED" Then
            RENOFailed = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "PRI1" Then
            RENOPri1 = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "PRI2" Then
            RENOPri2 = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "PRI3" Then
            RENOPri3 = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "PRI4" Then
            RENOPri4 = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "PRI5" Then
            RENOPri5 = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        ElseIf GroupName = "PRI6" Then
            RENOPri6 = Mid(thisline, (Tempvar3 + 2), (Len(thisline) - Tempvar3 - 5))
        End If
    End If
    
    thisline = f.readline
    'If line does not a new point, read new line
    Do While ((Mid(thisline, 2, 1) < "0") Or (Mid(thisline, 2, 1) > "Z")) And f.atendofstream <> True
        'Get RENO Group configuration
        Tempvar = InStr(thisline, ",")
        If Tempvar <> 0 Then
        Tempvar2 = InStr((Tempvar + 2), thisline, ",")
        GroupName = Mid(thisline, (Tempvar + 2), (Tempvar2 - Tempvar - 3))
        If GroupName = "NORMAL" Then
            RENORTN = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "FAILED" Then
            RENOFailed = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "PRI1" Then
            RENOPri1 = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "PRI2" Then
            RENOPri2 = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "PRI3" Then
            RENOPri3 = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "PRI4" Then
            RENOPri4 = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "PRI5" Then
            RENOPri5 = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        ElseIf GroupName = "PRI6" Then
            RENOPri6 = Mid(thisline, (Tempvar2 + 2), (Len(thisline) - Tempvar2 - 5))
        End If
        End If
        thisline = f.readline

    Loop

    Workbooks(writebook).Sheets(writesheet).Range("A:A").Cells.Find(What:=PointSystemName).Activate
          
    ' Write Point Information to worksheet
    ActiveCell.Offset(0, 23).Value = RENORTN
    ActiveCell.Offset(0, 24).Value = RENOFailed
    ActiveCell.Offset(0, 25).Value = RENOPri1
    ActiveCell.Offset(0, 26).Value = RENOPri2
    ActiveCell.Offset(0, 27).Value = RENOPri3
    ActiveCell.Offset(0, 28).Value = RENOPri4
    ActiveCell.Offset(0, 29).Value = RENOPri5
    ActiveCell.Offset(0, 30).Value = RENOPri6
Loop

'End Process file loop

'Close File
f.Close

End Sub