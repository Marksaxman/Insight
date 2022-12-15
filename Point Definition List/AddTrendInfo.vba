Public Sub AddTrendInfo()

' This macro sets up trend information based on selected trend report.
' It assumes CreateList macro has already been run.
'
' Setup
'
'


'Initialize Parameters
Dim fs, f
writebook = ActiveWorkbook.Name
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

'Prepare for reading file
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile(readfile)

'Process file
thisline = f.readline
Do While f.atendofstream <> True
    
    'Check for new point

    If InStr(thisline, "Point Name:") Then
        startpoint = 16
        endpoint = Len(thisline)
        PointName = Mid(thisline, startpoint, endpoint - startpoint)
        Workbooks(writebook).Sheets(writesheet).Range("B:B").Cells.Find(What:=PointName).Activate
        
        ' Write Point Information to worksheet
        ActiveCell.Offset(0, 21).Value = "Y"
        'ActiveCell.Offset(0, 31).Value = "Y"
        'ActiveCell.Offset(0, 32).Value = "Y"

    'If line does not a new point, read new line
    End If

    thisline = f.readline
Loop
'End Process file loop

'Close File
f.Close

End Sub