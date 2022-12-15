Public Sub EnhancedAlarms()

' This macro adds level delay information based on selected enhanced alarm report.
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
        
        'Initialize data
        LevelDelay = ""
        
        thisline = f.readline
        'Point Name
        If InStr(thisline, "Point Name:") Then
            startpoint = 16
            endpoint = Len(thisline)
            PointName = Mid(thisline, startpoint, endpoint - startpoint)
        End If
        
        EndOfLine = """" & """"
        Do While (Mid(thisline, 2, 11) <> "Point Name:") And (InStr(thisline, EndOfLine) = 0)
            thisline = f.readline
        
            'Level Delay (sec)
            If Mid(thisline, 2, 19) = "Level Delay (sec.):" Then
                LevelDelay = Mid(thisline, (19 + 5), Len(thisline) - (19 + 5))
            End If
        Loop
        Workbooks(writebook).Sheets(writesheet).Range("B:B").Cells.Find(What:=PointName).Activate
        
        ' Write Point Information to worksheet
        ActiveCell.Offset(0, 13).Value = LevelDelay

    'If line does not a new point, read new line
    End If

    thisline = f.readline
Loop
'End Process file loop

'Close File
f.Close

End Sub