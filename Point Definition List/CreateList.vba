Public Sub CreateList()

' This macro creates a point list from a point definition report.
' AddTrendInfo can then be run to add trend information from trend report.
'
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
HeaderRange = "A1:AE1"

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
Set writehere = Workbooks(writebook).Sheets(writesheet)
With writehere
    .Range("A1").Cells.Value = "Point System Name"
    .Range("B1").Cells.Value = "Point Name"
    .Range("C1").Cells.Value = "Panel Name"
    .Range("D1").Cells.Value = "Descriptor"
    .Range("E1").Cells.Value = "Point Type"
    .Range("F1").Cells.Value = "Point Address"
    .Range("G1").Cells.Value = "Proof Point Address"
    .Range("H1").Cells.Value = "Engineering Units"
    .Range("I1").Cells.Value = "COV Limit"
    .Range("J1").Cells.Value = "Sensor Type"
    .Range("K1").Cells.Value = "Slope"
    .Range("L1").Cells.Value = "Intercept"
    .Range("M1").Cells.Value = "# of Decimal Places"
    .Range("N1").Cells.Value = "Mode Delay (min)"
    .Range("O1").Cells.Value = "Level Delay (sec)"
    .Range("P1").Cells.Value = "Differential"
    .Range("Q1").Cells.Value = "Setpoint Value/Name"
    .Range("R1").Cells.Value = "Offset1"
    .Range("S1").Cells.Value = "Priority1"
    .Range("T1").Cells.Value = "Offset2"
    .Range("U1").Cells.Value = "Priority2"
    .Range("V1").Cells.Value = "Mode Point"
    .Range("W1").Cells.Value = "Trended"
    .Range("X1").Cells.Value = "RENO Normal"
    .Range("Y1").Cells.Value = "RENO Failed"
    .Range("Z1").Cells.Value = "RENO Pri1"
    .Range("AA1").Cells.Value = "RENO Pri2"
    .Range("AB1").Cells.Value = "RENO Pri3"
    .Range("AC1").Cells.Value = "RENO Pri4"
    .Range("AD1").Cells.Value = "RENO Pri5"
    .Range("AE1").Cells.Value = "RENO Pri6"
End With

'Format the header range
Range(HeaderRange).Select
Selection.Font.Bold = True
Range(HeaderRange).WrapText = True
Range(HeaderRange).HorizontalAlignment = xlLeft

'Freeze window to keep header alway at top of page
'Rows("2:2").Select
'ActiveWindow.FreezePanes = True

'Prepare for reading file
Set writehere = Workbooks(writebook).Sheets(writesheet).Range(startcell)
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.OpenTextFile(readfile)

'
'
' Read File
'
'

'Process file
thisline = f.readline
Do While f.atendofstream <> True
    
    'Check for new point
    If InStr(thisline, "Point System Name:") Then
        startpoint = 23
        endpoint = Len(thisline)
        PointSystemName = Mid(thisline, startpoint, endpoint - startpoint)

        'initialize all tags which may not exist in a point definition
        PointName = ""
        ProofPointAddress = ""
        SensorType = ""
        ProofDelay = ""
        ZeroState = ""
        OneState = ""
        Slope = ""
        Intercept = ""
        InitialValue = ""
        COVLimit = ""
        EngineeringUnits = ""
        DecimalPlaces = ""
        Equipment = ""
        ModePoint = ""
        LevelDelay = ""
        ModeDelay = ""
        Differential = ""
        DefaultDestination1 = ""
        SetpointValue = ""
        Offset1 = ""
        Priority1 = ""
        Offset2 = ""
        Priority2 = ""
        ModePoint = ""
        
        'Get point data
        thisline = f.readline
            'Point Name
            If Mid(thisline, 2, 11) = "Point Name:" Then
                PointName = Mid(thisline, (11 + 5), Len(thisline) - (11 + 5))
            End If
            
        Do While (Mid(thisline, 2, 18) <> "Point System Name:") And (InStr(thisline, "**********") = 0)
            thisline = f.readline
            'Point Type
            If Mid(thisline, 2, 11) = "Point Type:" Then
                PointType = Mid(thisline, (11 + 5), Len(thisline) - (11 + 5))
            End If
            'Descriptor
            If Mid(thisline, 2, 11) = "Descriptor:" Then
                Descriptor = Mid(thisline, (11 + 5), Len(thisline) - (11 + 5))
            End If
            'Panel Name (assumes all panel names are 15 characters in length)
            If Mid(thisline, 2, 11) = "Panel Name:" Then
                PanelName = Mid(thisline, (11 + 5), 15)
            End If
            'Point Address
            If Mid(thisline, 2, 14) = "Point Address:" Then
                PointAddress = Mid(thisline, (14 + 5), 31)
            End If
            If Mid(thisline, 2, 21) = "On/Off Point Address:" Then
                PointAddress = Mid(thisline, (21 + 5), 31)
            End If
            'Proof Point Address
            If Mid(thisline, 2, 20) = "Proof Point Address:" Then
                ProofPointAddress = Mid(thisline, (20 + 5), 31)
            End If
            'Sensor Type
            If Mid(thisline, 2, 12) = "Sensor Type:" Then
                SensorType = Mid(thisline, (12 + 5), Len(thisline) - (12 + 5))
            End If
            'Proof Delay
            If Mid(thisline, 2, 12) = "Proof Delay:" Then
                ProofDelay = Mid(thisline, (12 + 5), Len(thisline) - (12 + 5))
            End If
            'Zero State
            If Mid(thisline, 2, 29) = "                         0 - " Then
                ZeroState = Mid(thisline, (29 + 1), Len(thisline) - (29 + 1))
            End If
            'One State
            If Mid(thisline, 2, 29) = "                         1 - " Then
                OneState = Mid(thisline, (29 + 1), Len(thisline) - (29 + 1))
            End If
            'Slope and Intercept
            If Mid(thisline, 2, 6) = "Slope:" Then
                Tempvar = InStr(thisline, "Intercept:")
                Slope = Mid(thisline, (6 + 5), (Tempvar - 14))
                Intercept = Mid(thisline, (Tempvar + 13), Len(thisline) - (Tempvar + 13))
            End If
            'Initial Value
            If Mid(thisline, 2, 14) = "Initial Value:" Then
                InitialValue = Mid(thisline, (14 + 5), Len(thisline) - (14 + 5))
            End If
            'COV Limit
            If Mid(thisline, 2, 10) = "COV Limit:" Then
                If PointType = "LAI" Then
                    Tempvar = InStr(thisline, "Wire")
                    COVLimit = Mid(thisline, (10 + 5), (Tempvar - 18))
                Else
                    COVLimit = Mid(thisline, (10 + 5), Len(thisline) - (10 + 5))
                End If
            End If
            'Engineering Units
            If Mid(thisline, 2, 18) = "Engineering Units:" Then
                EngineeringUnits = Mid(thisline, (18 + 5), Len(thisline) - (18 + 5))
            End If
            '# of Decimal Places
            If Mid(thisline, 2, 20) = "# of decimal places:" Then
                DecimalPlaces = Mid(thisline, (20 + 5), Len(thisline) - (20 + 5))
            End If
            'Enabled for RENO
            If Mid(thisline, 2, 17) = "Enabled for RENO:" Then
                EnabledforRENO = Mid(thisline, (17 + 5), Len(thisline) - (17 + 5))
            End If
            'Alarm Issue Management and Equipment
            If Mid(thisline, 2, 23) = "Alarm Issue Management:" Then
                If (Mid(thisline, (23 + 5), 1)) = "N" Then
                    AlarmIssueManagement = Mid(thisline, (23 + 5), Len(thisline) - (23 + 5))
                Else
                    Tempvar = InStr(thisline, "Equipment:")
                    AlarmIssueManagement = Mid(thisline, (23 + 5), (Tempvar - 31))
                    Equipment = Mid(thisline, (Tempvar + 13), Len(thisline) - (Tempvar + 13))
                End If
            End If
            'Graphic Name
            If Mid(thisline, 2, 13) = "Graphic Name:" Then
                GraphicName = Mid(thisline, (13 + 5), Len(thisline) - (13 + 5))
            End If
            'Mode Point
            If Mid(thisline, 2, 11) = "Mode Point:" Then
                ModePoint = Mid(thisline, (11 + 5), Len(thisline) - (11 + 5))
            End If
            'Level Delay (sec)
            If Mid(thisline, 2, 19) = "Level Delay (sec.):" Then
                LevelDelay = Mid(thisline, (19 + 5), Len(thisline) - (19 + 5))
            End If
            'Mode Delay (min)
            If Mid(thisline, 2, 18) = "Mode Delay (min.):" Then
                ModeDelay = Mid(thisline, (18 + 5), Len(thisline) - (18 + 5))
            End If
            'Differential
            If Mid(thisline, 2, 13) = "Differential:" Then
                Differential = Mid(thisline, (13 + 5), Len(thisline) - (13 + 5))
            End If
            'Default Destination 1
            If Mid(thisline, 2, 22) = "Default Destination 1:" Then
                DefaultDestination1 = Mid(thisline, (22 + 5), Len(thisline) - (22 + 5))
            End If
            'Setpoint Value/Name
            If Mid(thisline, 2, 15) = "Setpoint Value:" Then
                SetpointValue = Mid(thisline, (15 + 5), Len(thisline) - (15 + 5))
            End If
            If Mid(thisline, 2, 14) = "Setpoint Name:" Then
                SetpointValue = Mid(thisline, (14 + 5), Len(thisline) - (14 + 5))
            End If
            'Offset1, Priority1, Offset2, Priority2
            If Mid(thisline, 2, 6) = "Offset" Then
                thisline = f.readline
                Tempvar = InStr(thisline, ",")
                Offset1 = Mid(thisline, 2, (Tempvar - 3))
                Tempvar2 = InStr((Tempvar + 1), thisline, ",")
                Priority1 = Mid(thisline, (Tempvar + 2), (Tempvar2 - Tempvar - 3))
                thisline = f.readline
                If Mid(thisline, 1, 2) <> """""" Then
                    Tempvar = InStr(thisline, ",")
                    Offset2 = Mid(thisline, 2, (Tempvar - 3))
                    Tempvar2 = InStr((Tempvar + 1), thisline, ",")
                    Priority2 = Mid(thisline, (Tempvar + 2), (Tempvar2 - Tempvar - 3))
                End If
            End If
        Loop
        'End Get point data loop
     
'
'
' Write Worksheet
'
'
     
        ' Write Point Information to worksheet
        sheetoffset = sheetoffset + 1
        With writehere
            .Offset(sheetoffset, 0).Value = PointSystemName
            .Offset(sheetoffset, 1).Value = PointName
            .Offset(sheetoffset, 2).Value = PanelName
            .Offset(sheetoffset, 3).Value = Descriptor
            .Offset(sheetoffset, 4).Value = PointType
            .Offset(sheetoffset, 5).Value = PointAddress
            .Offset(sheetoffset, 6).Value = ProofPointAddress
            .Offset(sheetoffset, 7).Value = EngineeringUnits
            .Offset(sheetoffset, 8).Value = COVLimit
            .Offset(sheetoffset, 9).Value = SensorType
            .Offset(sheetoffset, 10).Value = Slope
            .Offset(sheetoffset, 11).Value = Intercept
            .Offset(sheetoffset, 12).Value = DecimalPlaces
            .Offset(sheetoffset, 13).Value = ModeDelay
            .Offset(sheetoffset, 14).Value = LevelDelay
            .Offset(sheetoffset, 15).Value = Differential
            .Offset(sheetoffset, 16).Value = SetpointValue
            .Offset(sheetoffset, 17).Value = Offset1
            .Offset(sheetoffset, 18).Value = Priority1
            .Offset(sheetoffset, 19).Value = Offset2
            .Offset(sheetoffset, 20).Value = Priority2
            '.Offset(sheetoffset, 7).Value = ProofDelay
            '.Offset(sheetoffset, 8).Value = ZeroState
            '.Offset(sheetoffset, 9).Value = OneState
            '.Offset(sheetoffset, 12).Value = InitialValue
            '.Offset(sheetoffset, 16).Value = EnabledforRENO
            '.Offset(sheetoffset, 17).Value = AlarmIssueManagement
            '.Offset(sheetoffset, 18).Value = Equipment
            '.Offset(sheetoffset, 19).Value = GraphicName
            .Offset(sheetoffset, 21).Value = ModePoint
            '.Offset(sheetoffset, 24).Value = DefaultDestination1
        End With
        'End Write Point Information to Worksheet
   
    'If line does not a new point, read new line
    Else
        thisline = f.readline
    End If

Loop
'End Process file loop

'
'
' Cleanup
'
'

'Close File
f.Close

'Set Column Width
Range(HeaderRange).EntireColumn.AutoFit
'Columns("B:B").ColumnWidth = 6.71
'Columns("H:H").ColumnWidth = 7.29
'Columns("O:O").ColumnWidth = 12.57
'Columns("P:P").ColumnWidth = 8.71
'Columns("Q:Q").ColumnWidth = 9.71
'Columns("R:R").ColumnWidth = 12.57
'Columns("V:V").ColumnWidth = 7.71
'Columns("W:W").ColumnWidth = 7.71

End Sub