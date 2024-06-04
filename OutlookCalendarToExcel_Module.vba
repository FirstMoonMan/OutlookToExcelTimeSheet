Sub CreateTimesheetFromOutlook()
' Pulls all the calendar items starting with the BillingKeyword variable and makes a pivot table
' with a break-down by the "location" field, where the deliverables should be
' listed (instead of location).

' Pre-requisites are that you add the "Tools|References|Microsoft Excel 14.0 Object Library" under MS VBA in Outlook
' Suggested: Use conditional formatting under Outlook Calendar to highlight items with subject containing the BillingKeyword


Dim objView As CalendarView
Dim varArray As Variant
Dim FromDate As Date
Dim ToDate As Date
Dim BillingKeyword As String

' Change this value to the prefix you'll use in the subject field of appointments
BillingKeyword = "@Code:"

' Check if the current view is a calendar view.
If Application.ActiveExplorer.CurrentView.ViewType = olCalendarView Then
  ' Obtain a CalendarView object reference for the
  ' current calendar view.
   Set objView = Application.ActiveExplorer.CurrentView
   ' Obtain the DisplayedDates value, a string
   ' array of dates representing the dates displayed
   ' in the calendar view.
   varArray = objView.DisplayedDates
   ' If the example obtained a valid array, display
   ' a dialog box with a summary of its contents.
   If IsArray(varArray) Then
     FromDate = varArray(LBound(varArray))
     ToDate = varArray(UBound(varArray))
     Call GetCalData(BillingKeyword, FromDate, ToDate)
   End If
End If
End Sub
Private Sub GetCalData(BillingKeyword As String, StartDate As Date, Optional EndDate As Date)

' -------------------------------------------------
' Notes:
' If Outlook is not open, it still works, but much slower (~8 secs vs. 2 secs w/ Outlook open).
' Make sure to reference the Outlook and Excel object libraries (v 12.0) before running the code
' End Date is optional, if you want to pull from only one day, use: Call GetCalData("7/14/2008")
' -------------------------------------------------

Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Dim myCalItems As Outlook.Items
Dim ItemstoCheck As Outlook.Items
Dim ThisAppt As Outlook.AppointmentItem

Dim MyItem As Object

Dim StringToCheck As String

Dim MyBook As Excel.Workbook
Dim rngStart As Excel.Range

Dim i As Long
Dim NextRow As Long

Dim VOfficeLanguage As Long

VOfficeLanguage = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
   
dummy = EndDate - StartDate
If EndDate - StartDate > 28 Then
  ' ask if the requestor wants so much info
  If MsgBox("This could take some time (more than one month to process). Continue anyway?", vbInformation + vbYesNo) = vbNo Then
      GoTo ExitProc
  End If
End If

' get or create Outlook object and make sure it exists before continuing
On Error Resume Next
  Set olApp = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Set olApp = CreateObject("Outlook.Application")
  End If
On Error GoTo 0
If olApp Is Nothing Then
  MsgBox "Cannot start Outlook.", vbExclamation
  GoTo ExitProc
End If

Set olNS = olApp.GetNamespace("MAPI")
Set myCalItems = Application.ActiveExplorer.CurrentFolder.Items
' olNS.GetDefaultFolder(olFolderCalendar).Items

' ------------------------------------------------------------------
' the following code adapted from:
' http://www.outlookcode.com/article.aspx?id=30
'
With myCalItems
  .Sort "[Start]", False
  .IncludeRecurrences = True
End With
'
StringToCheck = "[Start] >= " & Quote(StartDate & " 12:00 AM") & " AND [End] <= " & Quote(EndDate & " 11:59 PM" & Chr(34) & " AND [Subject] >= " & Chr(34) & "@Code:" & Chr(34) & " AND [Subject] < " & Chr(34) & "@Code;")
Debug.Print StringToCheck
'
Set ItemstoCheck = myCalItems.Restrict(StringToCheck)
Debug.Print ItemstoCheck.Count
' ------------------------------------------------------------------

If ItemstoCheck.Count > 0 Then
  ' we found at least one appt
  ' check if there are actually any items in the collection, otherwise exit
  If ItemstoCheck.Item(1) Is Nothing Then GoTo ExitProc

  Set MyBook = Excel.Workbooks.Add
  ' Excel.Application.Visible = True
  Set rngStart = MyBook.Sheets(1).Range("A1")

  With rngStart
    .Offset(0, 0).Value = "Subject"
    .Offset(0, 1).Value = "Duration"
    .Offset(0, 2).Value = "Start Date"
    .Offset(0, 3).Value = "Start Time"
    .Offset(0, 4).Value = "End Date"
    .Offset(0, 5).Value = "End Time"
    .Offset(0, 6).Value = "Location"
    .Offset(0, 7).Value = "Categories"
  End With

  For Each MyItem In ItemstoCheck
    If MyItem.Class = olAppointment Then
    ' MyItem is the appointment or meeting item we want,
    ' set obj reference to it
      Set ThisAppt = MyItem
      NextRow = WorksheetFunction.CountA(Range("A:A"))

      With rngStart
        .End(xlDown).End(xlUp).Offset(NextRow, 0).Value = ThisAppt.Subject
        .End(xlDown).End(xlUp).Offset(NextRow, 1).Value = Format(ThisAppt.Duration / 1440, "HH:MM")
        .End(xlDown).End(xlUp).Offset(NextRow, 2).Value = Format(ThisAppt.Start, "MM/DD/YYYY")
        .End(xlDown).End(xlUp).Offset(NextRow, 3).Value = Format(ThisAppt.Start, "HH:MM AM/PM")
        .End(xlDown).End(xlUp).Offset(NextRow, 4).Value = Format(ThisAppt.End, "MM/DD/YYYY")
        .End(xlDown).End(xlUp).Offset(NextRow, 5).Value = Format(ThisAppt.End, "HH:MM AM/PM")
        .End(xlDown).End(xlUp).Offset(NextRow, 6).Value = ThisAppt.Location

        If ThisAppt.Categories <> "" Then
          .End(xlDown).End(xlUp).Offset(NextRow, 7).Value = ThisAppt.Categories
        Else
          .End(xlDown).End(xlUp).Offset(NextRow, 7).Value = "n/a"
        End If
      End With
    End If
  Next MyItem

'  In cell I2=IF(ISNUMBER(SEARCH("Graphe:",A2)),IF(ISNUMBER(FIND(" ",A2,9)),LEFT(A2,(FIND(" ",A2,9))),LEFT(A2,LEN(A2))),"")
'  Range("I2").Select
'  ActiveCell.FormulaR1C1 = "=IF(ISNUMBER(SEARCH(""Graphe:"",RC[-8])),IF(ISNUMBER(FIND("" "",RC[-8],9)),LEFT(RC[-8],(FIND("" "",RC[-8],9))),LEFT(RC[-8],LEN(RC[-8]))),"""")"
'  Range("I2").Select
'  Selection.AutoFill Destination:=Range("I2:I36")
'  Range("I2:I36").Select
'  Columns("I:I").EntireColumn.AutoFit

  ' make it pretty
  Call Cool_Colors(rngStart)

  ' Add Pivot Table
    Columns("A:H").Select
    Set VPivotTableTargetSheet = Sheets.Add
    ' Set VSourceData = ActiveSheet.Cells (to be used as a variable assigned to SourceData, not functional)
    Select Case VOfficeLanguage
        Case 1033, 4105 ' English US or Canada
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                "Sheet1!R1C1:R1048576C8").CreatePivotTable _
                TableDestination:="Sheet2!R3C1", TableName:="PivotTableTS"
'            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                "Sheet1!R1C1:R1048576C8", Version:=xlPivotTableVersion16).CreatePivotTable _
                TableDestination:=VPivotTableTargetSheet.Cell(R3C1), TableName:="PivotTableTS"
            VPivotTableTargetSheet.Select
        Case 1036 ' Français
            ' Correctif pour langue française
            ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                "Feuil1!R1C1:R1048576C8").CreatePivotTable _
                TableDestination:="Feuil2!R3C1", TableName:="PivotTableTS"
            VPivotTableTargetSheet.Select
    End Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTableTS").PivotFields("Subject")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTableTS").AddDataField ActiveSheet.PivotTables( _
        "PivotTableTS").PivotFields("Duration"), "Count of Duration", xlCount
    With ActiveSheet.PivotTables("PivotTableTS").PivotFields("Start Date")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableTS").PivotFields("Count of Duration")
        .Caption = "Sum of Duration"
        .Function = xlSum
        .NumberFormat = "[H]:mm;;"
    End With
        With ActiveSheet.PivotTables("PivotTableTS").PivotFields("Location")
        .Orientation = xlRowField
        .Position = 2
    End With
    Range("E5").Select
    ActiveSheet.PivotTables("PivotTableTS").AllowMultipleFilters = True
Rem    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Subject").PivotFilters.Add Type:=xlValueDoesNotEqual, DataField:=ActiveSheet.PivotTables("PivotTableTS").PivotFields("Temps"), Value1:=0
    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Subject").PivotFilters.Add Type:=xlCaptionContains, Value1:=BillingKeyword
    Range("B4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Range("B1").Activate
    Cells.EntireColumn.AutoFit
    Range("A4").Select
'   Collapse all location details
    Range("B5").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Subject").ShowDetail = False

    Selection.End(xlToRight).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Cells.FormatConditions.Delete
'    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1"
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    Selection.FormatConditions(Selection.FormatConditions.Count).NumberFormat = "d""jour(s) de 7.5 h, +"" h""h""mm""m"""
'    Selection.FormatConditions(1).StopIfTrue = False
    Cells.EntireColumn.AutoFit
'    ActiveSheet.PivotTables("PivotTableTS").RowGrand = False
    Rows("4:4").Select
    Selection.NumberFormat = "ddd, d/m/yy;@"
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Cells.Select
    ActiveSheet.PivotTables("PivotTableTS").TableStyle2 = "PivotStyleMedium9"
    Range("B1").Select
    ActiveSheet.PivotTables("PivotTableTS").CompactLayoutColumnHeader = ""
    Columns("B:B").EntireColumn.AutoFit
    ActiveWorkbook.ShowPivotTableFieldList = False
        Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").ColumnWidth = 2.86
    ActiveSheet.PivotTables("PivotTableTS").PivotSelect "", xlDataAndLabel, True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDot
        .ThemeColor = 5
        .TintAndShade = 0.599963377788629
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ThemeColor = 5
        .TintAndShade = 0.599963377788629
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveSheet.PivotTables("PivotTableTS").PivotSelect "", xlDataAndLabel, True
    Range("B2").Select
    ActiveSheet.PivotTables("PivotTableTS").DataPivotField.PivotItems("Sum of Duration").Caption = "Temps"
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTableTS").CompactLayoutRowHeader = "Graphes & activités"
    Range("B4").Select
    ActiveSheet.PivotTables("PivotTableTS").ShowTableStyleRowStripes = True
    ActiveSheet.PivotTables("PivotTableTS").TableStyle2 = "PivotStyleMedium9"
        Range("B2").Select
            ' Add Total Duration * 24 (to see in hours + fraction)
                ActiveSheet.PivotTables("PivotTableTS").PivotSelect "'Row Grand Total'", _
                xlDataAndLabel, True
            ActiveSheet.PivotTables("PivotTableTS").CalculatedFields.Add "Dec", "=Duration*24", True
            ActiveSheet.PivotTables("PivotTableTS").PivotFields("Dec").Orientation = xlDataField
'                    With ActiveSheet.PivotTables("PivotTableTS").PivotFields("Dec")
'                        .NumberFormat = "0.00"
'                    End With
   
            Select Case VOfficeLanguage
                Case 1033, 4105 ' English US or Canada
                    ActiveSheet.PivotTables("PivotTableTS").PivotSelect "'Sum of Dec'", xlDataAndLabel, True
                    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Sum of Dec").Caption = "T"
               Case 1036 ' Français
                    ' Correctif de langue française
                    ActiveSheet.PivotTables("PivotTableTS").PivotSelect "'Somme de Dec'", xlDataAndLabel, True
                    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Somme de Dec").Caption = "T"
            End Select
                        Selection.NumberFormat = "#.00;[Red]-#,###;"
                    With Selection.Font
                        .Name = "Calibri"
                        .Size = 7
                '        .Strikethrough = False
                '        .Superscript = False
                '        .Subscript = False
                '        .OutlineFont = False
                '        .Shadow = False
                '        .Underline = xlUnderlineStyleNone
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                '        .ThemeFont = xlThemeFontMinor
                    End With
                    With Selection
                        .HorizontalAlignment = xlRight
                    End With
                    Cells.Select
                    Cells.EntireColumn.AutoFit

                    ActiveSheet.PivotTables("PivotTableTS").PivotSelect "T", xlDataAndLabel, True
                    With Selection.Font
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0
                    End With
                    ActiveSheet.PivotTables("PivotTableTS").PivotSelect "T 'Row Grand Total'", _
                        xlDataAndLabel, True
                    With Selection.Font
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0
                    End With
                    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Subject").AutoSort xlAscending, "Temps"

                    Selection.Font.Bold = True
                    Selection.Font.Size = 12
                    With Selection.Font
                        .Name = "Calibri"
                        .Size = 11
                        .Strikethrough = False
                        .Superscript = False
                        .Subscript = False
                        .OutlineFont = False
                        .Shadow = False
                        .Underline = xlUnderlineStyleNone
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0
                        .ThemeFont = xlThemeFontMinor
                    End With
                    Columns("N:N").EntireColumn.AutoFit

    ' Conditional formatting

    Columns("B:B").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="*-2820", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'   With Selection.FormatConditions(1).Font
'       .ColorIndex = xlAutomatic
'       .TintAndShade = 0
'   End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        .Color = &HFA93FA
        ' http://www.colorpicker.com/ (invert the three Hex couples, as it's in reverse order)
        ' http://www.mathsisfun.com/binary-decimal-hexadecimal-converter.html
        ' .ColorIndex = 5
        ' .Color = RGB(200, 250, 200)
        ' .Color = &HC8EFAC8   'h=Hex,  o=Octal  anyone still use octal
        ' .Color = 14403539
        ' 10040319
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.80001220740379
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlTextString, String:="*:Oper", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        .Color = 12764108
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.80001220740379
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlTextString, String:="*-2830", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 0
        .Gradient.ColorStops.Clear
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        .Color = &HECFFCC
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.80001220740379
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    ' Font and color formatting
    Range("B5").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Font
        .Name = "SansSerif"
        .Size = 9
    End With

    Excel.Application.ActiveWindow.DisplayGridlines = False
    ' Excel.Application.ActiveWindow.DisplayHeadings = False
    ActiveSheet.PivotTables("PivotTableTS").ShowTableStyleRowStripes = False

    ' Add Emptoris format info at the end of the line
'    Range("B3").Select
'    Selection.End(xlDown).Select
'    Selection.End(xlDown).Select
'    Selection.End(xlToRight).Select
'    Selection.End(xlUp).Select
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.FormulaR1C1 = "=MID(RC[-13],9,11)"
    ''' Selection.AutoFill Destination:=Range("O5:O100")
    ''' Range("O5:O17").Select
'    Selection.End(xlDown).Select
'    Selection.ClearContents

'    Range("B3").Select
'    Selection.End(xlDown).Select
'    Selection.End(xlDown).Select
'    Selection.End(xlToRight).Select
'    Selection.End(xlUp).Select
'    ActiveCell.Offset(0, 2).Select
    ' Range("P5").Select
'    ActiveCell.FormulaR1C1 = "=MID(RC[-14],22,LEN(RC[-14])-21)"
    ' Selection.AutoFill Destination:=Range("P5:P100")
'    Range("P5:P100").Select
'    Selection.End(xlDown).Select
'    Selection.ClearContents

' Add Emptoris formated project info in hidden columns before the table
    ' Add two columns before the Pivot Table
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' Add formulas
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[1],23,LEN(RC[1])-21)"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[2],9,11)"
    ' Select and autofill
    ' Set SourceRange = Range("B5:C5") ----> was causing the script to stop on Ali's computer
    ' Set fillRange = Range("B5", Range("D" & Rows.Count).End(xlUp).Offset(-1, -1)) ----> was causing the script to stop on Ali's computer
    Set SourceRange = Range("B2:C2")
    Set fillRange = Range("B2", Range("D" & Rows.Count).End(xlUp).Offset(-1, -1))
    SourceRange.AutoFill Destination:=fillRange
    ' Format and hide the new columns
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:C").Select
    Selection.Columns.Group
    Columns(3).ShowDetail = False

    Range("D3").Select
    ActiveSheet.PivotTables("PivotTableTS").PivotFields("Subject").AutoSort xlAscending, "Subject"
    ' Show Excel
    Excel.Application.Visible = True
    ' End Excel script

Else
    MsgBox "There are no appointments or meetings during" & _
      "the time you specified. Exiting now.", vbCritical
End If

ExitProc:
Set myCalItems = Nothing
Set ItemstoCheck = Nothing
Set olNS = Nothing
Set olApp = Nothing
Set rngStart = Nothing
Set ThisAppt = Nothing
Set MyBook = Nothing
End Sub
Private Function Quote(myText)
' from Sue Mosher's excellent book "Microsoft Outlook Programming"
  Quote = Chr(34) & myText & Chr(34)
End Function
Private Sub Cool_Colors(rng As Excel.Range)
'
' Lt Blue BG with white letters
'
'
With Range(rng, rng.End(xlToRight))
  .Font.ColorIndex = 2
  .Font.Bold = True
  .HorizontalAlignment = xlCenter
  .MergeCells = False
  .AutoFilter
  .CurrentRegion.Columns.AutoFit
  With .Interior
    .ColorIndex = 41
    .Pattern = xlSolid
  End With
End With

End Sub

Function RegExTest(rngCell As Range, strPattern As String) As Boolean
    Dim objRegex As Object
    Set objRegex = CreateObject("VBScript.RegExp")
    objRegex.Global = True
    objRegex.IgnoreCase = True
    objRegex.Pattern = strPattern
    RegExTest = objRegex.Test(rngCell.Value)
End Function

Sub ApptFindReplace()
' This macro will find/replace specified text in the subject line of all appointments in a specified calendar

    Dim olApp As Outlook.Application
    Dim CalFolder As Outlook.MAPIFolder
    Dim Appt As Outlook.AppointmentItem
    Dim OldText As String
    Dim newText As String
    Dim CalChangedCount As Integer

    ' Set Outlook as active application
    Set olApp = Outlook.Application

    ' Get user inputs for find text, replace text and calendar folder
    MsgBox ("This script will perform a find/replace in the subject line of all appointments in a specified calendar.")
    OldText = InputBox("What is the text string that you would like to replace?")
    newText = InputBox("With what would you like to replace it?")
    MsgBox ("In the following dialog box, please select the calendar you would like to use.")
    Set CalFolder = Application.Session.PickFolder
    On Error GoTo ErrHandler:

    ' Check to be sure a Calendar folder was selected
    Do
    If CalFolder.DefaultItemType <> olAppointmentItem Then
        MsgBox ("This macro only works on calendar folders.  Please select a calendar folder from the following list.")
        Set CalFolder = Application.Session.PickFolder
        On Error GoTo ErrHandler:
    End If
    Loop Until CalFolder.DefaultItemType = olAppointmentItem

    ' Loop through appointments in calendar, change text where necessary, keep count
    CalChangedCount = 0
    For Each Appt In CalFolder.Items
        If InStr(Appt.Subject, OldText) <> 0 Then
            Debug.Print "Changed: " & Appt.Subject & " - " & Appt.Start
            Appt.Subject = Replace(Appt.Subject, OldText, newText)
            Appt.Save
       CalChangedCount = CalChangedCount + 1
       End If
    Next

    ' Display results and clear table
    MsgBox (CalChangedCount & " appointments had text in their subjects changed from '" & OldText & "' to '" & newText & "'.")
    Set Appt = Nothing
    Set CalFolder = Nothing
    Exit Sub

ErrHandler:
    MsgBox ("Macro terminated.")
    Exit Sub

End Sub

Sub ApptSetBusyAsFree()
' This macro will set the Free/Busy status to Free in a specified calendar
' Used source from: https://www.slipstick.com/developer/code-samples/change-the-all-day-event-default-freebusy-to-busy/

    Dim olApp As Outlook.Application
    Dim CalFolder As Outlook.MAPIFolder
    Dim Appt As Outlook.AppointmentItem
    Dim CalChangedCount As Integer
   
    BillingKeyword = "@Code:"
   
    ' Set Outlook as active application
    Set olApp = Outlook.Application

    ' Get user inputs for find text, replace text and calendar folder
    ' MsgBox ("In the following dialog box, please select the calendar you would like to use.")
    ' Set CalFolder = Application.Session.PickFolder
    Set CalFolder = Application.ActiveExplorer.CurrentFolder
    On Error GoTo ErrHandler:

    ' Check to be sure a Calendar folder was selected
    Do
        If CalFolder.DefaultItemType <> olAppointmentItem Then
            MsgBox ("This macro only works on calendar folders.  Please select a calendar folder from the following list.")
            Set CalFolder = Application.Session.PickFolder
            On Error GoTo ErrHandler:
        End If
    Loop Until CalFolder.DefaultItemType = olAppointmentItem

    ' Loop through appointments in calendar, change text where necessary, keep count
    CalChangedCount = 0
    For Each Appt In CalFolder.Items
        If InStr(Appt.Subject, BillingKeyword) <> 0 Then
            Debug.Print "Changed: " & Appt.Subject & " - " & Appt.Start
            If Appt.BusyStatus <> olFree Then
                Appt.BusyStatus = olFree
                Appt.Save
                CalChangedCount = CalChangedCount + 1
            End If
       End If
    Next

    ' Display results and clear table
    MsgBox (CalChangedCount & " appointments had their Free/Busy stataus set to Free.")
    Set Appt = Nothing
    Set CalFolder = Nothing
    Exit Sub

ErrHandler:
    MsgBox ("Macro terminated.")
    Exit Sub

End Sub

 
' Sub routine to find what language Office is set to
' Source: https://stackoverflow.com/questions/8588728/find-the-current-user-language
' At CGI, french was 1036 and english 1033
' Lookup (convert from Hex to Dec: https://msdn.microsoft.com/en-us/library/cc233982.aspx

    Sub GetXlLang()
        Dim lngCode As Long
        lngCode = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
        MsgBox "Code is: " & lngCode & vbNewLine
        ' & GetTxt(lngCode)
    End Sub

' Function to match language code with language (referenced link not functional anymore)
' Source: https://stackoverflow.com/questions/8588728/find-the-current-user-language

    Function GetTxt(ByVal lngCode) As String
        Dim objXmlHTTP As Object
        Dim objRegex As Object
        Dim objRegMC As Object
        Dim strResponse As String
        Dim strSite As String

        Set objXmlHTTP = CreateObject("MSXML2.XMLHTTP")
        strSite = "https://msdn.microsoft.com/en-us/library/cc233982.aspx"
        ' Original link (redirected to the above, which does not work with the marco
        ' strSite = "http://msdn.microsoft.com/en-us/goglobal/bb964664"

        On Error GoTo ErrHandler
        With objXmlHTTP
            .Open "GET", strSite, False
            .Send
            If .Status = 200 Then strResponse = .ResponseText
        End With
        On Error GoTo 0

        strResponse = Replace(strResponse, "</td><td>", vbNullString)
        Set objRegex = CreateObject("vbscript.regexp")
        With objRegex
            .Pattern = "><td>([a-zA-Z- ]+)[A-Fa-f0-9]{4}" & lngCode
            If .Test(strResponse) Then
                Set objRegMC = .Execute(strResponse)
                GetTxt = objRegMC(0).submatches(0)
            Else
                GetTxt = "Value not found from " & strSite
            End If
        End With
        Set objRegex = Nothing
        Set objXmlHTTP = Nothing
        Exit Function
ErrHandler:
        If Not objXmlHTTP Is Nothing Then Set objXmlHTTP = Nothing
        GetTxt = strSite & " unable to be accessed"
    End Function
