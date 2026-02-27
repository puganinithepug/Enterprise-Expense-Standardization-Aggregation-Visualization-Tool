Attribute VB_Name = "Module1"
Sub LaunchApp()
    frmMain.Show
End Sub
Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' function to assemble name of total quarter expenses sheet
Function GetQuarterSheetName(q As Integer) As String
    GetQuarterSheetName = "Quarter " & q & " Expenses"
End Function

' function to assemble name of total quarter expenses sheet
Function GetRawSheetName(q As Integer) As String
    GetRawSheetName = "RawData Quarter " & q & " Expenses"
End Function
' delete "Sheet"
Sub DeleteDefaultSheets()
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    
    For Each ws In Worksheets
        If ws.Name Like "Sheet*" Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
End Sub

' for random sample data population
Sub CreateSampleData(sheetCount As Long)

    Dim ws As Worksheet
    Dim i As Long
    Dim sheetIndex As Long

    Dim divisions As Variant
    Dim categories As Variant

    divisions = Array("East", "West", "North", "South")
    categories = Array( _
        "Overhead", "Technical Support", "Telephone", "Maintenance", _
        "Supplies", "Software", "Copying", "Contractors", _
        "Rent", "Consultants", "Telemarketing", "Advertising", _
        "Miscellaneous", "Salaries", "Clerical Support")

    Randomize

    ' create n raw data sheets
    For sheetIndex = 1 To sheetCount
    
        ' no overwriting

        Set ws = CreateNewSheet("RawSheet_" & sheetIndex)

        ' headers (generic, not quarter-based)
        ws.Range("A1:F1").Value = Array("Division", "Category", "Val1", "Val2", "Val3", "Total")

        For i = 2 To 21
            ws.Cells(i, 1).Value = divisions(Int(Rnd * 4))
            ws.Cells(i, 2).Value = categories(Int(Rnd * (UBound(categories) + 1)))
            ws.Cells(i, 3).Value = Round(Rnd * 5000 + 200, 2)
            ws.Cells(i, 4).Value = Round(Rnd * 5000 + 200, 2)
            ws.Cells(i, 5).Value = Round(Rnd * 5000 + 200, 2)
            ws.Cells(i, 6).Formula = "=SUM(C" & i & ":E" & i & ")"
        Next i

    Next sheetIndex

End Sub
Sub DeleteQuarterSheets(q As Integer)

    Dim ws As Worksheet
    Dim qName As String

    Application.DisplayAlerts = False

    qName = "Quarter " & q & " Expenses"

    For Each ws In Worksheets
        If ws.Name = qName _
           Or ws.Name Like "RawData Quarter " & q & " Expenses*" Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True

End Sub
Function CreateNewSheet(baseName As String) As Worksheet
    Dim ws As Worksheet
    Dim nameTry As String
    Dim i As Long

    i = 1
    nameTry = baseName

    Do While SheetExists(nameTry)
        i = i + 1
        nameTry = baseName & "_" & i
    Loop

    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
    ws.Name = nameTry
    Set CreateNewSheet = ws
End Function

' new sheet creation process
Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.count))
        ws.Name = sheetName
    End If
    
    Set GetOrCreateSheet = ws
End Function
' check for headers
Function HasHeader(ws As Worksheet) As Boolean
    HasHeader = _
        ws.Range("A1").Value = "Division" And _
        ws.Range("B1").Value = "Category" And _
        ws.Range("F1").Value = "Total"
End Function

Sub LoopQReport(q As Integer)
    Dim ws As Worksheet
    Dim first As Boolean
    
    Dim i As Long
    i = 1
    
    ' Dim q As Integer
    Dim qSheet As Worksheet
    
    ' result for user input on for same quarter gen
    Dim resp As VbMsgBoxResult

    ' q = GetQuarterFromUser
    ' If q = 0 Then Exit Sub
    
    ' if the stuff for that quarter already there
    If SheetExists(GetQuarterSheetName(q)) Then
        MsgBox "That quarter already has a report. vbInformation"

        If SheetExists(GetQuarterSheetName(q)) Then
            resp = MsgBox( _
                "A report for Quarter " & q & " already exists." & vbCrLf & vbCrLf & _
                "(Existing quarter sheets will be replaced)", _
                vbYesNo + vbQuestion, _
                "Repopulate Quarter?" _
            )

            If resp = vbNo Then Exit Sub

            ' on yes
            DeleteQuarterSheets q
        End If
    End If
    
    Set qSheet = CreateNewSheet(GetQuarterSheetName(q))
    
    first = True
    
    DeleteDefaultSheets
    
    For Each ws In Worksheets
        
        Worksheets(ws.Name).Select
        
        ' rename worksheets to include the quarter for which they were created
        
        If ws.Name <> qSheet.Name And ws.Name Like "RawSheet*" Then
            EnsureHeaders q
            AutomateTotalSum
            
            ws.Name = GetRawSheetName(q) & "_" & i
            i = i + 1
            
            ' select current data
            Range("A2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Range(Selection, Selection.End(xlToRight)).Select
            
            ' copy data
            Selection.Copy
            
            ' select quarter 1 report
            qSheet.Select
            
            ' paste data
            Range("A30000").Select
            Selection.End(xlUp).Select
            
            If Not first Then
                ActiveCell.Offset(1, 0).Select
            Else
                first = False
            End If
            
            ActiveSheet.Paste
        End If
    
    
    ' move to the next sheet in the loop

    Next ws
    
    qSheet.Select
    EnsureHeaders q
    AutomateTotalSum
    
    ' create the composite key
    BuildCompositeKey qSheet
    
    CreatePivotForQuarter q
    
End Sub
Public Sub AutomateTotalSum()
    Dim ws As Worksheet
    Dim lastCell As String
    
    'For Each ws in Worksheets
    'Worksheets(ws.Name).Select
        
    Range("F2").Select
        'now find bottom
    Cells(Rows.count, "F").End(xlUp).Select
        ' the cntrl down to get total coln
        ' assume total always in F coln
    
    If ActiveCell.Row < 2 Then Exit Sub
        
    If InStr(1, Selection.Formula, "SUM(", vbTextCompare) > 0 Then
        Exit Sub
    End If
        
        'get last cell to refernce
        
    lastCell = ActiveCell.Address(False, False)
        
        'make last cell be activecell's address
        ' if dont add the (false, false), will be absolute refernece
        ' add falses to make it relative so the address is dynamic
    ActiveCell.Offset(1, 0).Select
        
        'select 1 cell down from offset
    ActiveCell.Value = "=sum(F2:" & lastCell & ")"
        
        
        'creating sum function
        'go to very bottom of sheet
        'go xlup
        'move a row down below first cell in data
        'Range("A30000").Select
        'Selection.End(xlUp).Select
        'ActiveCell.Offset(1, 0).Select
        'ActiveSheet.Paste
        
    'Next ws
 
End Sub

Sub InsertheadersByQ(q As Integer)
'
' HeadersFormat Macro
' Headers and format added.
'
' Keyboard Shortcut: Ctrl+k
    
    Dim months As Variant
    
    Select Case q
        Case 1
            months = Array("Jan", "Feb", "Mar")
        Case 2
            months = Array("Apr", "May", "Jun")
        Case 3
            months = Array("Jul", "Aug", "Sep")
        Case 4
            months = Array("Oct", "Nov", "Dec")
    End Select
    
    ' Note to self - recorded Macro style
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Division"
    
    ' written VBA style
    Range("B1").Value = "Category"
    
    Range("C1").Value = months(0)
    
    Range("D1").Value = months(1)
    
    Range("E1").Value = months(2)
    
    Range("F1").Value = "Total"
    
    Range("A2").Select
End Sub
Sub FormatHeaders()
'
' HeaderFormatCont Macro
' More
'
' Keyboard Shortcut: Ctrl+m
'
    Range("A1:F1").Select
    Selection.Font.Size = 12
    Selection.Font.Size = 14
    Selection.Font.Size = 16
    Selection.Font.Size = 18
    Selection.Font.Size = 16
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    Range("A2").Select
    
End Sub

Sub EnsureHeaders(q As Integer)
    ' Check if headers already exist
    If Range("A1").Value = "Division" _
       And Range("B1").Value = "Category" _
       And Range("F1").Value = "Total" Then
        If Range("C1").Value = "Jan" And q = 1 _
            Or Range("C1").Value = "Apr" And q = 2 _
            Or Range("C1").Value = "Jul" And q = 3 _
            Or Range("C1").Value = "Oct" And q = 4 Then
                Exit Sub
        End If
        Rows(1).Clear
        InsertheadersByQ q
        FormatHeaders
        Exit Sub
    End If
    
    ' If not, insert them
    InsertheadersByQ q
    FormatHeaders
End Sub
' index match function
' this returns the total value
' so basically calls to this have to be captured and aggregated
Function IndexMatchTotal(ws As Worksheet, lookupKey As String) As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim total As Double
    Dim found As Boolean

    lastRow = ws.Cells(ws.Rows.count, "G").End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, "G").Value = lookupKey Then
            total = total + ws.Cells(i, "F").Value
            found = True
        End If
    Next i

    If Not found Then
        IndexMatchTotal = CVErr(xlErrNA)
    Else
        IndexMatchTotal = total
    End If
End Function
' Replacing INDEX/MATCH-related logic with aggregated division and category key-based lookup to get totals
' the division, category quarters, resultType are passed parameters from the user multiselects
Function LookupAggregate_IndexMatch( _
    div As String, _
    cat As String, _
    quarters As Object, _
    resultType As String) As Variant

    Dim ws As Worksheet
    Dim val As Variant
    
    ' this array stores values from index/matching on user's selected composite lookup key
    Dim values() As Double
    
    Dim count As Long
    
    ' this is the composite lookup key
    Dim key As String
    ' contructing the key
    
    key = div & "|" & cat
    
    count = 0

    For Each ws In Worksheets

        If quarters.Exists(ws.Name) Then

            val = IndexMatchTotal(ws, key)

            If Not IsError(val) Then
                count = count + 1
                ReDim Preserve values(1 To count)
                values(count) = val
            End If

        End If
    Next ws

    If count = 0 Then
        LookupAggregate_IndexMatch = CVErr(xlErrNA)
        Exit Function
    End If

    Select Case resultType
        Case "SUM"
            LookupAggregate_IndexMatch = Application.WorksheetFunction.sum(values)
        Case "AVG"
            LookupAggregate_IndexMatch = Application.WorksheetFunction.AVERAGE(values)
        
        Case "STD"
            If count < 2 Then
                LookupAggregate_IndexMatch = CVErr(xlErrDiv0)
            Else
                Dim tmp As Variant
                tmp = Application.StDev(values)

                If IsError(tmp) Then
                    LookupAggregate_IndexMatch = tmp
                Else
                    LookupAggregate_IndexMatch = tmp
                End If
            End If
    End Select

End Function

' build composite key for VLOOKUP -> upgrading to INDEX/MATCH to avoid copying "Total" column data since VLOOKUP
' index match also only matches just 1 lookup value so still need composite key

' vlookup requires that Lookup column must be first
' Cannot look left at Total column to lookup form composite key column, index match can look right and left  - better use

Sub BuildCompositeKey(ws As Worksheet)

    Dim lastRow As Long
    Dim i As Long

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ws.Range("G1").Value = "Key"

    For i = 2 To lastRow
        ws.Cells(i, "G").Value = _
            ws.Cells(i, "A").Value & "|" & ws.Cells(i, "B").Value
    Next i
    
    ' the composite key column is behind the scenes stuff
    ws.Columns("G").Hidden = True
    ws.Protect UserInterfaceOnly:=True

End Sub

Sub CreatePivotForQuarter(q As Integer)
    Dim wsData As Worksheet
    Dim wsPivotSheet As Worksheet
    Dim ptCache As PivotCache
    Dim pivotTable As pivotTable
    Dim lastRow As Long
    Dim pivotName As String
    
    ' Source data
    Set wsData = ThisWorkbook.Worksheets(GetQuarterSheetName(q))
    
    ' Determine last row
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row
    
    ' Create or clear Pivot sheet
    Set wsPivotSheet = CreateNewSheet("Pivot Table for Quarter " & q)
    
    ' Pivot cache
    Set ptCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1:F" & lastRow))
    
    ' Create pivot table
    pivotName = "Pivot_Quarter_" & q
    Set pivotTable = ptCache.CreatePivotTable( _
        TableDestination:=wsPivotSheet.Range("A3"), _
        TableName:=pivotName)
    
    ' Configure pivot table
    With pivotTable
        ' Row fields
        .PivotFields("Division").Orientation = xlRowField
        .PivotFields("Category").Orientation = xlRowField
        ' Add data field
        .AddDataField .PivotFields("Total"), "Sum of Total", xlSum
        ' Layout
        .RowAxisLayout xlTabularRow
    End With

    
    MsgBox "Pivot table created on sheet '" & wsPivotSheet.Name & "'", vbInformation
End Sub
