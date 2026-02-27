VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLookup 
   Caption         =   "Quick Analysis"
   ClientHeight    =   90
   ClientLeft      =   -906
   ClientTop       =   -3972
   ClientWidth     =   120
   OleObjectBlob   =   "Quick Analysis - User Form IndexMatch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private formLocked As Boolean

Private Sub UnlockForm()
    
    formLocked = False

    divisions.Enabled = True
    categories.Enabled = True

    q1.Enabled = True
    q2.Enabled = True
    q3.Enabled = True
    q4.Enabled = True

    sum.Enabled = True
    avg.Enabled = True
    stdDev.Enabled = True

    btnLookup.Enabled = False
End Sub

Private Sub LockForm()
    formLocked = True

    divisions.Enabled = False
    categories.Enabled = False

    q1.Enabled = False
    q2.Enabled = False
    q3.Enabled = False
    q4.Enabled = False

    sum.Enabled = False
    avg.Enabled = False
    stdDev.Enabled = False

    btnLookup.Enabled = False
End Sub

Private Sub backtoMain_Click()
    
    Me.Hide
    frmMain.Show
    
End Sub

Private Sub btnClearSel_Click()
    divisions.Value = ""
    categories.Value = ""

    q1.Value = False
    q2.Value = False
    q3.Value = False
    q4.Value = False

    sum.Value = False
    avg.Value = False
    stdDev.Value = False

    ClearResult
    UnlockForm

End Sub

Private Sub divisionInstruct_Click()

End Sub

Private Sub results_Click()

End Sub
Private Sub ClearResult()
    With resultsbox
        .Value = ""
        .BackColor = vbWhite
        .TextAlign = fmTextAlignLeft
    End With
End Sub

Private Sub resultsbox_Change()
    With resultsbox
        .Locked = True
        .BackColor = &HF0F0F0
        .TextAlign = fmTextAlignCenter
    End With
    
End Sub
' iniatialize dropdowns
' labels of division dropdown is divisions
' the array for divisions is divs
' labels of category dropdown is categories
' the array for categories is cats
Private Sub UserForm_Initialize()
    formLocked = False
    Me.Width = 500
    Me.Height = 440

    Dim divs As Variant
    Dim cats As Variant
    Dim i As Long

    divs = Array("East", "West", "North", "South")

    cats = Array( _
        "Overhead", "Technical Support", "Telephone", "Maintenance", _
        "Supplies", "Software", "Copying", "Contractors", _
        "Rent", "Consultants", "Telemarketing", "Advertising", _
        "Miscellaneous", "Salaries", "Clerical Support")

    divisions.Clear
    categories.Clear

    For i = LBound(divs) To UBound(divs)
        divisions.AddItem divs(i)
    Next i

    For i = LBound(cats) To UBound(cats)
        categories.AddItem cats(i)
    Next i

    ' Disable Run Lookup at start
    btnLookup.Enabled = False

End Sub
' validating form before the lookup button can be clicked
Private Function IsFormValid() As Boolean

    ' division + category selected
    If divisions.Value = "" Then Exit Function
    If categories.Value = "" Then Exit Function

    ' at least one quarter selected
    If Not (q1.Value Or q2.Value Or q3.Value Or q4.Value) Then Exit Function

    ' KPI selected
    If Not (sum.Value Or avg.Value Or stdDev.Value) Then Exit Function

    IsFormValid = True

End Function

Private Function ValidateQuarter(q As Integer) As Boolean

    Dim resp As VbMsgBoxResult
    
    Dim sheetCount As Integer
    
    Dim n As Long
    
    If SheetExists(GetQuarterSheetName(q)) Then
        ValidateQuarter = True
        Exit Function
    End If

    resp = MsgBox( _
        "Quarter " & q & " report does not exist." & vbCrLf & _
        "Do you want to generate it now?", _
        vbYesNo + vbQuestion, _
        "Missing Quarter")

    If resp = vbYes Then
    
        sheetCount = Application.InputBox( _
            Prompt:="Enter number of raw data sheets to populate:", _
            Title:="Populate Data", _
            Type:=1)
    
        ' Validate
        If sheetCount < 1 Or sheetCount = 0 Then
            MsgBox "Please enter a number greater than 0.", vbExclamation
            Exit Function
        End If
    
        n = CLng(sheetCount)
    
        CreateSampleData n
        ' RemoveDefaultSheetIfEmpty

        LoopQReport q
        ValidateQuarter = True
    Else
        ValidateQuarter = False
    End If

End Function

Private Sub btnLookup_Click()
    
    'clear once new results are on the way, when user presses button it means they are ready to get new value
    'if i clear after every selection its too many lines unneeded and user might still want it until the last moment
    ClearResult

    Dim quarters As Object
    Set quarters = CreateObject("Scripting.Dictionary")

    If q1.Value Then quarters.Add "Quarter 1 Expenses", True
    If q2.Value Then quarters.Add "Quarter 2 Expenses", True
    If q3.Value Then quarters.Add "Quarter 3 Expenses", True
    If q4.Value Then quarters.Add "Quarter 4 Expenses", True

    If quarters.count = 0 Then
        MsgBox "Select at least one quarter.", vbExclamation
        Exit Sub
    End If

    Dim resultType As String
    If sum.Value Then resultType = "SUM"
    If avg.Value Then resultType = "AVG"
    If stdDev.Value Then resultType = "STD"

    Dim result As Variant

    result = LookupAggregate_IndexMatch( _
        divisions.Value, _
        categories.Value, _
        quarters, _
        resultType)
        
    If IsError(result) Then

        If sum.Value Or avg.Value Then
            resultsbox.Value = "No matching data"
        ElseIf stdDev.Value Then
            MsgBox _
                "Standard deviation requires at least TWO data points." & vbCrLf & _
                "Please select two or more quarters.", _
                vbExclamation, _
                "Insufficient Data"

            stdDev.Value = False
            resultsbox.Value = ""
            UnlockForm
            Exit Sub
        End If

    Else
        resultsbox.Value = Format(result, "0.00")
        LockForm
    End If

End Sub

Private Sub categories_Change()
    If formLocked Then Exit Sub

    btnLookup.Enabled = IsFormValid
End Sub

Private Sub divisions_Change()
    If formLocked Then Exit Sub

    btnLookup.Enabled = IsFormValid
End Sub

Private Sub InstructionBox_Click()

End Sub

Private Sub q1_Click()
    If formLocked Then Exit Sub

    If q1.Value Then
        If Not ValidateQuarter(1) Then q1.Value = False
    End If
    btnLookup.Enabled = IsFormValid
End Sub

Private Sub q2_Click()
    If formLocked Then Exit Sub

    If q2.Value Then
        If Not ValidateQuarter(2) Then q2.Value = False
    End If
    btnLookup.Enabled = IsFormValid
End Sub

Private Sub q3_Click()
    If formLocked Then Exit Sub

    If q3.Value Then
        If Not ValidateQuarter(3) Then q3.Value = False
    End If
    btnLookup.Enabled = IsFormValid

End Sub

Private Sub q4_Click()
    If formLocked Then Exit Sub

    If q4.Value Then
        If Not ValidateQuarter(4) Then q4.Value = False
    End If
    btnLookup.Enabled = IsFormValid
End Sub

Private Sub stdDev_Click()
    If formLocked Then Exit Sub

    btnLookup.Enabled = IsFormValid

End Sub

Private Sub sum_Click()
    If formLocked Then Exit Sub

    btnLookup.Enabled = IsFormValid
End Sub

Private Sub avg_Click()
    If formLocked Then Exit Sub
    
    btnLookup.Enabled = IsFormValid

End Sub

Private Sub UserForm_Click()

End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub



