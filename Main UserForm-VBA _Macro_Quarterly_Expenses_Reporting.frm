VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Quarterly Expenses"
   ClientHeight    =   72
   ClientLeft      =   -228
   ClientTop       =   -1116
   ClientWidth     =   12
   OleObjectBlob   =   "Main UserForm-VBA _Macro_Quarterly_Expenses_Reporting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPopulate_Click()

    Dim sheetCount As Variant
    Dim n As Long

    sheetCount = Application.InputBox( _
        Prompt:="Enter number of raw data sheets to populate:", _
        Title:="Populate Data", _
        Type:=1)

    ' User cancelled
    If sheetCount = 0 Then Exit Sub

    ' Validate
    If sheetCount < 1 Then
        MsgBox "Please enter a number greater than 0.", vbExclamation
        Exit Sub
    End If

    n = CLng(sheetCount)

    CreateSampleData n
    ' RemoveDefaultSheetIfEmpty

    cboQuarter.Enabled = True
    MsgBox n & " raw data sheets created. Select a quarter to continue.", vbInformation

End Sub

Private Sub btnRefresh_Click()
    Dim ws As Worksheet
' button that creates Sheet1
' deletes all other sheets present
' basically a buttn to go back to default
    Sheets.Add.Name = "Sheet1"
    Application.DisplayAlerts = False ' Turns off confirmation prompts
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Sheet1" Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True ' Turns confirmation prompts back on
End Sub

Private Sub btnVlookup_Click()
    Me.Hide
    frmLookup.Show
    
End Sub

Private Sub cboQuarter_Change()
' this is second to be activated by user
' this is supposed to trigger data standardization
' can select quarter only after btn_populate has been clcked, so must check for presence of data before user can select quarter
' every logic related to quarter selection by user, user inout for quarter selection should be tied to this button
' when quarter is selected with dropdown, the sheet is formatted - activate the main module code - the LoopQReport
    Dim q As Integer
    
    ' Validate selection
    If cboQuarter.ListIndex = -1 Then
        MsgBox "Select a quarter (1–4) before continuing.", vbExclamation
        Exit Sub
    End If
    
    q = CInt(cboQuarter.Value)
    
    LoopQReport q
    cboQuarter.Enabled = False
    MsgBox n & "Populate Sample Data to select generate a new quarter report.", vbInformation
End Sub

Private Sub InstructionBox_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Width = 500
    Me.Height = 440


    ' Populate quarter selector
    With cboQuarter
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .ListIndex = -1 ' nothing selected by default
        .Enabled = False 'locked untl populate
        
    End With
End Sub
Private Sub UserForm_Click()

End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub
