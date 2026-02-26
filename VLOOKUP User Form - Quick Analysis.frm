VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLookup 
   Caption         =   "Quick Analysis"
   ClientHeight    =   2508
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   8010
   OleObjectBlob   =   "VLOOKUP User Form - Quick Analysis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AVERAGE_Click()

End Sub

Private Sub btnLookup_Click()

    Dim quarters As Object
    Set quarters = CreateObject("Scripting.Dictionary")

    If chkQ1.Value Then quarters.Add "Quarter 1 Expenses", True
    If chkQ2.Value Then quarters.Add "Quarter 2 Expenses", True
    If chkQ3.Value Then quarters.Add "Quarter 3 Expenses", True
    If chkQ4.Value Then quarters.Add "Quarter 4 Expenses", True

    If quarters.count = 0 Then
        MsgBox "Select at least one quarter.", vbExclamation
        Exit Sub
    End If

    Dim resultType As String
    If optSum.Value Then resultType = "SUM"
    If optAverage.Value Then resultType = "AVG"
    If optStdDev.Value Then resultType = "STD"

    Dim result As Double

    result = LookupAggregate( _
        cboDivision.Value, _
        cboCategory.Value, _
        quarters, _
        resultType)

    MsgBox "Result: " & Format(result, "0.00"), vbInformation

End Sub


Private Sub categories_Change()

End Sub

Private Sub diviisions_Change()

End Sub

Private Sub InstructionBox_Click()

End Sub



Private Sub q1_Click()

End Sub

Private Sub Q2_Click()

End Sub

Private Sub q3_Click()

End Sub

Private Sub q4_Click()

End Sub

Private Sub stdDev_Click()

End Sub

Private Sub sum_Click()

End Sub

Private Sub UserForm_Click()

End Sub
