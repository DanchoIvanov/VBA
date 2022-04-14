VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Calendar"
   ClientHeight    =   4665
   ClientLeft      =   180
   ClientTop       =   705
   ClientWidth     =   5250
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim YEARS_BEFORE As Integer
Dim YEARS_AFTER As Integer

Private Sub setConstants()
     
    YEARS_BEFORE = 20
    YEARS_AFTER = 80
    
End Sub

Private Sub UserForm_Activate()
   
    setConstants
    
    Dim i As Integer
    
    With Me.MonthComboBox
        For i = 1 To 12
            .AddItem Format(DateSerial(Year(Date), i, 1), "MMMM")
        Next i
        '.Value = Format(Date, "MMMM")
    End With
    
    
    With Me.YearComboBox
        For i = Year(Date) - YEARS_BEFORE To Year(Date) + YEARS_AFTER
            .AddItem i
        Next i
        '.Value = Format(Date, "YYYY")
    End With
    
    'showDate
    
    If Me.DateTextBox.Value <> "" Then
        Me.MonthComboBox.ListIndex = CInt(Month(CDate(Me.DateTextBox.Value))) - 1
        Me.YearComboBox.Value = Year(CDate(Me.DateTextBox.Value))
    Else
        Me.MonthComboBox.ListIndex = CInt(Month(Date)) - 1
        Me.YearComboBox.Value = Year(Date)
    End If
    
End Sub

Private Sub showDate()
    
    Dim firstDate As Date
    Dim lastDate As Date
    
    firstDate = CDate("1 " & Me.MonthComboBox.Value & " " & Me.YearComboBox.Value)
    lastDate = DateSerial(Year(firstDate), Month(firstDate) + 1, 1) - 1
    
    removeCaption
    setFirstDateOfmonth (firstDate)
    setAllDates (Day(lastDate))
    disableUnusedButtons
    highlightDate
    
End Sub

Private Sub removeCaption()
    
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    
    For i = 1 To 42
        Set btn = Me.Controls("CommandButton" & i)
        btn.Caption = ""
        btn.BackColor = RGB(240, 240, 240) 'RGB(216, 208, 200)
    Next i
    
End Sub

Private Sub setFirstDateOfmonth(firstDate As Date)
    
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    
    For i = 1 To 7
        Set btn = Me.Controls("CommandButton" & i)
        If VBA.Weekday(firstDate, vbMonday) = i Then
        btn.Caption = "1"
        GoTo breakCycle
        End If
    Next i

breakCycle:
End Sub

Private Sub setAllDates(lastDay As Integer)

    Dim i As Integer
    Dim btn1 As MSForms.CommandButton
    Dim btn2 As MSForms.CommandButton
    
    For i = 1 To 41
        Set btn1 = Me.Controls("CommandButton" & i)
        Set btn2 = Me.Controls("CommandButton" & i + 1)
        
        If btn1.Caption <> "" Then
            If CInt(btn1.Caption) < lastDay Then
                btn2.Caption = btn1.Caption + 1
            End If
        End If
    Next i

End Sub

Private Sub highlightDate()

    Dim i As Integer
    Dim btn As MSForms.CommandButton
    Dim dateToHighlight As Date
    
    If Me.DateTextBox.Value <> "" Then
        dateToHighlight = CDate(Me.DateTextBox.Value)
    Else
        dateToHighlight = Date
    End If
    
    If (Me.MonthComboBox.ListIndex = CInt(Month(dateToHighlight)) - 1) And (CInt(Me.YearComboBox.Value) = Year(dateToHighlight)) Then
        For i = 1 To 42
            Set btn = Me.Controls("CommandButton" & i)
            If CStr(Day(dateToHighlight)) = btn.Caption Then
                btn.BackColor = RGB(192, 192, 192)
            End If
        Next i
    End If

End Sub

Private Sub MonthComboBox_Change()

    If Me.MonthComboBox.Value <> "" And Me.YearComboBox.Value <> "" Then
        If Me.MonthComboBox.ListIndex >= 0 And isNumeric(Me.YearComboBox.Value) Then
            If Me.YearComboBox.ListIndex >= 0 Then
                showDate
                Me.MonthYearLabel.Caption = Me.MonthComboBox.Value & " - " & Me.YearComboBox.Value
            Else
                GoTo fixDate
            End If
        Else
fixDate:
            Me.MonthComboBox.ListIndex = CInt(Month(Date)) - 1
            Me.YearComboBox.Value = Year(Date)
            showDate
            Me.MonthYearLabel.Caption = Me.MonthComboBox.Value & " - " & Me.YearComboBox.Value
        End If
    End If
    
    Exit Sub

End Sub

Private Sub YearComboBox_Change()

    If Me.MonthComboBox.Value <> "" And Me.YearComboBox.Value <> "" Then
        If Me.MonthComboBox.ListIndex >= 0 And isNumeric(Me.YearComboBox.Value) Then
            If Me.YearComboBox.ListIndex >= 0 Then
                showDate
                Me.MonthYearLabel.Caption = Me.MonthComboBox.Value & " - " & Me.YearComboBox.Value
            Else
                GoTo fixDate
            End If
        Else
fixDate:
            Me.MonthComboBox.ListIndex = CInt(Month(Date)) - 1
            Me.YearComboBox.Value = Year(Date)
            showDate
            Me.MonthYearLabel.Caption = Me.MonthComboBox.Value & " - " & Me.YearComboBox.Value
        End If
    End If
    
    disableButton
    
End Sub

Private Sub PreviousMonthCommandButton_Click()
    
    If Me.MonthComboBox.ListIndex = 0 Then
        Me.MonthComboBox.ListIndex = 11
        Me.YearComboBox.Value = Me.YearComboBox - 1
    Else
        Me.MonthComboBox.ListIndex = Me.MonthComboBox.ListIndex - 1
    End If
    
    disableButton
    
End Sub

Private Sub NextMonthCommandButton_Click()

    If Me.MonthComboBox.ListIndex = 11 Then
        Me.MonthComboBox.ListIndex = 0
        Me.YearComboBox.Value = Me.YearComboBox + 1
    Else
        Me.MonthComboBox.ListIndex = Me.MonthComboBox.ListIndex + 1
    End If
    
    disableButton

End Sub

Private Sub disableButton()
    
    If (Me.MonthComboBox.ListIndex = 0 And CInt(Me.YearComboBox.Value) = Year(Date) - YEARS_BEFORE) Or (CInt(Me.YearComboBox.Value) < Year(Date) - YEARS_BEFORE) Then
        Me.PreviousMonthCommandButton.Enabled = False
    Else
        Me.PreviousMonthCommandButton.Enabled = True
    End If
    
    If (Me.MonthComboBox.ListIndex = 11 And CInt(Me.YearComboBox.Value) = Year(Date) + YEARS_AFTER) Or (CInt(Me.YearComboBox.Value) > Year(Date) + YEARS_AFTER) Then
        Me.NextMonthCommandButton.Enabled = False
    Else
        Me.NextMonthCommandButton.Enabled = True
    End If

End Sub

Private Sub disableUnusedButtons()
    
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    
    For i = 1 To 42
        Set btn = Me.Controls("CommandButton" & i)
        If btn.Caption = "" Then
            btn.Visible = False
        Else
            btn.Visible = True
        End If
    Next i
    
End Sub

Function datePicker(Optional dateInput As Object) As String
    
    Dim str As String
    
    'If TypeName(dateInput) = "Textbox" Or TypeName(dateInput) = "Range" Then
        'str = dateInput.Value
    If TypeName(dateInput) = "CommandBtton" Or TypeName(dateInput) = "Label" Then
        str = dateInput.Caption
    Else
        str = dateInput.Value
    End If
    
    If isDate(str) = True Then
        Me.DateTextBox.Value = Format(CDate(str), "MM/DD/YYYY")
    Else
        Me.DateTextBox.Value = ""
    End If
    
    Calendar.Show
    
    If TypeName(dateInput) = "Textbox" Or TypeName(dateInput) = "Range" Then
        dateInput.Value = Me.DateTextBox.Value
    ElseIf TypeName(dateInput) = "CommandBtton" Or TypeName(dateInput) = "Label" Then
        dateInput.Caption = Me.DateTextBox.Value
    Else
        datePicker = Me.DateTextBox.Value
    End If
    
End Function

Private Sub buttonClick(btn As MSForms.CommandButton)

    If btn.Caption <> "" Then
        Me.DateTextBox = Format(CDate(btn.Caption & "-" & Left(Me.MonthComboBox.Value, 3) & "-" & Me.YearComboBox), "MM/DD/YYYY")
    End If
    Unload Me
    
End Sub

Private Sub CommandButton1_Click()
    buttonClick Me.CommandButton1
End Sub
Private Sub CommandButton2_Click()
    buttonClick Me.CommandButton2
End Sub
Private Sub CommandButton3_Click()
    buttonClick Me.CommandButton3
End Sub
Private Sub CommandButton4_Click()
    buttonClick Me.CommandButton4
End Sub
Private Sub CommandButton5_Click()
    buttonClick Me.CommandButton5
End Sub
Private Sub CommandButton6_Click()
    buttonClick Me.CommandButton6
End Sub
Private Sub CommandButton7_Click()
    buttonClick Me.CommandButton7
End Sub
Private Sub CommandButton8_Click()
    buttonClick Me.CommandButton8
End Sub
Private Sub CommandButton9_Click()
    buttonClick Me.CommandButton9
End Sub
Private Sub CommandButton10_Click()
    buttonClick Me.CommandButton10
End Sub
Private Sub CommandButton11_Click()
    buttonClick Me.CommandButton11
End Sub
Private Sub CommandButton12_Click()
    buttonClick Me.CommandButton12
End Sub
Private Sub CommandButton13_Click()
    buttonClick Me.CommandButton13
End Sub
Private Sub CommandButton14_Click()
    buttonClick Me.CommandButton14
End Sub
Private Sub CommandButton15_Click()
    buttonClick Me.CommandButton15
End Sub
Private Sub CommandButton16_Click()
    buttonClick Me.CommandButton16
End Sub
Private Sub CommandButton17_Click()
    buttonClick Me.CommandButton17
End Sub
Private Sub CommandButton18_Click()
    buttonClick Me.CommandButton18
End Sub
Private Sub CommandButton19_Click()
    buttonClick Me.CommandButton19
End Sub
Private Sub CommandButton20_Click()
    buttonClick Me.CommandButton20
End Sub
Private Sub CommandButton21_Click()
    buttonClick Me.CommandButton21
End Sub
Private Sub CommandButton22_Click()
    buttonClick Me.CommandButton22
End Sub
Private Sub CommandButton23_Click()
    buttonClick Me.CommandButton23
End Sub
Private Sub CommandButton24_Click()
    buttonClick Me.CommandButton24
End Sub
Private Sub CommandButton25_Click()
    buttonClick Me.CommandButton25
End Sub
Private Sub CommandButton26_Click()
    buttonClick Me.CommandButton26
End Sub
Private Sub CommandButton27_Click()
    buttonClick Me.CommandButton27
End Sub
Private Sub CommandButton28_Click()
    buttonClick Me.CommandButton28
End Sub
Private Sub CommandButton29_Click()
    buttonClick Me.CommandButton29
End Sub
Private Sub CommandButton30_Click()
    buttonClick Me.CommandButton30
End Sub
Private Sub CommandButton31_Click()
    buttonClick Me.CommandButton31
End Sub
Private Sub CommandButton32_Click()
    buttonClick Me.CommandButton32
End Sub
Private Sub CommandButton33_Click()
    buttonClick Me.CommandButton33
End Sub
Private Sub CommandButton34_Click()
    buttonClick Me.CommandButton34
End Sub
Private Sub CommandButton35_Click()
    buttonClick Me.CommandButton35
End Sub
Private Sub CommandButton36_Click()
    buttonClick Me.CommandButton36
End Sub
Private Sub CommandButton37_Click()
    buttonClick Me.CommandButton37
End Sub
Private Sub CommandButton38_Click()
    buttonClick Me.CommandButton38
End Sub
Private Sub CommandButton39_Click()
    buttonClick Me.CommandButton39
End Sub
Private Sub CommandButton40_Click()
    buttonClick Me.CommandButton40
End Sub
Private Sub CommandButton41_Click()
    buttonClick Me.CommandButton41
End Sub
Private Sub CommandButton42_Click()
    buttonClick Me.CommandButton42
End Sub
