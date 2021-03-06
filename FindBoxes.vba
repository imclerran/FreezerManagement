' TODO: refactor (everything)

Sub Find_Boxes()
    Worksheets("New Shelf Grid").Activate
    Dim targetDate As Date
    Dim targetDate2 As Date
    targetDate = Cells(30, 20).Value
    If Not IsEmpty(Cells(31, 20)) Then
        targetDate2 = Cells(31, 20).Value
    Else
        targetDate2 = targetDate
    End If
    Call Search_Fry(targetDate, targetDate2)
    Call Search_L2(targetDate, targetDate2)
    Call Search_NSP(targetDate, targetDate2)
    Call Search_Overflow(targetDate, targetDate2)
    Call Search_Pallets(targetDate, targetDate2)
End Sub

Sub Search_Range(targetDate As Date, targetDate2, rowStart, rowEnd, colStart, colEnd)
    Dim Date1 As Date
    Dim date2 As Date
    For rowNum = rowStart To rowEnd Step 2
        For colNum = colStart To colEnd Step 1
            Date1 = Cells(rowNum, colNum)
            date2 = Cells(rowNum + 1, colNum)
            If IsEmpty(Cells(rowNum, colNum)) Then
                Call Highlight_Box(colNum, rowNum, BackgroundColor(rowNum))
            ElseIf Not targetDate = targetDate2 Then
                If Date1 = targetDate And date2 = targetDate2 Then Call Highlight_Box(colNum, rowNum, FoundColor)
            ElseIf Not IsEmpty(Cells(rowNum + 1, colNum)) Then
                If Date_In_Range(targetDate, Date1, date2) Then Call Highlight_Box(colNum, rowNum, FoundColor)
            ElseIf Date1 = targetDate Then
                Call Highlight_Box(colNum, rowNum, FoundColor)
            End If
        Next
    Next
End Sub

Sub Search_Vertical_To_End(targetDate As Date, targetDate2, rowStart As Integer, colStart As Integer)
    Dim rowNum As Integer
    Dim colNum As Integer
    Dim Date1 As Date
    Dim date2 As Date
    rowNum = rowStart
    colNum = colStart
    While (Not IsEmpty(Cells(rowNum, colNum)))
        Date1 = Cells(rowNum, colNum)
        date2 = Cells(rowNum, colNum + 1)
        If Not targetDate = targetDate2 Then ' Match exact date
            If Date1 = targetDate And date2 = targetDate2 Then
                Call Highlight_Wide_Box(colNum, rowNum, FoundColor)
            Else
                Call Highlight_Wide_Box(colNum, rowNum, Get_Background_Color(rowNum, colNum, False))
            End If
        ElseIf IsEmpty(Cells(rowNum, colNum + 1)) Then
            If Date1 = targetDate Then
                Call Highlight_Wide_Box(colNum, rowNum, FoundColor)
            Else
                Call Highlight_Wide_Box(colNum, rowNum, Get_Background_Color(rowNum, colNum, False))
            End If
        Else
            If Date_In_Range(targetDate, Date1, date2) Then
                Call Highlight_Wide_Box(colNum, rowNum, FoundColor)
            Else
                Call Highlight_Wide_Box(colNum, rowNum, Get_Background_Color(rowNum, colNum, False))
            End If
        End If
        rowNum = rowNum + 1
    Wend
End Sub

Sub Search_Fry(targetDate As Date, targetDate2 As Date)
    Call Search_Range(targetDate, targetDate2, 4, 19, 2, 17)
End Sub

Sub Search_L2(targetDate As Date, targetDate2 As Date)
    Call Search_Range(targetDate, targetDate2, 21, 22, 2, 17)
End Sub

Sub Search_NSP(targetDate As Date, targetDate2 As Date)
    Call Search_Range(targetDate, targetDate2, 24, 29, 2, 13)
End Sub

Sub Search_Overflow(targetDate As Date, targetDate2 As Date)
    Call Search_Range(targetDate, targetDate2, 31, 34, 2, 17)
End Sub

Sub Search_Pallets(targetDate As Date, targetDate2 As Date)
    Dim colNum As Integer
    For colNum = 1 To 15 Step 2
        Call Search_Vertical_To_End(targetDate, targetDate2, 38, colNum)
    Next
End Sub

Sub Reset_Pallets()
    Dim colNum As Integer
    Dim rowNum As Integer
    For colNum = 1 To 15 Step 2
        rowNum = 38
         While (Not IsEmpty(Cells(rowNum, colNum)))
         Call Highlight_Wide_Box(colNum, rowNum, Get_Background_Color(rowNum, colNum, False))
            rowNum = rowNum + 1
         Wend
    Next
End Sub

Function Date_In_Range(targetDate As Date, rangeStart As Date, rangeEnd As Date) As Boolean
    Date_In_Range = False
    If DateDiff("y", targetDate, rangeStart) <= 0 Then
        If DateDiff("y", targetDate, rangeEnd) >= 0 Then Date_In_Range = True
    End If
End Function

Sub Highlight_Box(colNum, rowNum, color As Double)
    Range(Cells(rowNum, colNum), Cells(rowNum + 1, colNum)).Interior.color = color
End Sub

Sub Highlight_Wide_Box(colNum, rowNum, color As Double)
    Range(Cells(rowNum, colNum), Cells(rowNum, colNum + 1)).Interior.color = color
End Sub

Function FoundColor() As Double
    FoundColor = RGB(0, 176, 240)
End Function

Function BackgroundColor(rowNum) As Double
    If rowNum >= 4 And rowNum <= 19 Then BackgroundColor = RGB(198, 239, 206)
    If rowNum >= 21 And rowNum <= 22 Then BackgroundColor = RGB(255, 235, 156)
    If rowNum >= 24 And rowNum <= 29 Then BackgroundColor = RGB(255, 199, 206)
    If rowNum >= 31 And rowNum <= 34 Then BackgroundColor = RGB(219, 219, 219)
End Function

Private Function Get_Background_Color(rowNum, colNum, isExpired As Boolean)
    Dim textColor As Long
    textColor = Cells(rowNum, colNum).Font.color
    If (isExpired) Then ' expired product
        Get_Background_Color = RGB(244, 176, 132)
    ElseIf (RGB(0, 97, 0) = textColor) Then ' fry line
        Get_Background_Color = RGB(198, 239, 206)
    ElseIf (RGB(156, 101, 0) = textColor) Then ' line 2
        Get_Background_Color = RGB(255, 235, 156)
    ElseIf (RGB(156, 0, 6) = textColor) Then ' nsp line
        Get_Background_Color = RGB(255, 199, 206)
    Else ' overflow or unknown
        Get_Background_Color = RGB(219, 219, 219)
    End If
End Function

