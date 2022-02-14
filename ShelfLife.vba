Enum LineType
    [_First]
    Fry
    L2
    NSP
    None
    [_Last]
End Enum

Public Sub Update_Expired()
    Worksheets("New Shelf Grid").Activate
    Update_Fry
    Update_L2
    Update_NSP
    Update_Overflow
End Sub

Private Sub Update_Fry()
    Const rowStart = 4
    Const rowEnd = 19
    Const colStart = 2
    Const colEnd = 17
    Const shelfRow = 7
    Const shelfCol = 20
    Call Update_Cells(rowStart, rowEnd, colStart, colEnd)
End Sub

Private Sub Update_L2()
    Const rowStart = 21
    Const rowEnd = 22
    Const colStart = 2
    Const colEnd = 17
    Const shelfRow = 8
    Const shelfCol = 20
    Call Update_Cells(rowStart, rowEnd, colStart, colEnd)
End Sub

Private Sub Update_NSP()
    Const rowStart = 24
    Const rowEnd = 29
    Const colStart = 2
    Const colEnd = 13
    Const shelfRow = 9
    Const shelfCol = 20
    Call Update_Cells(rowStart, rowEnd, colStart, colEnd)
End Sub

Private Sub Update_Overflow()
    Const rowStart = 31
    Const rowEnd = 34
    Const colStart = 2
    Const colEnd = 17
    Const shelfRow = 10
    Const shelfCol = 20
    Call Update_Cells(rowStart, rowEnd, colStart, colEnd)
End Sub

Private Sub Update_Cells(rowStart As Integer, rowEnd As Integer, colStart As Integer, colEnd As Integer)
    Dim val As Date
    Dim today As Date
    today = Date
    For rowNum = rowStart To rowEnd Step 2
        For colNum = colStart To colEnd Step 1
            Dim textColor As Long
            textColor = Cells(rowNum, colNum).Font.color
            If (Check_Cell_Expired(rowNum, colNum)) Then
                Range(Cells(rowNum, colNum), Cells(rowNum + 1, colNum)).Interior.color = Get_Background_Color(textColor, True)
            Else
                Range(Cells(rowNum, colNum), Cells(rowNum + 1, colNum)).Interior.color = Get_Background_Color(textColor, False)
            End If
        Next
    Next
End Sub

Private Function Check_Cell_Expired(ByVal rowNum As Integer, ByVal colNum As Integer) As Boolean
    Dim cellVal As Date
    Dim today As Date
    Dim val As Variant
    Dim shelfLife As Integer
    today = Date
    If Not IsEmpty(Cells(rowNum, colNum)) And VarType(Cells(rowNum, colNum).Value) <> VbVarType.vbDate Then
        val = Null
    ElseIf (Not IsEmpty(Cells(rowNum + 1, colNum))) Then
        val = Cells(rowNum + 1, colNum).Value
    Else
        val = Cells(rowNum, colNum).Value
    End If
    
    shelfLife = GetCellShelfLife(rowNum, colNum)
    If IsNull(val) Then
        Check_Cell_Expired = True
    ElseIf (DateDiff("y", val, today) >= shelfLife And Not IsEmpty(Cells(rowNum, colNum))) Then ' is expired
        Check_Cell_Expired = True
    Else ' not expired
        Check_Cell_Expired = False
    End If
End Function

Private Function Get_Background_Color(textColor As Long, isExpired As Boolean)
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

Private Function ShelfLifeFry()
    ShelfLifeFry = Cells(9, 20)
End Function

Private Function ShelfLifeL2()
    ShelfLifeL2 = Cells(9, 20)
End Function

Private Function ShelfLifeNSP()
    ShelfLifeNSP = Cells(9, 20)
End Function

Private Function ShelfLifeOverflow()
    ShelfLifeOverflow = Cells(10, 20)
End Function

Function GetCellLineType(rowNum As Integer, colNum As Integer) As Long
    Dim textColor As Long
    textColor = Cells(rowNum, colNum).Font.color
    If (RGB(0, 97, 0) = textColor) Then ' fry line
        GetCellLineType = LineType.Fry
    ElseIf (RGB(156, 101, 0) = textColor) Then ' line 2
        GetCellLineType = LineType.L2
    ElseIf (RGB(156, 0, 6) = textColor) Then ' nsp line
        GetCellLineType = LineType.NSP
    Else ' no line type
        GetCellLineType = LineType.None
    End If
End Function

Function GetCellShelfLife(rowNum As Integer, colNum As Integer)
    Dim cellType As Integer
    cellType = GetCellLineType(rowNum, colNum)
    If (cellType = LineType.Fry) Then
        GetCellShelfLife = ShelfLifeFry
    ElseIf (cellType = LineType.L2) Then
        GetCellShelfLife = ShelfLifeL2
    ElseIf (cellType = LineType.NSP) Then
        GetCellShelfLife = ShelfLifeNSP
    Else
        GetCellShelfLife = 0
    End If
End Function

