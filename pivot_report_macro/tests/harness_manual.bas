
' ---------- ручная проверка: сводная как на листе -> готовый отчёт ----------
Public Function Построить(ByVal s As String, ByVal номерЛиста As Long) As String
    Dim строки, i As Long, k As Long, j As Long, ячейки, f
    Dim lbl() As Variant, vals() As Variant, ws As Object, nV As Long, r As Long, c As Long
    Dim out As String, cel As Object, t As String
    On Error GoTo Ош
    строки = Split(s, Chr(10))
    nR = CLng(Split(строки(0), " ")(0))
    nLbl = CLng(Split(строки(0), " ")(1))
    nV = CLng(Split(строки(0), " ")(2))
    ReDim inTxt(1 To nR, 1 To nLbl): ReDim inCT(1 To nR, 1 To nLbl): ReDim inLev(1 To nR, 1 To nLbl)
    ReDim inInd(1 To nR, 1 To nLbl): ReDim inLT(1 To nR): ReDim inHasVal(1 To nR)
    ReDim lbl(1 To nR, 1 To nLbl): ReDim vals(1 To nR, 1 To nV)
    For i = 1 To nR
        ячейки = Split(строки(i), Chr(30))
        inLT(i) = CLng(ячейки(0))
        For k = 1 To nLbl
            f = Split(ячейки(k), Chr(31))
            inTxt(i, k) = f(0): inCT(i, k) = CLng(f(1)): inLev(i, k) = CLng(f(2)): inInd(i, k) = CLng(f(3))
            lbl(i, k) = inTxt(i, k)
        Next k
        For j = 1 To nV
            t = ячейки(nLbl + j)
            If Len(t) > 0 Then vals(i, j) = CDbl(t): inHasVal(i) = True
        Next j
    Next i
    ' шапка: последние столбец — общий итог
    nC = nV: nMeas = 1: ReDim mMeas(1 To 1): mMeas(1) = "Сумма": cMeasLev = 0
    ReDim cItem(1 To nC, 1 To 1): ReDim cCnt(1 To nC): ReDim cType(1 To nC): ReDim cDF(1 To nC)
    f = Split(строки(nR + 1), Chr(30))
    For j = 1 To nC
        cItem(j, 1) = f(j - 1): cCnt(j) = 1
        If f(j - 1) = "Общий итог" Then cType(j) = LT_GRAND
    Next j
    ReDim mLblCap(1 To nLbl)
    f = Split(строки(nR + 2), Chr(30))
    For k = 1 To nLbl: mLblCap(k) = f(k - 1): Next k
    СобратьШапку
    ПорядокСтолбцов
    ЯдроСтроки
    ПорядокСтрок
    ЯдроПослеПорядка
    ReDim mФорматКол(1 To nC): ReDim mФорматСтр(1 To nR)
    For j = 1 To nC: mФорматКол(j) = "General": Next j
    mФильтры = ""
    InitPalette 1
    ' каждый отчёт — на новом чистом листе, как в макросе
    Set ws = ActiveWorkbook.Worksheets.Add
    ВывестиОтчёт ws, lbl, vals
    ' текстом: строки отчёта
    For r = 5 To 5 + nLev + nOrd
        For c = 1 To 1 + nLbl + nCOrd
            Set cel = ws.Cells(r, c)
            t = cel.Text
            If cel.IndentLevel > 0 Then t = String(cel.IndentLevel * 2, "·") & t
            If cel.Font.Bold Then t = "*" & t
            out = out & t & Chr(31) & Hex(cel.Interior.Color) & Chr(30)
        Next c
        out = out & Chr(10)
    Next r
    Построить = out
    Exit Function
Ош:
    Построить = "ОШИБКА " & Err.Number & ": " & Err.Description
End Function
