
' ---------- реальная сводная из файла: лист -> отчёт ----------
' Вход — строка с разделителями: подписи строк, значения, шапка столбцов.
Public Function ИзФайла(ByVal s As String) As String
    Dim ч, i As Long, k As Long, j As Long, r As Long, c As Long, f, nV As Long
    Dim lbl() As Variant, vals() As Variant, ws As Object, out As String, cel As Object, t As String
    On Error GoTo Ош
    ч = Split(s, Chr(29))
    ' 0: заголовок
    f = Split(ч(0), Chr(30))
    nR = CLng(f(0)): nLbl = CLng(f(1)): nV = CLng(f(2)): nHR = CLng(f(3))
    nMeas = CLng(f(4))
    ReDim mMeas(1 To nMeas)
    For k = 1 To nMeas: mMeas(k) = f(4 + k): Next k
    mСлужебные = f(5 + nMeas)
    ReDim mLblCap(1 To nLbl)
    For k = 1 To nLbl: mLblCap(k) = f(5 + nMeas + k): Next k
    ' 1: строки
    ReDim inTxt(1 To nR, 1 To nLbl): ReDim inCT(1 To nR, 1 To nLbl): ReDim inLev(1 To nR, 1 To nLbl)
    ReDim inInd(1 To nR, 1 To nLbl): ReDim inLT(1 To nR): ReDim inHasVal(1 To nR)
    ReDim lbl(1 To nR, 1 To nLbl): ReDim vals(1 To nR, 1 To nV)
    Dim стр
    стр = Split(ч(1), Chr(10))
    For i = 1 To nR
        f = Split(стр(i - 1), Chr(30))
        inLT(i) = CLng(f(0))
        For k = 1 To nLbl
            Dim g
            g = Split(f(k), Chr(31))
            inTxt(i, k) = g(0): inCT(i, k) = CLng(g(1)): inLev(i, k) = 0: inInd(i, k) = CLng(g(2))
            lbl(i, k) = inTxt(i, k)
        Next k
        For j = 1 To nV
            If Len(f(nLbl + j)) > 0 Then vals(i, j) = CDbl(f(nLbl + j)): inHasVal(i) = True
        Next j
    Next i
    ' 2: шапка
    nC = nV
    ReDim hTxt(1 To nHR, 1 To nC): ReDim hCT(1 To nHR, 1 To nC): ReDim inCLT(1 To nC)
    стр = Split(ч(2), Chr(10))
    For r = 1 To nHR
        f = Split(стр(r - 1), Chr(30))
        For j = 1 To nC
            g = Split(f(j - 1), Chr(31))
            hTxt(r, j) = g(0): hCT(r, j) = CLng(g(1))
        Next j
    Next r
    f = Split(ч(3), Chr(30))
    For j = 1 To nC: inCLT(j) = CLng(f(j - 1)): Next j

    ШапкаПоЛисту
    СобратьШапку
    ПорядокСтолбцов
    ЯдроСтроки
    ПорядокСтрок
    ЯдроПослеПорядка
    ReDim mФорматКол(1 To nC): ReDim mФорматСтр(1 To nR)
    For j = 1 To nC: mФорматКол(j) = "General": Next j
    mФильтры = ""
    InitPalette 1
    Set ws = ActiveWorkbook.Worksheets.Add
    ВывестиОтчёт ws, lbl, vals
    For r = 5 To 5 + nLev + nOrd
        For c = 1 To 1 + nLbl + nCOrd
            Set cel = ws.Cells(r, c)
            t = cel.Text
            If cel.MergeCells Then
                If cel.MergeArea.Cells(1, 1).Address <> cel.Address Then t = "<"
            End If
            If cel.IndentLevel > 0 Then t = String(cel.IndentLevel * 2, "·") & t
            If cel.Font.Bold Then t = "*" & t
            out = out & t & Chr(31) & Hex(cel.Interior.Color) & Chr(30)
        Next c
        out = out & Chr(10)
    Next r
    ИзФайла = out
    Exit Function
Ош:
    ИзФайла = "ОШИБКА " & Err.Number & ": " & Err.Description & " (" & Erl & ")"
End Function
