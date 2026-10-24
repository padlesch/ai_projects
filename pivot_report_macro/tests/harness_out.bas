
' ---------- прогон вывода на лист в эмуляции Excel (LibreOffice VBA) ----------
Private Sub Ячейка(ByVal i As Long, ByVal k As Long, ByVal t As String, ByVal ct As Long, ByVal lv As Long, ByVal ind As Long)
    inTxt(i, k) = t: inCT(i, k) = ct: inLev(i, k) = lv: inInd(i, k) = ind
End Sub

Private Function Дамп(ws As Object, ByVal r1 As Long, ByVal r2 As Long, ByVal c2 As Long) As String
    Dim r As Long, c As Long, s As String, cel As Object, ф As String
    For r = r1 To r2
        For c = 1 To c2
            Set cel = ws.Cells(r, c)
            ф = Hex(cel.Interior.Color)
            s = s & "[" & cel.Text & "|" & ф & "|" & РольЯчейки(cel) & "|" & cel.IndentLevel & "|" & IIf(cel.Font.Bold, "b", "") & "]"
        Next c
        s = s & Chr(10)
    Next r
    Дамп = s
End Function

Public Function ТестВывод() As String
    Dim ws As Object, lbl() As Variant, vals() As Variant, i As Long, j As Long, s As String
    On Error GoTo Ош
    ' сжатый макет, итоги внизу: Север(Москва, Тверь), Юг(Сочи), общий итог
    nR = 8: nLbl = 1
    ReDim inTxt(1 To nR, 1 To 1): ReDim inCT(1 To nR, 1 To 1): ReDim inLev(1 To nR, 1 To 1)
    ReDim inInd(1 To nR, 1 To 1): ReDim inLT(1 To nR): ReDim inHasVal(1 To nR)
    Ячейка 1, 1, "Север", 1, 1, 0
    Ячейка 2, 1, "Москва", 1, 2, 1
    Ячейка 3, 1, "Тверь", 1, 2, 1
    Ячейка 4, 1, "Север Итог", 2, 1, 0
    Ячейка 5, 1, "Юг", 1, 1, 0
    Ячейка 6, 1, "Сочи", 1, 2, 1
    Ячейка 7, 1, "Юг Итог", 2, 1, 0
    Ячейка 8, 1, "Общий итог", 3, 0, 0
    For i = 1 To nR: inLT(i) = -1: inHasVal(i) = (i <> 1 And i <> 5): Next i
    ReDim lbl(1 To nR, 1 To 1)
    ReDim vals(1 To nR, 1 To 3)
    For i = 1 To nR
        lbl(i, 1) = inTxt(i, 1)
        If inHasVal(i) Then vals(i, 1) = i * 10: vals(i, 2) = i * 0.5: vals(i, 3) = i * 10.5
    Next i
    ' столбцы: 2023, 2024, общий итог (в конце, как в Excel)
    nC = 3: nMeas = 1: ReDim mMeas(1 To 1): mMeas(1) = "Выручка": cMeasLev = 0
    ReDim cItem(1 To 3, 1 To 1): ReDim cCnt(1 To 3): ReDim cType(1 To 3): ReDim cDF(1 To 3)
    cItem(1, 1) = "2023": cItem(2, 1) = "2024": cCnt(1) = 1: cCnt(2) = 1: cCnt(3) = 1
    cType(3) = LT_GRAND
    СобратьШапку
    ПорядокСтолбцов
    ЯдроСтроки
    ПорядокСтрок
    ЯдроПослеПорядка
    ReDim mLblCap(1 To 1): mLblCap(1) = "Регион"
    ReDim mФорматКол(1 To nC): ReDim mФорматСтр(1 To nR)
    For j = 1 To nC: mФорматКол(j) = "General": Next j
    mФорматКол(2) = "0.0%"
    mФильтры = "Год: 2024"
    InitPalette 1
    Set ws = ActiveWorkbook.Worksheets(1)
    ВывестиОтчёт ws, lbl, vals
    s = "ПОСТРОЕН меток=" & ТестМетки.Count & Chr(10) & Дамп(ws, 5, 13, 5)
    s = s & "ТЕКУЩАЯ=" & ТекущаяПалитра(ws) & Chr(10)
    If ПерекраситьЛист(ws, 4) Then s = s & "ПЕРЕКРАШЕН в 4" & Chr(10) Else s = s & "НЕ ПЕРЕКРАШЕН" & Chr(10)
    s = s & Дамп(ws, 5, 13, 5)
    s = s & "ТЕКУЩАЯ=" & ТекущаяПалитра(ws) & Chr(10)
    s = s & "формат C8=" & ws.Cells(8, 4).NumberFormat & " D8=" & ws.Cells(8, 5).NumberFormat
    ТестВывод = s
    Exit Function
Ош:
    ТестВывод = s & Chr(10) & "ОШИБКА " & Err.Number & ": " & Err.Description & " (строка " & Erl & ")"
End Function
