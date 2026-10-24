
' ---------- тестовый вход: строка с кейсами ----------
Private Function Поле(ByVal s As String, ByVal n As Long, ByVal sep As String) As String
    Dim a
    a = Split(s, sep)
    If n <= UBound(a) Then Поле = a(n)
End Function

Public ВремяЯдра As String

Public Function Время() As String
    Время = ВремяЯдра
End Function

Public Function ТестСтроки(ByVal все As String) As String
    Dim кейсы, c As Long, out As String
    кейсы = Split(все, Chr(29))
    For c = 0 To UBound(кейсы)
        If Len(кейсы(c)) > 0 Then out = out & ОдинКейс(кейсы(c)) & Chr(29)
    Next c
    ТестСтроки = out
End Function

Private Function ОдинКейс(ByVal s As String) As String
    Dim строки, i As Long, k As Long, ячейки, f, p As Long, r As String, t As String
    строки = Split(s, Chr(10))
    nR = CLng(Поле(строки(0), 0, " "))
    nLbl = CLng(Поле(строки(0), 1, " "))
    ReDim inTxt(1 To nR, 1 To nLbl)
    ReDim inCT(1 To nR, 1 To nLbl)
    ReDim inLev(1 To nR, 1 To nLbl)
    ReDim inInd(1 To nR, 1 To nLbl)
    ReDim inLT(1 To nR)
    ReDim inHasVal(1 To nR)
    For i = 1 To nR
        ячейки = Split(строки(i), Chr(30))
        inLT(i) = CLng(ячейки(0))
        inHasVal(i) = (ячейки(1) = "1")
        For k = 1 To nLbl
            f = Split(ячейки(1 + k), Chr(31))
            inTxt(i, k) = f(0)
            inCT(i, k) = CLng(f(1))
            inLev(i, k) = CLng(f(2))
            inInd(i, k) = CLng(f(3))
        Next k
    Next i
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
    t0 = Timer
    ЯдроСтроки
    t1 = Timer
    ПорядокСтрок
    t2 = Timer
    ЯдроПослеПорядка
    t3 = Timer
    ВремяЯдра = Format(t1 - t0, "0.00") & " / " & Format(t2 - t1, "0.00") & " / " & Format(t3 - t2, "0.00")
    Dim части() As String
    ReDim части(1 To nOrd)
    For p = 1 To nOrd
        i = mOrd(p)
        r = i & Chr(30) & mDepth(i) & Chr(30) & mType(i) & Chr(30) & mOutLvl(p) & Chr(30)
        If mType(i) = LT_GRAND Then
            r = r & "G"
        ElseIf mРодитель(i) Then
            r = r & "S"
        End If
        For k = 1 To nLbl
            If mOutR(p, k) > 0 Then t = inTxt(mOutR(p, k), mOutC(p, k)) Else t = mOutLbl(p, k)
            r = r & Chr(30) & t & Chr(31) & mOutInd(p, k)
        Next k
        части(p) = r
    Next p
    ОдинКейс = Join(части, Chr(10))
End Function

' ---------- столбцы ----------
Public Function ТестСтолбцы(ByVal все As String) As String
    Dim кейсы, c As Long, out As String
    кейсы = Split(все, Chr(29))
    For c = 0 To UBound(кейсы)
        If Len(кейсы(c)) > 0 Then out = out & ОдинСтолбцы(кейсы(c)) & Chr(29)
    Next c
    ТестСтолбцы = out
End Function

Private Function ОдинСтолбцы(ByVal s As String) As String
    Dim строки, j As Long, k As Long, f, mx As Long, r As String, ci As Long
    строки = Split(s, Chr(10))
    nC = CLng(Поле(строки(0), 0, " "))
    mx = CLng(Поле(строки(0), 1, " "))
    cMeasLev = CLng(Поле(строки(0), 2, " "))
    nMeas = CLng(Поле(строки(0), 3, " "))
    ReDim mMeas(1 To Макс(nMeas, 1))
    For k = 1 To nMeas
        mMeas(k) = Поле(строки(0), 3 + k, " ")
    Next k
    ReDim cItem(1 To nC, 1 To mx)
    ReDim cCnt(1 To nC)
    ReDim cType(1 To nC)
    ReDim cDF(1 To nC)
    For j = 1 To nC
        f = Split(строки(j), Chr(30))
        cType(j) = CLng(f(0))
        cDF(j) = f(1)
        cCnt(j) = UBound(f) - 1
        For k = 1 To cCnt(j)
            cItem(j, k) = f(1 + k)
        Next k
    Next j
    СобратьШапку
    ПорядокСтолбцов
    For ci = 1 To nCOrd
        r = r & cOrd(ci)
        For k = 1 To nLev
            r = r & Chr(30) & dLev(cOrd(ci), k)
        Next k
        r = r & Chr(10)
    Next ci
    ОдинСтолбцы = r
End Function
