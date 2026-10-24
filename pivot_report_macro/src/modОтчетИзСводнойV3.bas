Attribute VB_Name = "modОтчетИзСводнойV3"
Option Explicit

'==========================================================================
'  ОФОРМЛЕНИЕ ОТЧЁТА ИЗ СВОДНОЙ ТАБЛИЦЫ  — версия 3
'--------------------------------------------------------------------------
'  ОформитьОтчёт        — встать в любую ячейку сводной и запустить
'  ОформитьОтчётБезВопросов — то же молча, отчёт новым листом этой книги
'  ПалитраОтчёта         — сменить палитру готового отчёта (активный лист)
'  ПалитраСледующая      — переключить палитру по кругу
'
'  Что делает:
'   * снимает со сводной значения и структуру — в любом макете:
'     сжатом, структурном, табличном и смешанном, с итогами вверху,
'     внизу или без них, до 32 уровней строк;
'   * общий итог по строкам -> первой строкой данных;
'   * итог уровня всегда над своими детьми (итоги «внизу» поднимаются,
'     пустая строка-заголовок группы сливается со своим итогом);
'   * общий итог и промежуточные итоги по столбцам -> влево от своей группы;
'   * несколько мер -> каждая мера отдельным блоком столбцов;
'   * числовые форматы мер (%, валюта, дроби) берутся из сводной;
'   * строка фильтров сводной под заголовком;
'   * группировка строк по уровням (кнопки 1-2-3... слева) — свернуть до нужной глубины;
'   * заголовок, нумерация столбцов и строк, Arial Narrow, серые границы.
'
'  На листе после построения нет ничего, кроме отчёта: ни панелей,
'  ни служебных ячеек. Палитра меняется отдельным макросом.
'
'  Работает и с обычными сводными, и со сводными на модели данных (OLAP).
'==========================================================================

'#ДЕКЛ-НАЧАЛО
'--- переключатели поведения ---------------------------------------------
Private Const FNT           As String = "Arial Narrow"
Private Const FS_BODY       As Long = 9         ' кегль таблицы
Private Const FS_TITLE      As Long = 16        ' кегль заголовка
Private Const FS_NUM        As Long = 7         ' кегль нумерации
Private Const TXT_TITLE     As String = "Анализ данных"
Private Const TXT_SUB       As String = "ДЗМ, ЗЛ"
Private Const TXT_GRAND     As String = "Общий итог"
Private Const TXT_SUBTOT    As String = "Итог"
Private Const TXT_EMPTY     As String = "(пусто)"
Private Const MARGIN_CM     As Double = 0.3
Private Const ПАЛ_ПО_УМОЛЧ  As Long = 1         ' палитра нового отчёта
Private Const ГРУППИРОВАТЬ  As Boolean = True   ' кнопки свёртки уровней слева

'--- типы строк/столбцов сводной (как XlPivotLineType) --------------------
Private Const LT_REGULAR    As Long = 0
Private Const LT_SUBTOTAL   As Long = 1
Private Const LT_GRAND      As Long = 2
Private Const LT_BLANK      As Long = 3

'--- типы ячеек сводной (как XlPivotCellType) ------------------------------
Private Const CT_ITEM       As Long = 1
Private Const CT_SUBTOTAL   As Long = 2
Private Const CT_GRAND      As Long = 3
Private Const CT_DATAFIELD  As Long = 4
Private Const CT_CUSTOMSUB  As Long = 7
Private Const CT_BLANK      As Long = 9

'--- палитра --------------------------------------------------------------
Private cTitle As Long, cBordo As Long, cHead As Long, cNum As Long
Private cBrd As Long, cBrdOut As Long
Private mПал As Long

'--- вход ядра строк: то, что снято с листа ---------------------------------
Private nR As Long, nLbl As Long
Private inTxt() As String         ' (строка, столбец подписей) текст
Private inCT() As Long            ' тип ячейки сводной, -1 — неизвестен
Private inLev() As Long           ' номер поля строк (1..), 0 — неизвестен
Private inInd() As Long           ' отступ в ячейке
Private inLT() As Long            ' тип линии из оси сводной, -1 — неизвестен
Private inHasVal() As Boolean     ' в строке есть хоть одно значение

'--- результат ядра строк --------------------------------------------------
Private mType() As Long
Private mDepth() As Long          ' глубина строки (1 = верхний уровень)
Private mKey() As String          ' (строка, уровень) — подпись предка
Private mKeyR() As Long           ' откуда взята подпись: строка источника
Private mKeyC() As Long           '                       столбец источника
Private mPath() As String         ' (строка, уровень) — полный путь, для сравнения
Private mOwnCustom() As Boolean   ' пользовательский итог — своя подпись как есть
Private mDrop() As Boolean        ' строка поглощена своим итогом
Private mMaxD As Long
Private mColOf() As Long          ' в каком столбце подписей живёт уровень
Private mIndOf() As Long          ' отступ уровня внутри столбца
Private mCellLev() As Long
Private mOrd() As Long
Private nOrd As Long
Private mРодитель() As Boolean    ' у строки есть подчинённые строки
Private mЕстьРодитель() As Boolean ' (строка, уровень) — есть отдельная строка-предок
Private mOutLbl() As String       ' (позиция, столбец) — подпись в отчёте
Private mOutR() As Long           ' (позиция, столбец) — исходная ячейка подписи
Private mOutC() As Long
Private mOutInd() As Long         ' (позиция, столбец) — отступ, -1 не задан
Private mOutLvl() As Long         ' (позиция) — уровень группировки строк

'--- столбцы --------------------------------------------------------------
Private nC As Long, nLev As Long
Private cItem() As String         ' (столбец, уровень) элементы оси столбцов
Private cCnt() As Long            ' сколько уровней задано у столбца
Private cDF() As String           ' имя меры столбца, если известно
Private cMeasLev As Long          ' на каком уровне оси стоят меры (0 — нет)
Private mMeas() As String         ' имена мер по порядку
Private nMeas As Long
Private dLev() As String          ' dLev(столбец, уровень) — текст шапки
Private dDepth() As Long          ' сколько уровней у столбца задано
Private cType() As Long
Private cOrd() As Long
Private nCOrd As Long
'#ДЕКЛ-КОНЕЦ

Private mLblCap() As String       ' подписи столбцов с уровнями строк
Private mTot() As Double
Private mКлассСтр() As String     ' класс каждой строки отчёта
Private mКлассКол() As String     ' класс каждой колонки отчёта
Private mИсходник As Range        ' область значений исходной сводной
Private mФорматКол() As String    ' числовой формат по колонкам источника
Private mФорматСтр() As String    ' ... или по строкам (меры в строках)
Private mФильтры As String
Private mQuiet As Boolean         ' True — работать молча (вызов из другого макроса)


'==========================================================================
'  ТОЧКА ВХОДА
'==========================================================================
' Молчаливый режим: без вопросов, отчёт кладётся новым листом в ту же книгу.
Public Sub ОформитьОтчётБезВопросов()
    mQuiet = True
    ОформитьОтчёт
    mQuiet = False
End Sub

Public Sub ОформитьОтчёт()

    Dim pt As PivotTable
    Dim wsOut As Worksheet
    Dim lbl() As Variant, vals() As Variant
    Dim calc As Long

    On Error GoTo Сбой

    Set pt = ПолучитьСводную()
    If pt Is Nothing Then Exit Sub

    InitPalette ПАЛ_ПО_УМОЛЧ

    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If Not РазобратьСводную(pt, lbl, vals) Then GoTo Выход

    ПорядокСтолбцов
    ЯдроСтроки
    ПорядокСтрок
    ЯдроПослеПорядка

    Application.ScreenUpdating = True
    Set wsOut = СоздатьЛист(pt)
    If wsOut Is Nothing Then GoTo Выход
    Application.ScreenUpdating = False

    ВывестиОтчёт wsOut, lbl, vals

    Application.Calculation = calc
    Application.ScreenUpdating = True
    Set mИсходник = Nothing

    wsOut.Activate
    wsOut.Range("A1").Select
    If Not mQuiet Then
        MsgBox "Отчёт сформирован на листе """ & wsOut.Name & """." & vbLf & vbLf & _
               "Порядок строк взят из сводной; итоги подняты над своими группами." & vbLf & _
               "Сменить цвет: макрос ПалитраОтчёта.", _
               vbInformation, "Готово"
    End If
    Exit Sub

Выход:
    Application.Calculation = calc
    Application.ScreenUpdating = True
    Set mИсходник = Nothing
    Exit Sub

Сбой:
    Set mИсходник = Nothing
    ' Без этого Excel остался бы с выключенной отрисовкой и пересчётом
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.PrintCommunication = True
    If calc <> 0 Then Application.Calculation = calc
    Сообщить "Не удалось построить отчёт." & vbLf & _
             "Ошибка " & Err.Number & ": " & Err.Description

End Sub


'==========================================================================
'  ПАЛИТРЫ
'  Каждая задаётся двумя цветами: заголовок и шапка. Светлые тона итогов
'  считаются от цвета шапки, поэтому набор нельзя рассогласовать.
'  Чтобы добавить палитру: увеличить ЧислоПалитр и дописать Case.
'==========================================================================
Private Function ЧислоПалитр() As Long
    ЧислоПалитр = 4
End Function

Private Function ИмяПалитры(ByVal n As Long) As String
    Select Case n
        Case 2: ИмяПалитры = "Серый"
        Case 3: ИмяПалитры = "Дизайн 1"
        Case 4: ИмяПалитры = "Дизайн 2"
        Case Else: ИмяПалитры = "Тёмно-синий"
    End Select
End Function

Private Sub ЦветаПалитры(ByVal n As Long, ByRef заг As Long, ByRef шап As Long)
    Select Case n
        Case 2:  заг = RGB(74, 79, 85):   шап = RGB(112, 119, 126)
        ' TODO: цвета из дизайна — заменить на свои (заголовок, шапка)
        Case 3:  заг = RGB(31, 56, 100):  шап = RGB(47, 78, 111)
        Case 4:  заг = RGB(74, 79, 85):   шап = RGB(112, 119, 126)
        Case Else: заг = RGB(31, 56, 100): шап = RGB(47, 78, 111)
    End Select
End Sub

Private Sub InitPalette(ByVal n As Long)
    If n < 1 Or n > ЧислоПалитр() Then n = 1
    mПал = n
    ЦветаПалитры n, cTitle, cHead
    cBordo = RGB(139, 46, 60)        ' #8B2E3C  подзаголовок / фильтры
    cNum = RGB(166, 166, 166)        ' #A6A6A6  нумерация
    cBrd = RGB(191, 191, 191)        ' #BFBFBF  внутренние границы
    cBrdOut = RGB(140, 140, 140)     ' #8C8C8C  внешняя рамка
End Sub

' Осветление к белому: 0 — исходный цвет, 1 — чистый белый
Private Function Оттенок(ByVal c As Long, ByVal f As Double) As Long
    Dim r As Long, g As Long, b As Long
    If f < 0 Then f = 0
    If f > 1 Then f = 1
    r = c Mod 256
    g = (c \ 256) Mod 256
    b = (c \ 65536) Mod 256
    Оттенок = RGB(r + (255 - r) * f, g + (255 - g) * f, b + (255 - b) * f)
End Function

' Насколько светлой должна быть строка данного класса.
' "G" — общий итог, "S<n>" — родитель/промежуточный итог уровня n, "" — детализация
Private Function ФонСтроки(ByVal класс As String) As Double
    Dim d As Long
    ФонСтроки = -1
    If класс = "G" Then ФонСтроки = 0.74: Exit Function
    If Left$(класс, 1) = "S" Then
        d = 1
        If Len(класс) > 1 Then d = CLng(Mid$(класс, 2))
        ФонСтроки = 0.85 + 0.035 * (d - 1)
        If ФонСтроки > 0.95 Then ФонСтроки = 0.95
    End If
End Function

' Колонка главного итога заметно темнее промежуточных
Private Function ФонКолонки(ByVal класс As String, ByVal фонСтр As Double) As Double
    ФонКолонки = -1
    If класс <> "G" And класс <> "S" Then Exit Function
    If фонСтр < 0 Then
        If класс = "G" Then ФонКолонки = 0.86 Else ФонКолонки = 0.93
    Else
        If класс = "G" Then ФонКолонки = фонСтр - 0.1 Else ФонКолонки = фонСтр - 0.05
    End If
    If ФонКолонки < 0 Then ФонКолонки = 0
End Function


'==========================================================================
'  СМЕНА ПАЛИТРЫ ГОТОВОГО ОТЧЁТА — отдельные макросы
'  Служебных ячеек нет: цвета узнаются по самим заливкам. Каждая заливка
'  отчёта — это цвет шапки одной из палитр, осветлённый на известную долю,
'  поэтому её можно однозначно перевести в ту же долю новой палитры.
'==========================================================================
Public Sub ПалитраОтчёта()
    Dim s As String, i As Long, otvet As String, тек As Long
    тек = ТекущаяПалитра(ActiveSheet)
    For i = 1 To ЧислоПалитр()
        s = s & i & " — " & ИмяПалитры(i)
        If i = тек Then s = s & "   (сейчас)"
        s = s & vbLf
    Next i
    otvet = InputBox("Номер палитры для отчёта на активном листе:" & vbLf & vbLf & s, _
                     "Палитра отчёта", CStr(тек Mod ЧислоПалитр() + 1))
    If Len(otvet) = 0 Or Not IsNumeric(otvet) Then Exit Sub
    ПалитраНомер CLng(otvet)
End Sub

Public Sub ПалитраСледующая()
    Dim тек As Long
    тек = ТекущаяПалитра(ActiveSheet)
    If тек = 0 Then
        MsgBox "На активном листе не найден отчёт, построенный макросом.", vbExclamation
        Exit Sub
    End If
    ПалитраНомер тек Mod ЧислоПалитр() + 1
End Sub

' Можно повесить на кнопку панели быстрого доступа: ПалитраНомер 2
Public Sub ПалитраНомер(ByVal n As Long)
    If n < 1 Or n > ЧислоПалитр() Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Not ПерекраситьЛист(ActiveSheet, n) Then
        MsgBox "На активном листе не найден отчёт, построенный макросом.", vbExclamation
    End If
End Sub

' Доли осветления, которые вообще встречаются в отчёте (0 — сама шапка)
Private Function ДолиЗаливки() As Variant
    Dim a() As Double, n As Long, d As Long, f As Double, g As Double
    Dim классы As Variant, кк As Variant, c As Variant
    ReDim a(1 To 200)
    n = 1: a(1) = 0
    классы = Array("", "G")
    For Each c In классы
        f = ФонСтроки(CStr(c))
        ДобавитьДолю a, n, f
        For Each кк In Array("G", "S")
            ДобавитьДолю a, n, ФонКолонки(CStr(кк), f)
        Next кк
    Next c
    For d = 1 To 40
        f = ФонСтроки("S" & d)
        ДобавитьДолю a, n, f
        For Each кк In Array("G", "S")
            g = ФонКолонки(CStr(кк), f)
            ДобавитьДолю a, n, g
        Next кк
    Next d
    ReDim Preserve a(1 To n)
    ДолиЗаливки = a
End Function

Private Sub ДобавитьДолю(ByRef a() As Double, ByRef n As Long, ByVal f As Double)
    Dim i As Long
    If f < 0 Then Exit Sub
    For i = 1 To n
        If Abs(a(i) - f) < 0.000001 Then Exit Sub
    Next i
    n = n + 1
    a(n) = f
End Sub

' Номер палитры, в которой сейчас окрашен отчёт; 0 — отчёта нет
Private Function ТекущаяПалитра(ws As Object) As Long
    Dim cel As Range, p As Long, заг As Long, шап As Long
    On Error Resume Next
    If TypeName(ws) <> "Worksheet" Then Exit Function
    For Each cel In ws.Range(ws.Cells(1, 1), ws.Cells(40, 6)).Cells
        If cel.Interior.Pattern <> xlNone Then
            For p = 1 To ЧислоПалитр()
                ЦветаПалитры p, заг, шап
                If cel.Interior.Color = шап Then ТекущаяПалитра = p: Exit Function
            Next p
        End If
    Next cel
End Function

Private Function ПерекраситьЛист(ws As Worksheet, ByVal n As Long) As Boolean

    Dim доли As Variant, старые As New Collection, новые As New Collection
    Dim p As Long, i As Long, заг As Long, шап As Long, новЗаг As Long, новШап As Long
    Dim ur As Range, rw As Range, cel As Range, v As Variant
    Dim нашли As Boolean

    If ТекущаяПалитра(ws) = 0 Then Exit Function

    доли = ДолиЗаливки()
    ЦветаПалитры n, новЗаг, новШап

    ' таблица «старый цвет -> новый»; первая запись выигрывает
    On Error Resume Next
    For p = 1 To ЧислоПалитр()
        ЦветаПалитры p, заг, шап
        For i = LBound(доли) To UBound(доли)
            новые.Add Оттенок(новШап, доли(i)), "c" & Оттенок(шап, доли(i))
        Next i
        старые.Add новЗаг, "t" & заг
    Next p
    On Error GoTo 0

    Application.ScreenUpdating = False
    Set ur = ws.UsedRange
    For Each rw In ur.Rows
        v = rw.Interior.Color
        If Not IsNull(v) Then
            ' вся строка одного цвета — одним вызовом
            If rw.Cells(1, 1).Interior.Pattern <> xlNone Then
                If ЕстьКлюч(новые, "c" & v) Then rw.Interior.Color = новые("c" & v): нашли = True
            End If
        Else
            For Each cel In rw.Cells
                If cel.Interior.Pattern <> xlNone Then
                    v = cel.Interior.Color
                    If ЕстьКлюч(новые, "c" & v) Then cel.Interior.Color = новые("c" & v): нашли = True
                End If
            Next cel
        End If
    Next rw
    ' заголовок отчёта — цветом палитры
    For Each cel In ws.Range(ws.Cells(1, 1), ws.Cells(3, 4)).Cells
        v = cel.Font.Color
        If Not IsNull(v) Then
            If ЕстьКлюч(старые, "t" & v) Then cel.Font.Color = старые("t" & v)
        End If
    Next cel
    Application.ScreenUpdating = True
    ПерекраситьЛист = нашли

End Function

Private Function ЕстьКлюч(c As Collection, ByVal k As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(k)
    ЕстьКлюч = (Err.Number = 0)
    Err.Clear
End Function


'==========================================================================
'  ПОИСК СВОДНОЙ
'==========================================================================
Private Function ПолучитьСводную() As PivotTable

    Dim ws As Worksheet, i As Long, n As Long
    Dim spisok As String, otvet As String

    If TypeName(ActiveSheet) <> "Worksheet" Then
        Сообщить "Активный лист не является рабочим листом."
        Exit Function
    End If
    Set ws = ActiveSheet

    ' 1) курсор внутри сводной
    On Error Resume Next
    Set ПолучитьСводную = ActiveCell.PivotTable
    On Error GoTo 0
    If Not ПолучитьСводную Is Nothing Then Exit Function

    ' 2) единственная сводная на листе
    n = ws.PivotTables.Count
    If n = 0 Then
        Сообщить "На листе """ & ws.Name & """ нет сводных таблиц." & vbLf & _
                 "Откройте лист со сводной и запустите макрос ещё раз."
        Exit Function
    End If
    If n = 1 Then
        Set ПолучитьСводную = ws.PivotTables(1)
        Exit Function
    End If

    ' 3) несколько — спрашиваем
    If mQuiet Then
        Set ПолучитьСводную = ws.PivotTables(1)
        Exit Function
    End If
    For i = 1 To n
        spisok = spisok & i & " — " & ws.PivotTables(i).Name & _
                 "   (" & ws.PivotTables(i).TableRange1.Address(False, False) & ")" & vbLf
    Next i
    otvet = InputBox("На листе несколько сводных таблиц. Введите номер нужной:" & _
                     vbLf & vbLf & spisok, "Выбор сводной таблицы", "1")
    If Not IsNumeric(otvet) Then Exit Function
    i = CLng(otvet)
    If i < 1 Or i > n Then Exit Function
    Set ПолучитьСводную = ws.PivotTables(i)

End Function


'==========================================================================
'  РАЗБОР СВОДНОЙ — снятие данных с листа
'==========================================================================
Private Function РазобратьСводную(pt As PivotTable, _
                                  ByRef lbl() As Variant, _
                                  ByRef vals() As Variant) As Boolean

    Dim rr As Range, db As Range
    Dim i As Long, k As Long, j As Long, смещСтрок As Long

    On Error Resume Next
    Set db = pt.DataBodyRange
    Set rr = pt.RowRange
    On Error GoTo 0

    If db Is Nothing Then
        Сообщить "В сводной нет значений: добавьте хотя бы одно поле в область «Значения»."
        Exit Function
    End If

    nR = db.Rows.Count
    nC = db.Columns.Count
    If nR = 0 Or nC = 0 Then Exit Function

    Set mИсходник = db
    vals = ЧитатьДиапазон(db, nR, nC)

    If rr Is Nothing Then
        ' Сводная без полей в строках: одна строка общего итога
        nLbl = 1
        ReDim lbl(1 To nR, 1 To 1)
        For i = 1 To nR
            lbl(i, 1) = TXT_GRAND
        Next i
        ReDim mLblCap(1 To 1)
        mLblCap(1) = ""
        СнятьСтрокиБезПодписей
    Else
        nLbl = rr.Columns.Count
        ' RowRange включает сверху строку подписей полей — снимаем сдвиг,
        ' иначе подписи разъедутся с числами на одну строку.
        смещСтрок = db.Row - rr.Row
        If смещСтрок < 0 Then смещСтрок = 0
        If смещСтрок + nR > rr.Rows.Count Then смещСтрок = rr.Rows.Count - nR
        If смещСтрок < 0 Then смещСтрок = 0
        Set rr = rr.Offset(смещСтрок, 0).Resize(nR, nLbl)
        lbl = ЧитатьДиапазон(rr, nR, nLbl)

        ' подписи столбцов уровней строк — в строке прямо над данными
        ReDim mLblCap(1 To nLbl)
        For k = 1 To nLbl
            mLblCap(k) = ""
            On Error Resume Next
            mLblCap(k) = ВСтроку(db.Worksheet.Cells(db.Row - 1, rr.Column + k - 1).Value)
            On Error GoTo 0
            If Len(Trim$(mLblCap(k))) = 0 Then mLblCap(k) = "Уровень " & k
        Next k

        СнятьСтроки pt, rr, lbl
    End If

    ReDim inHasVal(1 To nR)
    For i = 1 To nR
        For j = 1 To nC
            If Not ПустоеЗначение(vals(i, j)) Then inHasVal(i) = True: Exit For
        Next j
    Next i

    СнятьСтолбцы pt
    СобратьШапку

    СнятьФорматы db
    mФильтры = ТекстФильтров(pt)

    РазобратьСводную = True

End Function

Private Sub СнятьСтрокиБезПодписей()
    Dim i As Long
    ReDim inTxt(1 To nR, 1 To 1)
    ReDim inCT(1 To nR, 1 To 1)
    ReDim inLev(1 To nR, 1 To 1)
    ReDim inInd(1 To nR, 1 To 1)
    ReDim inLT(1 To nR)
    For i = 1 To nR
        inTxt(i, 1) = TXT_GRAND
        inCT(i, 1) = CT_GRAND
        inLT(i) = LT_GRAND
    Next i
End Sub

' Для каждой ячейки подписей: текст, отступ и — главное — что о ней
' знает сама сводная: тип ячейки и поле, к которому она относится.
' Поле однозначно задаёт уровень в любом макете (сжатом, структурном,
' табличном, смешанном), а отступ и столбец — запасной вариант.
Private Sub СнятьСтроки(pt As PivotTable, rr As Range, ByRef lbl() As Variant)

    Dim i As Long, k As Long, n As Long, f As Long
    Dim поля() As String, имяДанных As String, имя As String
    Dim cel As Range, ct As Long, пусто As Boolean
    Dim ax As Object, ok As Boolean
    Dim образец As Long, полный As Boolean

    ReDim inTxt(1 To nR, 1 To nLbl)
    ReDim inCT(1 To nR, 1 To nLbl)
    ReDim inLev(1 To nR, 1 To nLbl)
    ReDim inInd(1 To nR, 1 To nLbl)
    ReDim inLT(1 To nR)

    n = 0
    On Error Resume Next
    n = pt.RowFields.Count
    If n > 0 Then
        ReDim поля(1 To n)
        For f = 1 To n
            поля(f) = pt.RowFields(f).Name
        Next f
    End If
    имяДанных = ""
    имяДанных = pt.DataPivotField.Name
    Err.Clear
    On Error GoTo 0

    ' типы линий из оси сводной — самый надёжный признак итога
    ok = False
    On Error Resume Next
    Set ax = pt.PivotRowAxis
    If Not ax Is Nothing Then ok = (ax.PivotLines.Count = nR)
    For i = 1 To nR
        inLT(i) = -1
        If ok Then
            inLT(i) = ax.PivotLines(i).LineType
            If Err.Number <> 0 Then inLT(i) = -1: Err.Clear
        End If
    Next i
    On Error GoTo 0

    ' Опрос сводной по каждой ячейке — сотни тысяч обращений на большом
    ' отчёте. Поэтому полностью опрашиваем первые строки; если в них каждое
    ' место на листе (столбец + отступ) занимает ровно одно поле, уровень
    ' однозначно читается по месту, и дальше хватает одного отступа.
    образец = nR
    If ok And nR > 300 Then образец = 300
    полный = True
    For i = 1 To nR
        If i = образец + 1 Then полный = Not МестоОднозначно(образец)
        пусто = True
        For k = 1 To nLbl
            inTxt(i, k) = ВСтроку(lbl(i, k))
            If Len(Trim$(inTxt(i, k))) > 0 Then пусто = False
        Next k
        For k = 1 To nLbl
            inCT(i, k) = -1
            ' по пустым ячейкам спрашиваем сводную только в строке без подписей
            ' (пустой элемент) — иначе на больших отчётах это долго
            If (Len(Trim$(inTxt(i, k))) > 0 And Not полный) And Not пусто Then
                On Error Resume Next
                inInd(i, k) = rr.Cells(i, k).IndentLevel
                On Error GoTo 0
            ElseIf Len(Trim$(inTxt(i, k))) > 0 Or пусто Then
                Set cel = rr.Cells(i, k)
                On Error Resume Next
                inInd(i, k) = cel.IndentLevel
                ct = -1
                ct = cel.PivotCell.PivotCellType
                If Err.Number <> 0 Then ct = -1: Err.Clear
                inCT(i, k) = ct
                имя = ""
                If ct = CT_ITEM Or ct = CT_SUBTOTAL Or ct = CT_CUSTOMSUB Then
                    имя = cel.PivotCell.PivotField.Name
                    If Err.Number <> 0 Then
                        Err.Clear
                        имя = cel.PivotCell.PivotItem.Parent.Name
                        If Err.Number <> 0 Then имя = "": Err.Clear
                    End If
                ElseIf ct = CT_DATAFIELD Then
                    имя = имяДанных
                End If
                On Error GoTo 0
                If Len(имя) > 0 Then
                    For f = 1 To n
                        If StrComp(поля(f), имя, vbTextCompare) = 0 Then inLev(i, k) = f: Exit For
                    Next f
                End If
            End If
        Next k
    Next i

End Sub

' В образце каждое место (столбец + отступ) занято одним полем, и каждое
' поле живёт в одном месте? Тогда уровень по месту равен уровню по полю.
Private Function МестоОднозначно(ByVal образец As Long) As Boolean
    Dim i As Long, k As Long, n As Long, a As Long
    Dim место() As Long, поле() As Long, м As Long
    ReDim место(1 To образец * nLbl + 1)
    ReDim поле(1 To образец * nLbl + 1)
    For i = 1 To образец
        For k = 1 To nLbl
            If Len(Trim$(inTxt(i, k))) > 0 And inCT(i, k) <> CT_GRAND Then
                If inLev(i, k) <= 0 Then Exit Function       ' сводная не ответила
                м = k * 1000 + inInd(i, k)
                For a = 1 To n
                    If место(a) = м Xor поле(a) = inLev(i, k) Then Exit Function
                    If место(a) = м Then Exit For
                Next a
                If a > n Then n = n + 1: место(n) = м: поле(n) = inLev(i, k)
            End If
        Next k
    Next i
    МестоОднозначно = True
End Function

' Пустое значение сводной: пусто, ошибка или пустая строка
Private Function ПустоеЗначение(v As Variant) As Boolean
    If IsEmpty(v) Or IsError(v) Then ПустоеЗначение = True: Exit Function
    If Not IsNumeric(v) Then
        If Len(Trim$(ВСтроку(v))) = 0 Then ПустоеЗначение = True
    End If
End Function

' Числовые форматы источника: обычно у каждой меры (колонки) свой,
' а если меры стоят в строках — у каждой строки.
Private Sub СнятьФорматы(db As Range)
    Dim j As Long, i As Long, v As Variant, поКол As Boolean
    ReDim mФорматКол(1 To nC)
    ReDim mФорматСтр(1 To nR)
    поКол = True
    On Error Resume Next
    For j = 1 To nC
        v = Null
        v = db.Columns(j).NumberFormat
        If IsNull(v) Then поКол = False Else mФорматКол(j) = CStr(v)
    Next j
    If Not поКол Then
        For j = 1 To nC
            mФорматКол(j) = ""
        Next j
        For i = 1 To nR
            v = Null
            v = db.Rows(i).NumberFormat
            If Not IsNull(v) Then mФорматСтр(i) = CStr(v)
        Next i
    End If
    On Error GoTo 0
End Sub

' "Регион: Север; Год: 2024" — значения фильтров сводной
Private Function ТекстФильтров(pt As PivotTable) As String
    Dim pf As PivotField, s As String, t As String, n As Long
    On Error Resume Next
    n = pt.PageFields.Count
    If n = 0 Then Exit Function
    For Each pf In pt.PageFields
        t = ""
        If pf.EnableMultiplePageItems Then
            If Not ВыбранОдин(pf, t) Then t = "(несколько элементов)"
        End If
        If Len(t) = 0 Then
            t = pf.CurrentPage.Caption
            If Err.Number <> 0 Then Err.Clear: t = ""
        End If
        If Len(t) = 0 Then
            t = ЧистоеИмя(pf.CurrentPageName)
            If Err.Number <> 0 Then Err.Clear: t = ""
        End If
        If Len(t) = 0 Then t = "(все)"
        If LCase$(t) <> "(все)" And LCase$(t) <> "(all)" Then
            If Len(s) > 0 Then s = s & ";   "
            s = s & ЧистоеИмя(pf.Caption) & ": " & ЧистоеИмя(t)
        End If
    Next pf
    On Error GoTo 0
    ТекстФильтров = s
End Function

Private Function ВыбранОдин(pf As PivotField, ByRef t As String) As Boolean
    Dim it As PivotItem, n As Long, всего As Long
    On Error GoTo Нет
    For Each it In pf.PivotItems
        всего = всего + 1
        If it.Visible Then n = n + 1: t = it.Caption
    Next it
    If n = всего Then t = "": ВыбранОдин = True: Exit Function   ' выбраны все
    ВыбранОдин = (n = 1)
    If Not ВыбранОдин Then t = ""
    Exit Function
Нет:
    t = ""
End Function


'--------------------------------------------------------------------------
'  Столбцы: снятие оси столбцов
'--------------------------------------------------------------------------
Private Sub СнятьСтолбцы(pt As PivotTable)

    Dim ax As Object, ln As Object
    Dim j As Long, k As Long, ok As Boolean, maxLev As Long, d As String

    ReDim cType(1 To nC)
    ReDim cCnt(1 To nC)
    ReDim cDF(1 To nC)
    cMeasLev = 0

    nMeas = 0
    On Error Resume Next
    nMeas = pt.DataFields.Count
    If nMeas > 0 Then
        ReDim mMeas(1 To nMeas)
        For j = 1 To nMeas
            mMeas(j) = ЧистоеИмя(pt.DataFields(j).Name)
        Next j
    Else
        ReDim mMeas(1 To 1)
    End If
    Err.Clear
    On Error GoTo 0

    ok = False
    On Error Resume Next
    Set ax = pt.PivotColumnAxis
    If Not ax Is Nothing Then ok = (ax.PivotLines.Count = nC)
    Err.Clear
    On Error GoTo 0

    If Not ok Then
        СнятьСтолбцыПоЛисту pt
        Exit Sub
    End If

    maxLev = 1
    On Error Resume Next
    For j = 1 To nC
        cType(j) = ax.PivotLines(j).LineType
        cCnt(j) = ax.PivotLines(j).PivotLineCells.Count
        If cCnt(j) < 1 Then cCnt(j) = 1
        If cCnt(j) > maxLev Then maxLev = cCnt(j)
    Next j
    Err.Clear
    On Error GoTo 0

    ReDim cItem(1 To nC, 1 To maxLev)
    On Error Resume Next
    For j = 1 To nC
        Set ln = ax.PivotLines(j)
        For k = 1 To cCnt(j)
            cItem(j, k) = ТекстЭлемента(ln, k)
            d = ""
            d = ln.PivotLineCells(k).DataField.Name
            If Err.Number <> 0 Then d = "": Err.Clear
            If Len(d) > 0 Then
                cDF(j) = ЧистоеИмя(d)
                ' уровень мер определяем только по полноразмерным линиям
                If cCnt(j) = maxLev And cMeasLev = 0 Then cMeasLev = k
            End If
        Next k
    Next j
    Err.Clear
    On Error GoTo 0

End Sub

' Подпись элемента оси. У сводных на модели данных Caption часто пуст,
' а Name возвращает уникальное имя вида [ginek].[КССКП новый].&[КС] —
' из него берём то, что в последних скобках.
Private Function ТекстЭлемента(ln As Object, ByVal k As Long) As String
    Dim s As String
    On Error Resume Next
    Err.Clear
    s = ln.PivotLineCells(k).PivotItem.Caption
    If Err.Number <> 0 Then Err.Clear: s = ""
    If Len(s) = 0 Then
        Err.Clear
        s = ln.PivotLineCells(k).PivotItem.Name
        If Err.Number <> 0 Then Err.Clear: s = ""
    End If
    On Error GoTo 0
    ТекстЭлемента = ЧистоеИмя(s)
End Function

' Запасной разбор шапки — по тексту и типам ячеек на листе.
' Подпись группы стоит только над первым столбцом группы, поэтому пустая
' клетка наследует подпись слева; итог «Север Итог» относится к группе
' «Север», которая стоит прямо перед ним.
Private Sub СнятьСтолбцыПоЛисту(pt As PivotTable)

    Dim cr As Range, db As Range, j As Long, k As Long, r As Long, s As String
    Dim строк As Long, верх As Long, ct As Long
    Dim стек() As String, послНепуст As Long, база As String
    Dim нужна() As Boolean, t As Long, сырые() As String, типы() As Long, q As Long

    On Error Resume Next
    Set cr = pt.ColumnRange
    Set db = pt.DataBodyRange
    On Error GoTo 0
    If cr Is Nothing Then
        ReDim cItem(1 To nC, 1 To 1)
        For j = 1 To nC
            cItem(j, 1) = mMeas(1)
            If nMeas > 1 And j <= nMeas Then cItem(j, 1) = mMeas(j)
            cCnt(j) = 1
            cType(j) = LT_REGULAR
        Next j
        Exit Sub
    End If

    верх = cr.Row
    строк = db.Row - верх
    If строк < 1 Then строк = 1: верх = db.Row - 1
    ReDim сырые(1 To строк, 1 To nC)
    ReDim типы(1 To строк, 1 To nC)
    ReDim нужна(1 To строк)

    ' строки шапки, где только кнопки полей и надписи «Названия столбцов»,
    ' уровня не задают — выкидываем
    For r = 1 To строк
        For j = 1 To nC
            сырые(r, j) = ВСтроку(db.Worksheet.Cells(верх + r - 1, db.Column + j - 1).Value)
            ct = -1
            On Error Resume Next
            ct = db.Worksheet.Cells(верх + r - 1, db.Column + j - 1).PivotCell.PivotCellType
            If Err.Number <> 0 Then ct = -1: Err.Clear
            On Error GoTo 0
            типы(r, j) = ct
            If ct = CT_ITEM Or ct = CT_SUBTOTAL Or ct = CT_GRAND Or ct = CT_DATAFIELD _
               Or ct = CT_CUSTOMSUB Then нужна(r) = True
            If ct = -1 And Len(Trim$(сырые(r, j))) > 0 Then нужна(r) = True
        Next j
    Next r
    t = 0
    For r = 1 To строк
        If нужна(r) Then t = t + 1
    Next r
    If t = 0 Then t = 1: нужна(строк) = True

    ReDim cItem(1 To nC, 1 To t)
    ReDim стек(1 To t)
    For j = 1 To nC
        cType(j) = LT_REGULAR
        k = 0
        послНепуст = 0
        For r = 1 To строк
            If нужна(r) Then
                k = k + 1
                s = сырые(r, j)
                ct = типы(r, j)
                If ct = CT_GRAND Then cType(j) = LT_GRAND
                If ct = CT_SUBTOTAL Or ct = CT_CUSTOMSUB Then cType(j) = LT_SUBTOTAL
                If ct = -1 Then
                    If ЭтоОбщийИтог(s) Then
                        cType(j) = LT_GRAND
                    ElseIf Len(СуффиксИтога(s)) > 0 Then
                        cType(j) = LT_SUBTOTAL
                    End If
                End If
                If Len(Trim$(s)) > 0 Then послНепуст = k
                If cType(j) <> LT_REGULAR And Len(Trim$(s)) > 0 Then Exit For
                If Len(Trim$(s)) > 0 Then
                    стек(k) = s
                    For q = k + 1 To t
                        стек(q) = ""
                    Next q
                End If
                cItem(j, k) = стек(k)
            End If
        Next r
        If cType(j) = LT_SUBTOTAL Then
            ' путь итога — это подписи его группы
            база = СуффиксИтога(s)
            If Len(база) = 0 Then база = стек(k)
            For q = 1 To k - 1
                cItem(j, q) = стек(q)
            Next q
            cItem(j, k) = база
            cCnt(j) = k
        ElseIf cType(j) = LT_GRAND Then
            cItem(j, 1) = TXT_GRAND
            cCnt(j) = 1
        Else
            cCnt(j) = послНепуст
            If cCnt(j) < 1 Then cCnt(j) = 1
            For q = 1 To cCnt(j)
                cItem(j, q) = стек(q)
            Next q
        End If
        ' мера столбца — если в подписи есть её имя
        For q = 1 To nMeas
            If Len(mMeas(q)) > 0 And InStr(1, s, mMeas(q), vbTextCompare) > 0 Then cDF(j) = mMeas(q)
        Next q
    Next j

End Sub


'==========================================================================
'  ПЕЧАТЬ ФИЛЬТРОВ, МЕЛОЧИ РАЗБОРА
'==========================================================================
' Уникальное имя OLAP вида [ginek].[КССКП новый].&[КС] или
' [Measures].[Число элементов] сводим к тому, что в последних скобках.
' Обычный текст не трогаем: требуем именно форму [..].[..]
Private Function ЧистоеИмя(ByVal s As String) As String
    Dim p As Long
    ЧистоеИмя = s
    If Left$(s, 1) <> "[" Then Exit Function
    If Right$(s, 1) <> "]" Then Exit Function
    If InStr(s, "].[") = 0 And InStr(s, "].&[") = 0 Then Exit Function
    p = InStrRev(s, "&[")
    If p > 0 Then
        ЧистоеИмя = Mid$(s, p + 2, Len(s) - p - 2)
        Exit Function
    End If
    p = InStrRev(s, "[")
    If p > 0 Then ЧистоеИмя = Mid$(s, p + 1, Len(s) - p - 1)
End Function


'#ЯДРО-НАЧАЛО
'==========================================================================
'  ЯДРО: СТРУКТУРА СТРОК
'  Не обращается к Excel — работает только с массивами in*, поэтому
'  одинаково ведёт себя на любом макете и проверяется автотестами.
'==========================================================================
Public Sub ЯдроСтроки()

    Dim i As Long, k As Long, q As Long, lv As Long, глуб As Long
    Dim посл As Long, s As String, база As String, прежГлуб As Long
    Dim стек() As String, стекR() As Long, стекC() As Long

    ReDim mType(1 To nR)
    ReDim mDepth(1 To nR)
    ReDim mDrop(1 To nR)
    ReDim mOwnCustom(1 To nR)

    For i = 1 To nR
        mType(i) = ТипСтроки(i)
    Next i

    ПостроитьУровни

    ReDim mKey(1 To nR, 1 To mMaxD)
    ReDim mKeyR(1 To nR, 1 To mMaxD)
    ReDim mKeyC(1 To nR, 1 To mMaxD)
    ReDim mPath(1 To nR, 1 To mMaxD)
    ReDim стек(1 To mMaxD)
    ReDim стекR(1 To mMaxD)
    ReDim стекC(1 To mMaxD)

    прежГлуб = 0
    For i = 1 To nR
        If mType(i) = LT_GRAND Or mType(i) = LT_BLANK Then
            mDepth(i) = 1
            GoTo СледСтрока
        End If

        посл = 0
        For k = nLbl To 1 Step -1
            If mCellLev(i, k) > 0 Then посл = k: Exit For
        Next k

        If посл = 0 Then
            ' Строка без подписи — пустой элемент. Уровень знает сводная;
            ' если нет — это ребёнок предыдущей строки (в табличном макете
            ' предыдущая строка уже на самом глубоком уровне).
            lv = ПустойУровень(i)
            If lv = 0 Then lv = прежГлуб + 1
            If lv > mMaxD Then lv = mMaxD
            If lv < 1 Then lv = 1
            стек(lv) = TXT_EMPTY: стекR(lv) = 0: стекC(lv) = 0
            For q = lv + 1 To mMaxD
                стек(q) = "": стекR(q) = 0: стекC(q) = 0
            Next q
            глуб = lv
        Else
            глуб = 0
            For k = 1 To посл
                lv = mCellLev(i, k)
                If lv > 0 Then
                    s = inTxt(i, k)
                    If k = посл And mType(i) = LT_SUBTOTAL Then
                        ' «Север Итог» — итог группы «Север», которая только
                        ' что закончилась (или начнётся, если итоги вверху)
                        база = СуффиксИтога(s)
                        If Len(база) > 0 Then
                            If StrComp(база, стек(lv), vbTextCompare) <> 0 Then
                                стек(lv) = база: стекR(lv) = 0: стекC(lv) = 0
                            End If
                        ElseIf Len(стек(lv)) = 0 Then
                            стек(lv) = s: стекR(lv) = i: стекC(lv) = k
                        End If
                        If inCT(i, k) = CT_CUSTOMSUB Then mOwnCustom(i) = True
                    Else
                        стек(lv) = s: стекR(lv) = i: стекC(lv) = k
                    End If
                    For q = lv + 1 To mMaxD
                        стек(q) = "": стекR(q) = 0: стекC(q) = 0
                    Next q
                    If lv > глуб Then глуб = lv
                End If
            Next k
        End If

        mDepth(i) = глуб
        For k = 1 To глуб
            mKey(i, k) = стек(k)
            mKeyR(i, k) = стекR(k)
            mKeyC(i, k) = стекC(k)
            If k = 1 Then
                mPath(i, k) = LCase$(стек(k))
            Else
                mPath(i, k) = mPath(i, k - 1) & Chr$(1) & LCase$(стек(k))
            End If
        Next k
        прежГлуб = глуб
СледСтрока:
    Next i

    ПоглотитьЗаголовки

End Sub

' Тип строки: из оси сводной, иначе из типов ячеек, иначе по тексту
Private Function ТипСтроки(ByVal i As Long) As Long
    Dim k As Long, есть As Boolean, известно As Boolean, t As Long, s As String

    For k = 1 To nLbl
        If Len(Trim$(inTxt(i, k))) > 0 Then есть = True
    Next k

    t = -1
    If inLT(i) >= 0 Then
        t = inLT(i)
    Else
        For k = 1 To nLbl
            Select Case inCT(i, k)
                Case CT_GRAND: t = LT_GRAND
                Case CT_SUBTOTAL, CT_CUSTOMSUB: If t <> LT_GRAND Then t = LT_SUBTOTAL
            End Select
            If inCT(i, k) >= 0 Then известно = True
        Next k
        If t < 0 And Not известно Then
            For k = nLbl To 1 Step -1
                s = inTxt(i, k)
                If Len(Trim$(s)) > 0 Then
                    If ЭтоОбщийИтог(s) Then
                        t = LT_GRAND
                    ElseIf Len(СуффиксИтога(s)) > 0 Then
                        t = LT_SUBTOTAL
                    End If
                    Exit For
                End If
            Next k
        End If
        If t < 0 Then t = LT_REGULAR
    End If

    ' пустая строка-разделитель («пустая строка после каждого элемента»)
    If t <> LT_GRAND And Not есть And Not inHasVal(i) Then t = LT_BLANK
    If t = LT_BLANK And есть Then t = LT_REGULAR
    ТипСтроки = t
End Function

Public Function ЭтоОбщийИтог(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    ЭтоОбщийИтог = (s = "общий итог" Or s = "grand total" Or s = LCase$(TXT_GRAND))
End Function

' «Север Итог» -> «Север»; не итог -> ""
Public Function СуффиксИтога(ByVal s As String) As String
    Dim t As String, x As Variant
    t = Trim$(s)
    For Each x In Array(" итог", " total", " всего", " итого")
        If Len(t) > Len(x) Then
            If StrComp(Right$(t, Len(x)), x, vbTextCompare) = 0 Then
                СуффиксИтога = Trim$(Left$(t, Len(t) - Len(x)))
                Exit Function
            End If
        End If
    Next x
End Function

' Уровень каждой ячейки подписей.
' Если сводная сказала, к какому полю относится каждая непустая ячейка, —
' уровень = порядковый номер поля. Иначе уровень задаёт место на листе:
' столбец, а внутри столбца — отступ (сжатый макет держит несколько полей
' в одном столбце и различает их только отступом).
Private Sub ПостроитьУровни()

    Dim i As Long, k As Long, j As Long, n As Long, v As Long
    Dim поПолям As Boolean, есть As Boolean
    Dim знач() As Long, сыр() As Long

    ReDim mCellLev(1 To nR, 1 To nLbl)
    ReDim сыр(1 To nR, 1 To nLbl)

    поПолям = True
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 Then
                    есть = True
                    If inLev(i, k) <= 0 Then поПолям = False
                End If
            Next k
        End If
    Next i
    If Not есть Then поПолям = False

    ' сырые координаты и их различные значения по возрастанию
    ReDim знач(1 To nR * nLbl + 1)
    n = 0
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 Then
                    If поПолям Then
                        v = inLev(i, k)
                    Else
                        v = k * 1000 + Мин(Макс(inInd(i, k), 0), 999)
                    End If
                    сыр(i, k) = v
                    ВставитьЗнач знач, n, v
                End If
            Next k
        End If
    Next i
    If n = 0 Then n = 1: знач(1) = 1000

    mMaxD = n
    ReDim mColOf(1 To mMaxD)
    ReDim mIndOf(1 To mMaxD)

    For i = 1 To nR
        For k = 1 To nLbl
            If сыр(i, k) > 0 Then
                For j = 1 To n
                    If знач(j) = сыр(i, k) Then Exit For
                Next j
                mCellLev(i, k) = j
                If mColOf(j) = 0 Then mColOf(j) = k
            End If
        Next k
    Next i
    For j = 1 To mMaxD
        If mColOf(j) = 0 Then mColOf(j) = 1
        If mColOf(j) > nLbl Then mColOf(j) = nLbl
    Next j
    ' уровень, ушедший в столбец левее предыдущего (бывает в смешанном
    ' макете), не должен ломать подписи: столбцы уровней не убывают
    For j = 2 To mMaxD
        If mColOf(j) < mColOf(j - 1) Then mColOf(j) = mColOf(j - 1)
    Next j
    ' отступ внутри столбца — сколько уровней выше живут в том же столбце
    For j = 1 To mMaxD
        mIndOf(j) = 0
        For k = 1 To j - 1
            If mColOf(k) = mColOf(j) Then mIndOf(j) = mIndOf(j) + 1
        Next k
    Next j

End Sub

Private Sub ВставитьЗнач(ByRef a() As Long, ByRef n As Long, ByVal v As Long)
    Dim i As Long, j As Long
    For i = 1 To n
        If a(i) = v Then Exit Sub
        If a(i) > v Then Exit For
    Next i
    For j = n To i Step -1
        a(j + 1) = a(j)
    Next j
    a(i) = v
    n = n + 1
End Sub

' Уровень пустого элемента, если его знает сводная
Private Function ПустойУровень(ByVal i As Long) As Long
    Dim k As Long, j As Long, v As Long, kk As Long
    For k = nLbl To 1 Step -1
        If inLev(i, k) > 0 Then
            v = inLev(i, k)
            ' перевод номера поля в номер уровня
            For j = 1 To nR
                For kk = 1 To nLbl
                    If mCellLev(j, kk) > 0 And inLev(j, kk) = v Then
                        ПустойУровень = mCellLev(j, kk)
                        Exit Function
                    End If
                Next kk
            Next j
        End If
    Next k
End Function

' Структурный макет с итогами внизу даёт для группы ДВЕ строки: пустой
' заголовок «Север» сверху и «Север Итог» снизу. В отчёте нужна одна —
' итоговая, поднятая наверх. Заголовок без значений поглощается итогом.
Private Sub ПоглотитьЗаголовки()
    Dim i As Long, j As Long, d As Long
    For i = 1 To nR
        If mType(i) = LT_SUBTOTAL And Not mOwnCustom(i) Then
            d = mDepth(i)
            For j = i - 1 To 1 Step -1
                If mType(j) = LT_GRAND Then Exit For
                If mType(j) <> LT_BLANK And Not mDrop(j) Then
                    If mDepth(j) <= d Then
                        If mDepth(j) = d And mType(j) = LT_REGULAR And Not inHasVal(j) Then
                            If mPath(j, d) = mPath(i, d) Then
                                mDrop(j) = True
                                ' подпись заголовка точнее, чем «Север» из «Север Итог»
                                mKey(i, d) = mKey(j, d)
                                mKeyR(i, d) = mKeyR(j, d)
                                mKeyC(i, d) = mKeyC(j, d)
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
End Sub


'==========================================================================
'  ЯДРО: ПОРЯДОК СТРОК
'  Порядок групп — как в сводной (по первому появлению), итог уровня —
'  над своими детьми, общий итог — первой строкой. Сортировка устойчивая:
'  внутри группы строки не перемешиваются.
'==========================================================================
Public Sub ПорядокСтрок()

    Dim i As Long, k As Long, n As Long, счёт As Long, r As Long
    Dim ранг As New Collection
    Dim idx() As Long, ключ() As String, s As String

    ReDim mOrd(1 To nR)
    nOrd = 0

    ' общий итог — первой строкой
    For i = 1 To nR
        If mType(i) = LT_GRAND Then nOrd = nOrd + 1: mOrd(nOrd) = i
    Next i

    ReDim idx(1 To nR)
    ReDim ключ(1 To nR)
    n = 0
    счёт = 0
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK And Not mDrop(i) Then
            n = n + 1
            idx(n) = i
            s = ""
            For k = 1 To mMaxD
                If k <= mDepth(i) Then
                    r = ВзятьРанг(ранг, mPath(i, k))
                    If r = 0 Then
                        счёт = счёт + 1
                        r = счёт
                        ранг.Add r, "k" & mPath(i, k)
                    End If
                Else
                    r = 0          ' своя строка группы — раньше детей
                End If
                s = s & Format$(r, "0000000")
            Next k
            ключ(n) = s & Format$(i, "0000000")
        End If
    Next i

    If n > 0 Then СортироватьСлиянием idx, ключ, n
    For i = 1 To n
        nOrd = nOrd + 1
        mOrd(nOrd) = idx(i)
    Next i

End Sub

Private Function ВзятьРанг(c As Collection, ByVal k As String) As Long
    On Error Resume Next
    ВзятьРанг = c("k" & k)
    If Err.Number <> 0 Then ВзятьРанг = 0
    Err.Clear
End Function

' Устойчивая сортировка слиянием индексов по строковым ключам
Private Sub СортироватьСлиянием(ByRef idx() As Long, ByRef ключ() As String, ByVal n As Long)
    Dim tI() As Long, tK() As String
    Dim ширина As Long, лев As Long, сер As Long, пр As Long
    Dim a As Long, b As Long, c As Long
    ReDim tI(1 To n)
    ReDim tK(1 To n)
    ширина = 1
    Do While ширина < n
        лев = 1
        Do While лев <= n
            сер = лев + ширина - 1
            If сер > n Then сер = n
            пр = лев + 2 * ширина - 1
            If пр > n Then пр = n
            a = лев: b = сер + 1: c = лев
            Do While a <= сер And b <= пр
                If ключ(b) < ключ(a) Then
                    tI(c) = idx(b): tK(c) = ключ(b): b = b + 1
                Else
                    tI(c) = idx(a): tK(c) = ключ(a): a = a + 1
                End If
                c = c + 1
            Loop
            Do While a <= сер
                tI(c) = idx(a): tK(c) = ключ(a): a = a + 1: c = c + 1
            Loop
            Do While b <= пр
                tI(c) = idx(b): tK(c) = ключ(b): b = b + 1: c = c + 1
            Loop
            лев = лев + 2 * ширина
        Loop
        For a = 1 To n
            idx(a) = tI(a): ключ(a) = tK(a)
        Next a
        ширина = ширина * 2
    Loop
End Sub

' После упорядочивания: кто кому родитель, какие подписи показывать,
' уровень группировки строк.
Public Sub ЯдроПослеПорядка()

    Dim p As Long, i As Long, k As Long, d As Long, q As Long, c As Long
    Dim свой() As String, пред As Long, приставка As Long, лв As Long
    Dim txt As String

    ReDim mРодитель(1 To nR)
    ReDim mЕстьРодитель(1 To nR, 1 To mMaxD)
    ReDim mOutLbl(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutR(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutC(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutInd(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutLvl(1 To nOrd + 1)
    ReDim свой(1 To mMaxD)

    ' родитель — если следующей идёт строка той же группы глубже
    For p = 1 To nOrd
        i = mOrd(p)
        If mType(i) = LT_SUBTOTAL Then mРодитель(i) = True
        If mType(i) <> LT_GRAND And p < nOrd Then
            q = mOrd(p + 1)
            d = mDepth(i)
            If mType(q) <> LT_GRAND And mDepth(q) > d Then
                If mPath(q, d) = mPath(i, d) Then mРодитель(i) = True
            End If
        End If
    Next p

    ' есть ли над строкой отдельная строка её предка уровня k
    пред = 0
    For p = 1 To nOrd
        i = mOrd(p)
        For k = 1 To nLbl
            mOutInd(p, k) = -1
        Next k
        If mType(i) = LT_GRAND Then
            For k = 1 To mMaxD
                свой(k) = ""
            Next k
            mOutLbl(p, 1) = TXT_GRAND
            mOutLvl(p) = 1
            пред = 0
            GoTo След
        End If

        d = mDepth(i)
        свой(d) = mPath(i, d)
        For k = d + 1 To mMaxD
            свой(k) = ""
        Next k
        For k = 1 To d
            mЕстьРодитель(i, k) = (свой(k) = mPath(i, k) And Len(свой(k)) > 0)
        Next k

        ' собственная подпись
        c = mColOf(d)
        If mOwnCustom(i) Then
            txt = ""
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 Then txt = inTxt(i, k)
            Next k
            mOutLbl(p, c) = txt
        Else
            mOutLbl(p, c) = mKey(i, d)
            mOutR(p, c) = mKeyR(i, d)
            mOutC(p, c) = mKeyC(i, d)
        End If
        mOutInd(p, c) = mIndOf(d)

        ' подписи предков, у которых нет своей строки (табличный макет без
        ' итогов): показываем на первой строке группы
        For k = d - 1 To 1 Step -1
            If Not mЕстьРодитель(i, k) Then
                If пред = 0 Then
                    приставка = 1
                ElseIf mType(пред) = LT_GRAND Or mDepth(пред) < k Then
                    приставка = 1
                ElseIf mPath(пред, k) <> mPath(i, k) Then
                    приставка = 1
                Else
                    приставка = 0
                End If
                If приставка = 1 Then
                    If mColOf(k) <> c And Len(mOutLbl(p, mColOf(k))) = 0 Then
                        mOutLbl(p, mColOf(k)) = mKey(i, k)
                        mOutR(p, mColOf(k)) = mKeyR(i, k)
                        mOutC(p, mColOf(k)) = mKeyC(i, k)
                        mOutInd(p, mColOf(k)) = mIndOf(k)
                    Else
                        ' тот же столбец — дописываем путь перед подписью
                        mOutLbl(p, c) = mKey(i, k) & " / " & ТекстПодписи(p, c)
                        mOutR(p, c) = 0
                        mOutC(p, c) = 0
                    End If
                End If
            End If
        Next k

        ' уровень группировки: 1 + число предков, у которых есть своя строка
        лв = 1
        For k = 1 To d - 1
            If mЕстьРодитель(i, k) Then лв = лв + 1
        Next k
        If лв > 8 Then лв = 8
        mOutLvl(p) = лв
        пред = i
След:
    Next p

End Sub

Private Function ТекстПодписи(ByVal p As Long, ByVal c As Long) As String
    ТекстПодписи = mOutLbl(p, c)
End Function

Private Function Мин(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Мин = a Else Мин = b
End Function

Private Function Макс(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then Макс = a Else Макс = b
End Function


'==========================================================================
'  ЯДРО: ШАПКА СТОЛБЦОВ
'  Вход: cItem/cCnt/cType/cDF/cMeasLev/mMeas. Выход: dLev/dDepth/nLev.
'==========================================================================
Public Sub СобратьШапку()

    Dim j As Long, k As Long, maxLev As Long
    Dim мера() As String, выносить As Boolean
    Dim пор() As Long, np As Long, сдвиг As Long, задано As Long
    Dim счёт As Long, пусто As Boolean

    maxLev = 1
    For j = 1 To nC
        If cCnt(j) > maxLev Then maxLev = cCnt(j)
    Next j
    If UBound(cItem, 2) < maxLev Then maxLev = UBound(cItem, 2)

    ВосстановитьПутьИтогов cItem, cCnt, cType, nC, maxLev

    выносить = (cMeasLev > 0 And nMeas > 1)

    ' --- мера для каждого столбца ---
    ReDim мера(1 To nC)
    If выносить Then
        счёт = 0
        For j = 1 To nC
            If cType(j) = LT_SUBTOTAL Or (cType(j) = LT_GRAND And Len(cDF(j)) = 0) Then
                ' Excel выдаёт по одному столбцу итога на каждую меру подряд,
                ' в порядке мер — по позиции в серии и определяем меру
                If j > 1 Then
                    If cType(j - 1) = cType(j) And ОдинКлюч(cItem, cCnt, j - 1, j) Then
                        счёт = счёт + 1
                    Else
                        счёт = 1
                    End If
                Else
                    счёт = 1
                End If
                If счёт >= 1 And счёт <= nMeas Then мера(j) = mMeas(счёт)
            ElseIf Len(cDF(j)) > 0 Then
                мера(j) = cDF(j)
            ElseIf cMeasLev <= cCnt(j) Then
                мера(j) = cItem(j, cMeasLev)
            End If
        Next j
    End If

    ' --- уровни в порядке вывода: мера — вперёд ---
    ReDim пор(1 To maxLev)
    np = 0
    For k = 1 To maxLev
        If Not (выносить And k = cMeasLev) Then np = np + 1: пор(np) = k
    Next k

    If выносить Then nLev = np + 1 Else nLev = np
    If nLev < 1 Then nLev = 1
    If выносить Then сдвиг = 1 Else сдвиг = 0

    ReDim dLev(1 To nC, 1 To nLev)
    ReDim dDepth(1 To nC)

    For j = 1 To nC
        If выносить Then dLev(j, 1) = мера(j)
        If cType(j) = LT_GRAND Then
            dDepth(j) = сдвиг + 1
            If сдвиг + 1 <= nLev Then dLev(j, сдвиг + 1) = TXT_GRAND
        Else
            ' сколько «настоящих» уровней у столбца (без уровня меры)
            задано = 0
            For k = 1 To np
                If пор(k) <= cCnt(j) Then
                    dLev(j, сдвиг + k) = cItem(j, пор(k))
                    задано = k
                End If
            Next k
            dDepth(j) = сдвиг + задано
            If dDepth(j) < 1 Then dDepth(j) = 1
            If cType(j) = LT_SUBTOTAL And сдвиг + задано < nLev Then
                dLev(j, сдвиг + задано + 1) = TXT_SUBTOT
            End If
        End If
    Next j

    УбратьПустыеУровни

    ' Если у сводной нет полей в столбцах, шапка окажется пустой —
    ' подписываем колонку именем меры.
    For j = 1 To nC
        пусто = True
        For k = 1 To nLev
            If Len(Trim$(dLev(j, k))) > 0 Then пусто = False: Exit For
        Next k
        If пусто Then
            If Len(cDF(j)) > 0 Then
                dLev(j, 1) = cDF(j)
            ElseIf nMeas >= 1 Then
                dLev(j, 1) = mMeas(1)
            End If
            If Len(Trim$(dLev(j, 1))) = 0 Then dLev(j, 1) = "Значение"
        End If
    Next j

End Sub

' Уровень, пустой у всех столбцов (так бывает у безымянных членов OLAP),
' занимает строку в шапке и ничего не сообщает — выкидываем его.
Private Sub УбратьПустыеУровни()

    Dim k As Long, j As Long, n As Long, t As Long
    Dim держать() As Boolean
    Dim новый() As String, новГлуб() As Long

    If nLev <= 1 Then Exit Sub

    ReDim держать(1 To nLev)
    n = 0
    For k = 1 To nLev
        держать(k) = False
        For j = 1 To nC
            If Len(Trim$(dLev(j, k))) > 0 Then держать(k) = True: Exit For
        Next j
        If держать(k) Then n = n + 1
    Next k
    If n = nLev Or n = 0 Then Exit Sub

    ReDim новый(1 To nC, 1 To n)
    ReDim новГлуб(1 To nC)
    For j = 1 To nC
        t = 0
        новГлуб(j) = 0
        For k = 1 To nLev
            If держать(k) Then
                t = t + 1
                новый(j, t) = dLev(j, k)
                If k <= dDepth(j) Then новГлуб(j) = t
            End If
        Next k
        If новГлуб(j) < 1 Then новГлуб(j) = 1
    Next j

    nLev = n
    dLev = новый
    dDepth = новГлуб

End Sub

' Для линии промежуточного итога Excel отдаёт не путь предков, а сам
' итожимый элемент (на модели данных — ещё и продублированный). Из-за этого
' итог оказывается в отдельной группе и уезжает вправо. Восстанавливаем путь
' по БЛИЖАЙШЕЙ обычной линии с этим элементом: итог стоит вплотную к своей
' группе, а одноимённые элементы других родителей — дальше.
Public Sub ВосстановитьПутьИтогов(ByRef item() As String, ByRef cnt() As Long, _
                                  ByRef тип() As Long, ByVal n As Long, ByVal maxLev As Long)
    Dim j As Long, q As Long, k As Long, шаг As Long
    Dim посл As String, строка As Long, уровень As Long

    For j = 1 To n
        If тип(j) = LT_SUBTOTAL Then
            ' Обычная сводная уже отдаёт полный путь предков — трогать нельзя,
            ' иначе одноимённые элементы разных родителей склеятся в одну группу.
            If ПутьСогласован(item, cnt, тип, n, j) Then GoTo Дальше
            посл = item(j, cnt(j))
            строка = 0: уровень = 0
            For шаг = 1 To 2 * n
                ' сначала ближайшая слева, потом справа, расширяя окно
                If шаг Mod 2 = 1 Then q = j - (шаг + 1) \ 2 Else q = j + шаг \ 2
                    If q >= 1 And q <= n Then
                        If тип(q) = LT_REGULAR Then
                            For k = 1 To cnt(q)
                                If StrComp(item(q, k), посл, vbTextCompare) = 0 Then
                                    строка = q: уровень = k
                                    Exit For
                                End If
                            Next k
                        End If
                    End If
                If строка > 0 Then Exit For
            Next шаг
            If строка > 0 Then
                For k = 1 To maxLev
                    If k <= уровень Then item(j, k) = item(строка, k) Else item(j, k) = ""
                Next k
                cnt(j) = уровень
            End If
        End If
Дальше:
    Next j
End Sub

' Есть ли обычная линия, у которой первые cnt(j) уровней совпадают с путём
' линии итога. Если есть — путь настоящий, восстанавливать нечего.
Private Function ПутьСогласован(ByRef item() As String, ByRef cnt() As Long, _
                                ByRef тип() As Long, ByVal n As Long, _
                                ByVal j As Long) As Boolean
    Dim q As Long, k As Long, совпало As Boolean
    For q = 1 To n
        If тип(q) = LT_REGULAR And cnt(q) >= cnt(j) Then
            совпало = True
            For k = 1 To cnt(j)
                If StrComp(item(q, k), item(j, k), vbTextCompare) <> 0 Then
                    совпало = False
                    Exit For
                End If
            Next k
            If совпало Then ПутьСогласован = True: Exit Function
        End If
    Next q
End Function

Private Function ОдинКлюч(ByRef item() As String, ByRef cnt() As Long, _
                          ByVal a As Long, ByVal b As Long) As Boolean
    Dim k As Long
    If cnt(a) <> cnt(b) Then Exit Function
    For k = 1 To cnt(a)
        If StrComp(item(a, k), item(b, k), vbTextCompare) <> 0 Then Exit Function
    Next k
    ОдинКлюч = True
End Function


'==========================================================================
'  ЯДРО: ПОРЯДОК СТОЛБЦОВ — итог группы левее её столбцов
'==========================================================================
Public Sub ПорядокСтолбцов()
    Dim j As Long, ост() As Long, n As Long
    ReDim cOrd(1 To nC)
    nCOrd = 0

    ReDim ост(1 To nC)
    n = 0
    For j = 1 To nC
        If cType(j) <> LT_BLANK Then n = n + 1: ост(n) = j
    Next j
    If n > 0 Then ПорядокСтолбцовРек 1, ост, n
End Sub

Private Sub ПорядокСтолбцовРек(ByVal d As Long, ByRef idx() As Long, ByVal n As Long)

    Dim i As Long, k As Long
    Dim ключи() As String, nk As Long
    Dim свои() As Long, nс As Long
    Dim дети() As Long, nд As Long
    Dim найден As Boolean
    Dim пер As Long, ts As String

    If n = 0 Then Exit Sub
    If d > nLev Then
        For i = 1 To n
            nCOrd = nCOrd + 1: cOrd(nCOrd) = idx(i)
        Next i
        Exit Sub
    End If

    ' различные ключи уровня в порядке первого появления
    ReDim ключи(1 To n)
    nk = 0
    For i = 1 To n
        найден = False
        For k = 1 To nk
            If StrComp(dLev(idx(i), d), ключи(k), vbTextCompare) = 0 Then найден = True: Exit For
        Next k
        If Not найден Then nk = nk + 1: ключи(nk) = dLev(idx(i), d)
    Next i

    ' группы, целиком состоящие из общего итога, выносим в начало уровня
    пер = 0
    For k = 1 To nk
        If ТолькоИтог(idx, n, d, ключи(k)) Then
            пер = пер + 1
            If k > пер Then
                ts = ключи(k)
                For i = k To пер + 1 Step -1
                    ключи(i) = ключи(i - 1)
                Next i
                ключи(пер) = ts
            End If
        End If
    Next k

    For k = 1 To nk
        ReDim свои(1 To n): ReDim дети(1 To n)
        nс = 0: nд = 0
        For i = 1 To n
            If StrComp(dLev(idx(i), d), ключи(k), vbTextCompare) = 0 Then
                If dDepth(idx(i)) <= d Then
                    nс = nс + 1: свои(nс) = idx(i)      ' итог уровня — впереди
                Else
                    nд = nд + 1: дети(nд) = idx(i)
                End If
            End If
        Next i
        For i = 1 To nс
            nCOrd = nCOrd + 1: cOrd(nCOrd) = свои(i)
        Next i
        If nд > 0 Then ПорядокСтолбцовРек d + 1, дети, nд
    Next k

End Sub

Private Function ТолькоИтог(ByRef idx() As Long, ByVal n As Long, _
                            ByVal d As Long, ByVal ключ As String) As Boolean
    Dim i As Long, есть As Boolean
    For i = 1 To n
        If StrComp(dLev(idx(i), d), ключ, vbTextCompare) = 0 Then
            есть = True
            If cType(idx(i)) <> LT_GRAND Then Exit Function
        End If
    Next i
    ТолькоИтог = есть
End Function
'#ЯДРО-КОНЕЦ


'==========================================================================
'  ЛИСТ-ПРИЁМНИК
'==========================================================================
Private Function СоздатьЛист(pt As PivotTable) As Worksheet

    Dim otvet As VbMsgBoxResult
    Dim wb As Workbook
    Dim baza As String, imya As String, i As Long

    If mQuiet Then
        otvet = vbYes
    Else
        otvet = MsgBox("Куда положить отчёт?" & vbLf & vbLf & _
                       "Да — новый лист в этой книге" & vbLf & _
                       "Нет — новая книга", _
                       vbYesNoCancel + vbQuestion, "Размещение отчёта")
    End If
    If otvet = vbCancel Then Exit Function

    If otvet = vbYes Then
        Set wb = pt.Parent.Parent
    Else
        Set wb = Workbooks.Add(xlWBATWorksheet)
    End If

    baza = Left$(ЧистоеИмяЛиста(pt.Parent.Name), 24) & "_отчет"
    imya = baza
    i = 1
    Do While ЛистЕсть(wb, imya)
        i = i + 1
        imya = baza & CStr(i)
    Loop

    If otvet = vbNo Then
        Set СоздатьЛист = wb.Worksheets(1)
    Else
        Set СоздатьЛист = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    End If
    СоздатьЛист.Name = imya

End Function

Private Function ЧистоеИмяЛиста(ByVal s As String) As String
    Dim x As Variant
    For Each x In Array(":", "\", "/", "?", "*", "[", "]")
        s = Replace(s, x, "_")
    Next x
    ЧистоеИмяЛиста = s
End Function

Private Function ЛистЕсть(wb As Workbook, ByVal imya As String) As Boolean
    Dim ws As Object
    For Each ws In wb.Sheets
        If StrComp(ws.Name, imya, vbTextCompare) = 0 Then ЛистЕсть = True: Exit Function
    Next ws
End Function


'==========================================================================
'  ВЫВОД И ОФОРМЛЕНИЕ
'==========================================================================
Private Sub ВывестиОтчёт(ws As Worksheet, ByRef lbl() As Variant, ByRef vals() As Variant)

    Dim rHdr As Long, rNum As Long, rData As Long, rLast As Long
    Dim cLbl As Long, cData As Long, cLast As Long
    Dim i As Long, j As Long, k As Long, ri As Long, ci As Long
    Dim v As Variant, a() As Variant, rng As Range

    rHdr = 5
    rNum = rHdr + nLev
    rData = rNum + 1
    rLast = rData + nOrd - 1

    cLbl = 2
    cData = cLbl + nLbl
    cLast = cData + nCOrd - 1

    ws.Cells.Font.Name = FNT

    ' ---------- шапка листа ----------
    With ws.Cells(1, cLbl)
        .Value = TXT_TITLE
        .Font.Size = FS_TITLE: .Font.Bold = True: .Font.Color = cTitle
    End With
    With ws.Cells(2, cLbl)
        .Value = TXT_SUB
        .Font.Size = 10: .Font.Bold = True: .Font.Color = cBordo
    End With
    With ws.Cells(3, cLbl)               ' фильтры сводной
        .Value = ТекстЯчейки(mФильтры)
        .Font.Size = 9: .Font.Color = cBordo
    End With

    ' ---------- шапка таблицы ----------
    ReDim a(1 To nLev, 1 To nCOrd)
    For k = 1 To nLev
        For ci = 1 To nCOrd
            a(k, ci) = ТекстЯчейки(dLev(cOrd(ci), k))
        Next ci
    Next k
    ws.Range(ws.Cells(rHdr, cData), ws.Cells(rNum - 1, cLast)).Value = a
    For k = 1 To nLbl
        ws.Cells(rNum - 1, cLbl + k - 1).Value = ТекстЯчейки(mLblCap(k))
    Next k
    ws.Cells(rNum - 1, 1).Value = "№"

    With ws.Range(ws.Cells(rHdr, 1), ws.Cells(rNum - 1, cLast))
        .Interior.Color = cHead
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = FS_BODY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range(ws.Cells(rHdr, cLbl), ws.Cells(rNum - 1, cLbl + nLbl - 1)).HorizontalAlignment = xlLeft

    ' подписи слева живут на всю высоту шапки
    If nLev > 1 Then
        Dim предупр As Boolean
        предупр = Application.DisplayAlerts
        Application.DisplayAlerts = False
        ws.Range(ws.Cells(rHdr, 1), ws.Cells(rNum - 1, 1)).Merge
        For k = 1 To nLbl
            ws.Range(ws.Cells(rHdr, cLbl + k - 1), ws.Cells(rNum - 1, cLbl + k - 1)).Merge
        Next k
        Application.DisplayAlerts = предупр
    End If

    ОбъединитьШапку ws, rHdr, cData

    ' ---------- нумерация столбцов ----------
    ReDim a(1 To 1, 1 To nLbl + nCOrd)
    For j = 1 To nLbl + nCOrd
        a(1, j) = j
    Next j
    With ws.Range(ws.Cells(rNum, cLbl), ws.Cells(rNum, cLast))
        .Value = a
        .Font.Size = FS_NUM: .Font.Color = cNum
        .HorizontalAlignment = xlCenter
    End With

    ' ---------- подписи строк — одним массивом ----------
    ReDim a(1 To nOrd, 1 To nLbl)
    For ri = 1 To nOrd
        For k = 1 To nLbl
            If mOutR(ri, k) > 0 Then
                a(ri, k) = ЗначениеПодписи(lbl(mOutR(ri, k), mOutC(ri, k)))
            Else
                a(ri, k) = ТекстЯчейки(mOutLbl(ri, k))
            End If
        Next k
    Next ri
    With ws.Range(ws.Cells(rData, cLbl), ws.Cells(rLast, cLbl + nLbl - 1))
        .Value = a
        .HorizontalAlignment = xlLeft
        .Font.Size = FS_BODY
    End With
    ' отступы — сплошными кусками одного отступа, не по ячейке
    Dim нач As Long, отс As Long
    For k = 1 To nLbl
        нач = 1
        For ri = 2 To nOrd + 1
            If ri > nOrd Then отс = -99 Else отс = mOutInd(ri, k)
            If отс <> mOutInd(нач, k) Then
                If mOutInd(нач, k) > 0 Then
                    ws.Range(ws.Cells(rData + нач - 1, cLbl + k - 1), _
                             ws.Cells(rData + ri - 2, cLbl + k - 1)).IndentLevel = Мин(mOutInd(нач, k), 15)
                End If
                нач = ri
            End If
        Next ri
    Next k

    ' ---------- значения — одним массивом ----------
    ReDim a(1 To nOrd, 1 To nCOrd)
    For ri = 1 To nOrd
        i = mOrd(ri)
        For ci = 1 To nCOrd
            v = vals(i, cOrd(ci))
            If ПустоеЗначение(v) Then a(ri, ci) = "-" Else a(ri, ci) = v
        Next ci
    Next ri
    Set rng = ws.Range(ws.Cells(rData, cData), ws.Cells(rLast, cLast))
    rng.Value = a
    rng.HorizontalAlignment = xlCenter
    rng.Font.Size = FS_BODY
    ФорматыЧисел ws, rData, cData, vals

    ' прочерки вместо пустых — серым
    ' (SpecialCells на одной ячейке берёт весь лист — её проверяем сами)
    If rng.Cells.Count = 1 Then
        If VarType(rng.Value) = vbString Then rng.Font.Color = cNum
    Else
        On Error Resume Next
        rng.SpecialCells(xlCellTypeConstants, xlTextValues).Font.Color = cNum
        On Error GoTo 0
    End If

    ' нумерация строк
    ReDim a(1 To nOrd, 1 To 1)
    For ri = 1 To nOrd
        a(ri, 1) = ri
    Next ri
    With ws.Range(ws.Cells(rData, 1), ws.Cells(rLast, 1))
        .Value = a
        .Font.Size = FS_NUM: .Font.Color = cNum
        .HorizontalAlignment = xlCenter
    End With

    ' ---------- подсветка итогов ----------
    ' Красим только итоги: детализация остаётся белой, её держат серые границы.
    ReDim mКлассСтр(1 To nOrd)
    ReDim mКлассКол(1 To nCOrd)
    For ci = 1 To nCOrd
        j = cOrd(ci)
        If cType(j) = LT_GRAND Then
            mКлассКол(ci) = "G"
        ElseIf cType(j) = LT_SUBTOTAL Then
            mКлассКол(ci) = "S"
        Else
            mКлассКол(ci) = ""
        End If
    Next ci
    For ri = 1 To nOrd
        i = mOrd(ri)
        If mType(i) = LT_GRAND Then
            mКлассСтр(ri) = "G"
        ElseIf mРодитель(i) Then
            mКлассСтр(ri) = "S" & mDepth(i)
        Else
            mКлассСтр(ri) = ""
        End If
    Next ri

    ЗакраситьТело ws, rData, rLast, cLbl, cData, cLast

    ' ---------- границы ----------
    Границы ws.Range(ws.Cells(rHdr, 1), ws.Cells(rLast, cLast))

    ' ---------- геометрия ----------
    ws.Columns(1).ColumnWidth = 4
    ws.Range(ws.Cells(rHdr, cLbl), ws.Cells(rLast, cLast)).Columns.AutoFit
    For j = cLbl To cLast
        ' небольшой запас: жирные итоги шире обычного текста, иначе получим ###
        ws.Columns(j).ColumnWidth = ws.Columns(j).ColumnWidth + 1.2
        If ws.Columns(j).ColumnWidth < 7 Then ws.Columns(j).ColumnWidth = 7
        If ws.Columns(j).ColumnWidth > 45 Then ws.Columns(j).ColumnWidth = 45
    Next j
    ws.Rows(rNum).RowHeight = 10
    ws.Rows(4).RowHeight = 6

    ' ---------- группировка строк по уровням ----------
    If ГРУППИРОВАТЬ Then Сгруппировать ws, rData

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Cells(rData, cData).Select
    ActiveWindow.FreezePanes = True

    ПеренестиУсловноеФорматирование ws, rData, cData

    НастроитьПечать ws, rHdr, rNum, rLast, cLast

End Sub

' Строка, похожая на число или дату, но бывшая в сводной текстом,
' должна остаться текстом — иначе «1-2» превратится в 1 февраля.
Private Function ТекстЯчейки(ByVal s As String) As Variant
    If Len(s) > 0 Then
        If IsNumeric(s) Or IsDate(s) Or Left$(s, 1) = "=" Then
            ТекстЯчейки = "'" & s
            Exit Function
        End If
    End If
    ТекстЯчейки = s
End Function

Private Function ЗначениеПодписи(v As Variant) As Variant
    If VarType(v) = vbString Then
        ЗначениеПодписи = ТекстЯчейки(CStr(v))
    ElseIf IsError(v) Then
        ЗначениеПодписи = ""
    Else
        ЗначениеПодписи = v
    End If
End Function

' Формат каждой колонки берём у меры в сводной; «Общий» заменяем
' на разделители разрядов, дроби — только если они есть в данных.
Private Sub ФорматыЧисел(ws As Worksheet, ByVal rData As Long, ByVal cData As Long, _
                         ByRef vals() As Variant)
    Dim ci As Long, ri As Long, j As Long, i As Long, f As String
    Dim поСтрокам As Boolean
    Dim rLast As Long
    rLast = rData + nOrd - 1

    For j = 1 To nC
        If Len(mФорматКол(j)) = 0 Then поСтрокам = True
    Next j

    If Not поСтрокам Then
        For ci = 1 To nCOrd
            j = cOrd(ci)
            f = mФорматКол(j)
            If ОбщийФормат(f) Then f = АвтоФормат(vals, 0, j)
            ws.Range(ws.Cells(rData, cData + ci - 1), ws.Cells(rLast, cData + ci - 1)).NumberFormat = f
        Next ci
    Else
        For ri = 1 To nOrd
            i = mOrd(ri)
            f = mФорматСтр(i)
            If ОбщийФормат(f) Then f = АвтоФормат(vals, i, 0)
            ws.Range(ws.Cells(rData + ri - 1, cData), ws.Cells(rData + ri - 1, cData + nCOrd - 1)).NumberFormat = f
        Next ri
    End If
End Sub

Private Function ОбщийФормат(ByVal f As String) As Boolean
    f = LCase$(Trim$(f))
    ОбщийФормат = (f = "" Or f = "general" Or f = "общий" Or f = "основной" Or f = "@")
End Function

' i = 0 — по колонке j, j = 0 — по строке i
Private Function АвтоФормат(ByRef vals() As Variant, ByVal i As Long, ByVal j As Long) As String
    Dim a As Long, v As Variant, дроби As Boolean
    If i = 0 Then
        For a = 1 To nR
            v = vals(a, j)
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) <> Int(CDbl(v)) Then дроби = True: Exit For
            End If
        Next a
    Else
        For a = 1 To nC
            v = vals(i, a)
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) <> Int(CDbl(v)) Then дроби = True: Exit For
            End If
        Next a
    End If
    If дроби Then
        АвтоФормат = "#,##0.00;[Red]-#,##0.00;""-"""
    Else
        АвтоФормат = "#,##0;[Red]-#,##0;""-"""
    End If
End Function

' Уровни строк -> структура листа: кнопки 1-2-3... сворачивают отчёт
' до нужной глубины. Итог над группой, поэтому SummaryRow = над.
Private Sub Сгруппировать(ws As Worksheet, ByVal rData As Long)
    Dim ri As Long, нач As Long, лв As Long
    On Error Resume Next
    ws.Outline.SummaryRow = xlAbove
    нач = 1
    For ri = 2 To nOrd + 1
        If ri > nOrd Then
            лв = -1
        Else
            лв = mOutLvl(ri)
        End If
        If лв <> mOutLvl(нач) Then
            If mOutLvl(нач) > 1 Then
                ws.Range(ws.Rows(rData + нач - 1), ws.Rows(rData + ri - 2)).OutlineLevel = mOutLvl(нач)
            End If
            нач = ri
        End If
    Next ri
    Err.Clear
    On Error GoTo 0
End Sub


'==========================================================================
'  ЗАКРАСКА, ГРАНИЦЫ, УСЛОВНОЕ ФОРМАТИРОВАНИЕ
'==========================================================================
' Условное форматирование сводной (цветовые шкалы, гистограммы, наборы
' значков, правила по значению) переносится на соответствующие ячейки
' отчёта с учётом того, что строки и колонки переставлены.
Private Sub ПеренестиУсловноеФорматирование(ws As Worksheet, _
                                            ByVal rData As Long, ByVal cData As Long)

    Dim fc As Object, обл As Range, целевой As Range, кус As Range
    Dim строкаЕсть() As Boolean, столбецЕсть() As Boolean
    Dim k As Long, i As Long, j As Long, ri As Long, ci As Long
    Dim нач As Long, попал As Boolean, всего As Long

    If mИсходник Is Nothing Then Exit Sub
    On Error Resume Next
    всего = mИсходник.FormatConditions.Count
    On Error GoTo 0
    If всего = 0 Then Exit Sub

    ReDim строкаЕсть(1 To nR)
    ReDim столбецЕсть(1 To nC)

    For k = 1 To всего

        Set fc = Nothing
        Set обл = Nothing
        On Error Resume Next
        Set fc = mИсходник.FormatConditions(k)
        Set обл = Application.Intersect(fc.AppliesTo, mИсходник)
        On Error GoTo 0
        If обл Is Nothing Then GoTo Дальше

        For i = 1 To nR
            строкаЕсть(i) = Not (Application.Intersect(обл, mИсходник.Rows(i)) Is Nothing)
        Next i
        For j = 1 To nC
            столбецЕсть(j) = Not (Application.Intersect(обл, mИсходник.Columns(j)) Is Nothing)
        Next j

        ' собираем адресат сплошными кусками — по ячейке было бы слишком долго
        Set целевой = Nothing
        For ci = 1 To nCOrd
            If столбецЕсть(cOrd(ci)) Then
                нач = 0
                For ri = 1 To nOrd + 1
                    If ri <= nOrd Then попал = строкаЕсть(mOrd(ri)) Else попал = False
                    If попал And нач = 0 Then нач = ri
                    If (Not попал) And нач > 0 Then
                        Set кус = ws.Range(ws.Cells(rData + нач - 1, cData + ci - 1), _
                                           ws.Cells(rData + ri - 2, cData + ci - 1))
                        If целевой Is Nothing Then
                            Set целевой = кус
                        Else
                            Set целевой = Application.Union(целевой, кус)
                        End If
                        нач = 0
                    End If
                Next ri
            End If
        Next ci

        If Not целевой Is Nothing Then СоздатьУсловие fc, целевой
Дальше:
    Next k

End Sub

Private Sub СоздатьУсловие(fc As Object, rng As Range)

    Dim n As Object, i As Long

    On Error Resume Next
    Select Case fc.Type

        Case 3      ' цветовая шкала
            Set n = rng.FormatConditions.AddColorScale(fc.ColorScaleCriteria.Count)
            For i = 1 To fc.ColorScaleCriteria.Count
                n.ColorScaleCriteria(i).Type = fc.ColorScaleCriteria(i).Type
                n.ColorScaleCriteria(i).Value = fc.ColorScaleCriteria(i).Value
                n.ColorScaleCriteria(i).FormatColor.Color = fc.ColorScaleCriteria(i).FormatColor.Color
            Next i

        Case 4      ' гистограмма
            Set n = rng.FormatConditions.AddDatabar
            n.BarColor.Color = fc.BarColor.Color
            n.BarFillType = fc.BarFillType
            n.Direction = fc.Direction
            n.ShowValue = fc.ShowValue
            n.MinPoint.Modify newtype:=fc.MinPoint.Type, newvalue:=fc.MinPoint.Value
            n.MaxPoint.Modify newtype:=fc.MaxPoint.Type, newvalue:=fc.MaxPoint.Value

        Case 6      ' набор значков
            Set n = rng.FormatConditions.AddIconSetCondition
            n.IconSet = fc.IconSet
            n.ReverseOrder = fc.ReverseOrder
            n.ShowIconOnly = fc.ShowIconOnly
            For i = 2 To fc.IconCriteria.Count
                n.IconCriteria(i).Type = fc.IconCriteria(i).Type
                n.IconCriteria(i).Value = fc.IconCriteria(i).Value
                n.IconCriteria(i).Operator = fc.IconCriteria(i).Operator
            Next i

        Case 1      ' правило по значению ячейки
            If fc.Operator = xlBetween Or fc.Operator = xlNotBetween Then
                Set n = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=fc.Operator, _
                                                 Formula1:=fc.Formula1, Formula2:=fc.Formula2)
            Else
                Set n = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=fc.Operator, _
                                                 Formula1:=fc.Formula1)
            End If
            КопироватьФорматУсловия fc, n

        Case 2      ' правило-формула
            Set n = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fc.Formula1)
            КопироватьФорматУсловия fc, n

        Case Else
            ' остальные типы переносить нечем — пропускаем молча

    End Select
    Err.Clear
    On Error GoTo 0

End Sub

Private Sub КопироватьФорматУсловия(fc As Object, n As Object)
    On Error Resume Next
    n.Interior.Pattern = fc.Interior.Pattern
    n.Interior.Color = fc.Interior.Color
    n.Font.Color = fc.Font.Color
    n.Font.Bold = fc.Font.Bold
    n.Font.Italic = fc.Font.Italic
    n.Borders.LineStyle = fc.Borders.LineStyle
    n.Borders.Color = fc.Borders.Color
    Err.Clear
    On Error GoTo 0
End Sub

' Строка итога — целиком; колонки итогов — сплошными кусками одного цвета.
Private Sub ЗакраситьТело(ws As Worksheet, ByVal rData As Long, ByVal rLast As Long, _
                          ByVal cLbl As Long, ByVal cData As Long, ByVal cLast As Long)

    Dim ri As Long, ci As Long, нач As Long, n As Long
    Dim фонСтр() As Double, фон() As Double

    n = rLast - rData + 1
    ReDim фонСтр(1 To n)
    ReDim фон(1 To n + 1)

    For ri = 1 To n
        фонСтр(ri) = ФонСтроки(mКлассСтр(ri))
        If фонСтр(ri) >= 0 Then
            With ws.Range(ws.Cells(rData + ri - 1, cLbl), ws.Cells(rData + ri - 1, cLast))
                .Interior.Color = Оттенок(cHead, фонСтр(ri))
                .Font.Bold = True
            End With
        End If
    Next ri

    For ci = 1 To UBound(mКлассКол)
        If Len(mКлассКол(ci)) > 0 Then
            For ri = 1 To n
                фон(ri) = ФонКолонки(mКлассКол(ci), фонСтр(ri))
            Next ri
            фон(n + 1) = -2
            нач = 1
            For ri = 2 To n + 1
                If фон(ri) <> фон(нач) Then
                    If фон(нач) >= 0 Then
                        With ws.Range(ws.Cells(rData + нач - 1, cData + ci - 1), _
                                      ws.Cells(rData + ri - 2, cData + ci - 1))
                            .Interior.Color = Оттенок(cHead, фон(нач))
                            .Font.Bold = True
                        End With
                    End If
                    нач = ri
                End If
            Next ri
        End If
    Next ci

End Sub

' Объединение одинаковых заголовков по горизонтали и «висящих» — по вертикали
Private Sub ОбъединитьШапку(ws As Worksheet, ByVal rHdr As Long, ByVal cData As Long)

    Dim k As Long, ci As Long, a As Long, п As Long
    Dim былиПредупреждения As Boolean
    былиПредупреждения = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ' по горизонтали — соседние столбцы с одинаковым текстом на всех уровнях до k
    For k = 1 To nLev
        ci = 1
        Do While ci <= nCOrd
            a = ci
            Do While ci + 1 <= nCOrd
                If Not ОдинаковыДо(cOrd(a), cOrd(ci + 1), k) Then Exit Do
                If Len(Trim$(dLev(cOrd(a), k))) = 0 Then Exit Do
                ci = ci + 1
            Loop
            If ci > a Then
                ws.Range(ws.Cells(rHdr + k - 1, cData + a - 1), _
                         ws.Cells(rHdr + k - 1, cData + ci - 1)).Merge
            End If
            ci = ci + 1
        Loop
    Next k

    ' по вертикали — у столбцов, чьи нижние уровни пустые (итоги):
    ' подпись растягиваем от последнего непустого уровня до низа шапки
    For ci = 1 To nCOrd
        k = 1
        For п = nLev To 1 Step -1
            If Len(Trim$(dLev(cOrd(ci), п))) > 0 Then k = п: Exit For
        Next п
        If k < nLev Then
            On Error Resume Next
            ws.Range(ws.Cells(rHdr + k - 1, cData + ci - 1), _
                     ws.Cells(rHdr + nLev - 1, cData + ci - 1)).Merge
            On Error GoTo 0
        End If
    Next ci

    Application.DisplayAlerts = былиПредупреждения
End Sub

Private Function ОдинаковыДо(ByVal a As Long, ByVal b As Long, ByVal k As Long) As Boolean
    Dim i As Long
    For i = 1 To k
        If StrComp(dLev(a, i), dLev(b, i), vbTextCompare) <> 0 Then Exit Function
    Next i
    ОдинаковыДо = True
End Function

Private Sub Границы(rng As Range)
    Dim i As Variant
    For Each i In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
                        xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(i)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = cBrd
        End With
    Next i
    For Each i In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        rng.Borders(i).Color = cBrdOut
    Next i
End Sub

Private Sub НастроитьПечать(ws As Worksheet, ByVal rHdr As Long, ByVal rNum As Long, _
                            ByVal rLast As Long, ByVal cLast As Long)

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$1:" & ws.Cells(rLast, cLast).Address(True, True)
        .Orientation = xlLandscape
        .LeftMargin = Application.CentimetersToPoints(MARGIN_CM)
        .RightMargin = Application.CentimetersToPoints(MARGIN_CM)
        .TopMargin = Application.CentimetersToPoints(MARGIN_CM)
        .BottomMargin = Application.CentimetersToPoints(MARGIN_CM)
        .HeaderMargin = 0
        .FooterMargin = 0
        .PrintTitleRows = "$" & rHdr & ":$" & rNum
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True

    ' А3 ставим ОТДЕЛЬНО: если драйвер принтера его не поддерживает,
    ' присвоение падает и в отложенном режиме обнуляет всю пачку настроек.
    Err.Clear
    ws.PageSetup.PaperSize = xlPaperA3
    Err.Clear
    On Error GoTo 0

End Sub


'==========================================================================
'  МЕЛОЧИ
'==========================================================================
Private Function ЧитатьДиапазон(rng As Range, ByVal nrow As Long, ByVal ncol As Long) As Variant
    Dim a() As Variant, i As Long, j As Long, v As Variant
    ReDim a(1 To nrow, 1 To ncol)
    If nrow = 1 And ncol = 1 Then
        a(1, 1) = rng.Cells(1, 1).Value
    Else
        v = rng.Value
        For i = 1 To nrow
            For j = 1 To ncol
                a(i, j) = v(i, j)
            Next j
        Next i
    End If
    ЧитатьДиапазон = a
End Function

Private Sub Сообщить(ByVal текст As String)
    If Not mQuiet Then MsgBox текст, vbExclamation, "Отчёт из сводной"
End Sub

Private Function ВСтроку(v As Variant) As String
    On Error Resume Next
    If IsError(v) Then ВСтроку = "" Else ВСтроку = CStr(v)
    On Error GoTo 0
End Function
