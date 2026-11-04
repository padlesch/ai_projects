Attribute VB_Name = "modќтчет»з—воднойV3"
Option Explicit

'==========================================================================
'  ќ‘ќ–ћЋ≈Ќ»≈ ќ“„®“ј »« —¬ќƒЌќ… “јЅЋ»÷џ  Ч верси€ 3
'--------------------------------------------------------------------------
'  ќформитьќтчЄт        Ч встать в любую €чейку сводной и запустить
'  ќформитьќтчЄтЅез¬опросов Ч то же молча, отчЄт новым листом этой книги
'  ѕалитраќтчЄта         Ч сменить палитру готового отчЄта (активный лист)
'
'  „то делает:
'   * снимает со сводной значени€ и структуру Ч в любом макете:
'     сжатом, структурном, табличном и смешанном, с итогами вверху,
'     внизу или без них, до 32 уровней строк;
'   * общий итог по строкам -> первой строкой данных;
'   * итог уровн€ всегда над своими детьми (итоги Ђвнизуї поднимаютс€,
'     пуста€ строка-заголовок группы сливаетс€ со своим итогом);
'   * общий итог и промежуточные итоги по столбцам -> влево от своей группы;
'   * несколько мер -> кажда€ мера отдельным блоком столбцов;
'   * числовые форматы мер (%, валюта, дроби) берутс€ из сводной;
'   * строка фильтров сводной под заголовком;
'   * заголовок, нумераци€ столбцов и строк, Arial Narrow, серые границы.
'
'  Ќа листе после построени€ нет ничего, кроме отчЄта: ни панелей,
'  ни служебных €чеек. ѕалитра мен€етс€ отдельным макросом.
'
'  –аботает и с обычными сводными, и со сводными на модели данных (OLAP).
'==========================================================================

'#ƒ≈ Ћ-Ќј„јЋќ
'--- переключатели поведени€ ---------------------------------------------
Private Const FNT           As String = "Arial Narrow"
Private Const FS_BODY       As Long = 9         ' кегль таблицы
Private Const FS_TITLE      As Long = 16        ' кегль заголовка
Private Const FS_NUM        As Long = 7         ' кегль нумерации
Private Const TXT_TITLE     As String = "јнализ данных"
Private Const TXT_SUB       As String = "ƒ«ћ, «Ћ"
Private Const TXT_GRAND     As String = "ќбщий итог"
Private Const TXT_SUBTOT    As String = "»тог"
Private Const TXT_EMPTY     As String = "(пусто)"
Private Const MARGIN_CM     As Double = 0.3
Private Const ѕјЋ_ѕќ_”ћќЋ„  As Long = 1         ' палитра нового отчЄта
Private Const √–”ѕѕ»–ќ¬ј“№  As Boolean = False  ' True Ч кнопки свЄртки уровней слева

'--- типы строк/столбцов сводной (как XlPivotLineType) --------------------
Private Const LT_REGULAR    As Long = 0
Private Const LT_SUBTOTAL   As Long = 1
Private Const LT_GRAND      As Long = 2
Private Const LT_BLANK      As Long = 3

'--- типы €чеек сводной (как XlPivotCellType) ------------------------------
Private Const CT_ITEM       As Long = 1
Private Const CT_SUBTOTAL   As Long = 2
Private Const CT_GRAND      As Long = 3
Private Const CT_DATAFIELD  As Long = 4
Private Const CT_CUSTOMSUB  As Long = 7
Private Const CT_BLANK      As Long = 9

'--- палитра --------------------------------------------------------------
Private cTitle As Long, cBordo As Long, cHead As Long, cNum As Long
Private cBrd As Long, cBrdOut As Long, cHeadFont As Long
Private p“он As Boolean           ' палитра считаетс€ от цвета шапки
Private p»тог As Long, p”р(1 To 4) As Long, p люч As Long, pѕром ол As Long
Private mѕал As Long
Private Const C_ѕ–ќ„≈–  As Long = 10921638   ' RGB(166,166,166) Ч Ђ-ї вместо пусто
Private Const C_“≈ —“ As Long = 2039583       ' RGB(31,31,31) Ч основной текст

'--- вход €дра строк: то, что сн€то с листа ---------------------------------
Private nR As Long, nLbl As Long
Private inTxt() As String         ' (строка, столбец подписей) текст
Private inCT() As Long            ' тип €чейки сводной, -1 Ч неизвестен
Private inLev() As Long           ' номер пол€ строк (1..), 0 Ч неизвестен
Private inInd() As Long           ' отступ в €чейке
Private inLT() As Long            ' тип линии из оси сводной, -1 Ч неизвестен
Private inHasVal() As Boolean     ' в строке есть хоть одно значение

'--- результат €дра строк --------------------------------------------------
Private mType() As Long
Private mDepth() As Long          ' глубина строки (1 = верхний уровень)
Private mKey() As String          ' (строка, уровень) Ч подпись предка
Private mKeyR() As Long           ' откуда вз€та подпись: строка источника
Private mKeyC() As Long           '                       столбец источника
Private mPath() As String         ' (строка, уровень) Ч полный путь, дл€ сравнени€
Private mOwnCustom() As Boolean   ' пользовательский итог Ч сво€ подпись как есть
Private mDrop() As Boolean        ' строка поглощена своим итогом
Private mMaxD As Long
Private mColOf() As Long          ' в каком столбце подписей живЄт уровень
Private mIndOf() As Long          ' отступ уровн€ внутри столбца
Private mCellLev() As Long
Private mOrd() As Long
Private nOrd As Long
Private m–одитель() As Boolean    ' у строки есть подчинЄнные строки
Private m≈сть–одитель() As Boolean ' (строка, уровень) Ч есть отдельна€ строка-предок
Private mOutLbl() As String       ' (позици€, столбец) Ч подпись в отчЄте
Private mOutR() As Long           ' (позици€, столбец) Ч исходна€ €чейка подписи
Private mOutC() As Long
Private mOutInd() As Long         ' (позици€, столбец) Ч отступ, -1 не задан
Private mOutLvl() As Long         ' (позици€) Ч уровень группировки строк

'--- столбцы --------------------------------------------------------------
Private nC As Long, nLev As Long
Private cItem() As String         ' (столбец, уровень) элементы оси столбцов
Private cCnt() As Long            ' сколько уровней задано у столбца
Private cDF() As String           ' им€ меры столбца, если известно
Private cMeasLev As Long          ' на каком уровне оси сто€т меры (0 Ч нет)
Private mMeas() As String         ' имена мер по пор€дку
Private nMeas As Long
Private dLev() As String          ' dLev(столбец, уровень) Ч текст шапки
Private dDepth() As Long          ' сколько уровней у столбца задано
Private cType() As Long
Private cOrd() As Long
Private nCOrd As Long
Private mGrandTxt As String       ' как общий итог назван в самой сводной
Private mѕоѕол€м As Boolean       ' уровни вз€ты по пол€м сводной, а не по месту
Private m—лот() As Long           ' место (столбец*1000+отступ) каждого уровн€
Private nHR As Long               ' строк шапки столбцов над значени€ми
Private hTxt() As String          ' (строка шапки, столбец) Ч текст на листе
Private hCT() As Long             ' (строка шапки, столбец) Ч тип €чейки сводной
Private inCLT() As Long           ' тип столбца из оси сводной, -1 Ч неизвестен
Private m—лужебные As String      ' |подписи полей|Ђ«начени€ї| Ч не уровни шапки
Private nќбъед As Long            ' план объединени€ шапки
Private mќб 1() As Long, mќб—1() As Long, mќб 2() As Long, mќб—2() As Long
'#ƒ≈ Ћ- ќЌ≈÷

Private mLblCap() As String       ' подписи столбцов с уровн€ми строк
Private mTot() As Double
Private m ласс—тр() As String     ' класс каждой строки отчЄта
Private m ласс ол() As String     ' класс каждой колонки отчЄта
Private m»сходник As Range        ' область значений исходной сводной
Private m‘ормат ол() As String    ' числовой формат по колонкам источника
Private m‘ормат—тр() As String    ' ... или по строкам (меры в строках)
Private m‘ильтры As String
Private mQuiet As Boolean         ' True Ч работать молча (вызов из другого макроса)
Private Const ћ≈“ ј As String = "ќтчЄт_"   ' префикс стилей-меток ролей


'==========================================================================
'  “ќ„ ј ¬’ќƒј
'==========================================================================
' ћолчаливый режим: без вопросов, отчЄт кладЄтс€ новым листом в ту же книгу.
Public Sub ќформитьќтчЄтЅез¬опросов()
    mQuiet = True
    ќформитьќтчЄт
    mQuiet = False
End Sub

Public Sub ќформитьќтчЄт()

    Dim pt As PivotTable
    Dim wsOut As Worksheet
    Dim lbl() As Variant, vals() As Variant
    Dim calc As Long

    On Error GoTo —бой

    Set pt = ѕолучить—водную()
    If pt Is Nothing Then Exit Sub

    InitPalette ѕјЋ_ѕќ_”ћќЋ„

    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If Not –азобрать—водную(pt, lbl, vals) Then GoTo ¬ыход

    ѕор€док—толбцов
    ядро—троки
    ѕор€док—трок
    ядроѕослеѕор€дка

    Application.ScreenUpdating = True
    Set wsOut = —оздатьЋист(pt)
    If wsOut Is Nothing Then GoTo ¬ыход
    Application.ScreenUpdating = False

    ¬ывестиќтчЄт wsOut, lbl, vals

    Application.Calculation = calc
    Application.ScreenUpdating = True
    Set m»сходник = Nothing

    wsOut.Activate
    wsOut.Range("A1").Select
    If Not mQuiet Then
        MsgBox "ќтчЄт сформирован на листе """ & wsOut.Name & """." & vbLf & vbLf & _
               "ѕор€док строк вз€т из сводной; итоги подн€ты над своими группами." & vbLf & _
               "—менить цвет: макрос ѕалитраќтчЄта.", _
               vbInformation, "√отово"
    End If
    Exit Sub

¬ыход:
    Application.Calculation = calc
    Application.ScreenUpdating = True
    Set m»сходник = Nothing
    Exit Sub

—бой:
    Set m»сходник = Nothing
    ' Ѕез этого Excel осталс€ бы с выключенной отрисовкой и пересчЄтом
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.PrintCommunication = True
    If calc <> 0 Then Application.Calculation = calc
    —ообщить "Ќе удалось построить отчЄт." & vbLf & _
             "ќшибка " & Err.Number & ": " & Err.Description

End Sub


'==========================================================================
'  ѕјЋ»“–џ
'  1Ц2 Ч Ђтоновыеї: всЄ считаетс€ от цвета шапки (светлые доли к белому).
'  3Ц4 Ч из дизайна: у каждого уровн€ свой заданный цвет, шапка светла€
'        с тЄмным текстом, свои границы и цвет подзаголовка.
'  „тобы добавить палитру: увеличить „ислоѕалитр, дописать им€ и Case.
'==========================================================================
Private Function „ислоѕалитр() As Long
    „ислоѕалитр = 4
End Function

Private Function »м€ѕалитры(ByVal n As Long) As String
    Select Case n
        Case 2: »м€ѕалитры = "—ерый"
        Case 3: »м€ѕалитры = "A Ј старый формат (персиковые итоги)"
        Case 4: »м€ѕалитры = "B Ј стандартный шаблон (сине-серые итоги)"
        Case Else: »м€ѕалитры = "“Ємно-синий"
    End Select
End Function

Private Sub InitPalette(ByVal n As Long)
    If n < 1 Or n > „ислоѕалитр() Then n = 1
    mѕал = n
    p“он = False
    cHeadFont = RGB(255, 255, 255)
    cBordo = RGB(139, 46, 60)        ' подзаголовок / фильтры
    cNum = RGB(166, 166, 166)        ' нумераци€
    cBrd = RGB(191, 191, 191)        ' внутренние границы
    cBrdOut = RGB(140, 140, 140)     ' внешн€€ рамка
    Select Case n
        Case 2          ' серый
            p“он = True
            cTitle = RGB(74, 79, 85): cHead = RGB(112, 119, 126)
        Case 3          ' A Ј старый формат
            cTitle = RGB(31, 56, 100)       ' #1F3864 тЄмно-синий
            cBordo = RGB(154, 59, 44)       ' #9A3B2C кирпичный
            cHead = RGB(242, 242, 242)      ' #F2F2F2 шапка таблицы
            cHeadFont = RGB(31, 31, 31)     ' #1F1F1F
            cNum = RGB(107, 107, 107)       ' #6B6B6B служебный
            cBrdOut = RGB(166, 166, 166)    ' #A6A6A6
            cBrd = RGB(208, 208, 208)       ' #D0D0D0
            p»тог = RGB(247, 199, 172)      ' #F7C7AC общий итог
            p”р(1) = RGB(251, 226, 213)     ' #FBE2D5
            p”р(2) = RGB(255, 235, 205)     ' #FFEBCD
            p”р(3) = RGB(233, 243, 227)     ' #E9F3E3
            p люч = RGB(237, 237, 237)      ' #EDEDED ключева€ колонка
            pѕром ол = RGB(221, 235, 247)   ' #DDEBF7 промежуточные итоги
        Case 4          ' B Ј стандартный шаблон
            cTitle = RGB(27, 58, 92)        ' #1B3A5C глубокий синий
            cBordo = RGB(45, 106, 115)      ' #2D6A73 морска€ волна
            cHead = RGB(232, 236, 240)      ' #E8ECF0 шапка таблицы
            cHeadFont = RGB(31, 31, 31)     ' #1F1F1F
            cNum = RGB(102, 112, 122)       ' #66707A служебный
            cBrdOut = RGB(154, 165, 177)    ' #9AA5B1
            cBrd = RGB(211, 217, 223)       ' #D3D9DF
            p»тог = RGB(201, 220, 236)      ' #C9DCEC общий итог
            p”р(1) = RGB(220, 232, 243)     ' #DCE8F3
            p”р(2) = RGB(234, 241, 248)     ' #EAF1F8
            p”р(3) = RGB(238, 244, 234)     ' #EEF4EA
            p люч = RGB(236, 239, 242)      ' #ECEFF2 ключева€ колонка
            pѕром ол = RGB(214, 236, 236)   ' #D6ECEC бирюзовый светлый
        Case Else       ' тЄмно-синий
            p“он = True
            cTitle = RGB(31, 56, 100): cHead = RGB(47, 78, 111)
    End Select
    ' уровни глубже третьего дизайн не задаЄт Ч ещЄ светлее третьего
    p”р(4) = ќттенок(p”р(3), 0.5)
End Sub

' ќсветление к белому: 0 Ч исходный цвет, 1 Ч чистый белый
Private Function ќттенок(ByVal c As Long, ByVal f As Double) As Long
    Dim r As Long, g As Long, b As Long
    If f < 0 Then f = 0
    If f > 1 Then f = 1
    r = c Mod 256
    g = (c \ 256) Mod 256
    b = (c \ 65536) Mod 256
    ќттенок = RGB(r + (255 - r) * f, g + (255 - g) * f, b + (255 - b) * f)
End Function

' ÷вет €чейки по классу строки и колонки; -1 Ч без заливки.
' —трока: "G" Ч общий итог, "S<n>" Ч итог/родитель уровн€ n, "" Ч детализаци€.
'  олонка: "G" Ч общий итог, "S" Ч промежуточный, "" Ч обычна€.
Private Function ÷вет–оли(ByVal стр As String, ByVal кол As String) As Long
    Dim f As Double, g As Double, d As Long
    If p“он Then
        f = ‘он—троки(стр)
        g = ‘он олонки(кол, f)
        If g >= 0 Then
            ÷вет–оли = ќттенок(cHead, g)
        ElseIf f >= 0 Then
            ÷вет–оли = ќттенок(cHead, f)
        Else
            ÷вет–оли = -1
        End If
        Exit Function
    End If
    ' дизайн: строка итога краситс€ целиком, колонки выдел€ютс€ на детализации
    If стр = "G" Then
        ÷вет–оли = p»тог
    ElseIf Left$(стр, 1) = "S" Then
        d = 1
        If Len(стр) > 1 Then d = CLng(Mid$(стр, 2))
        If d > 4 Then d = 4
        If d < 1 Then d = 1
        ÷вет–оли = p”р(d)
    ElseIf кол = "G" Then
        ÷вет–оли = p люч
    ElseIf кол = "S" Then
        ÷вет–оли = pѕром ол
    Else
        ÷вет–оли = -1
    End If
End Function

' Ќасколько светлой должна быть строка данного класса (тоновые палитры)
Private Function ‘он—троки(ByVal класс As String) As Double
    Dim d As Long
    ‘он—троки = -1
    If класс = "G" Then ‘он—троки = 0.74: Exit Function
    If Left$(класс, 1) = "S" Then
        d = 1
        If Len(класс) > 1 Then d = CLng(Mid$(класс, 2))
        ‘он—троки = 0.85 + 0.035 * (d - 1)
        If ‘он—троки > 0.95 Then ‘он—троки = 0.95
    End If
End Function

'  олонка главного итога заметно темнее промежуточных (тоновые палитры)
Private Function ‘он олонки(ByVal класс As String, ByVal фон—тр As Double) As Double
    ‘он олонки = -1
    If класс <> "G" And класс <> "S" Then Exit Function
    If фон—тр < 0 Then
        If класс = "G" Then ‘он олонки = 0.86 Else ‘он олонки = 0.93
    Else
        If класс = "G" Then ‘он олонки = фон—тр - 0.1 Else ‘он олонки = фон—тр - 0.05
    End If
    If ‘он олонки < 0 Then ‘он олонки = 0
End Function


'==========================================================================
'  ћ≈“ » –ќЋ≈…
'   аждой €чейке отчЄта, которую красит палитра, назначаетс€ стиль-метка
'  (ЂќтчЄт_Hї Ч шапка, ЂќтчЄт_RS2ї Ч строка итога 2-го уровн€ и т.п.).
'  —тиль ничего не форматирует (все Include* = False) Ч он только помнит
'  роль €чейки. ћетки переживают сохранение, удаление и вставку строк,
'  и по ним палитра мен€етс€ без служебных €чеек на листе.
'==========================================================================
Private Sub ѕометить(rng As Range, ByVal роль As String)
    Dim wb As Workbook, st As Style
    Set wb = rng.Worksheet.Parent
    On Error Resume Next
    Set st = wb.Styles(ћ≈“ ј & роль)
    On Error GoTo 0
    If st Is Nothing Then
        Set st = wb.Styles.Add(ћ≈“ ј & роль)
        On Error Resume Next
        st.IncludeNumber = False
        st.IncludeFont = False
        st.IncludeAlignment = False
        st.IncludeBorder = False
        st.IncludePatterns = False
        st.IncludeProtection = False
        On Error GoTo 0
    End If
    rng.Style = ћ≈“ ј & роль
End Sub

Private Function –ольячейки(cel As Range) As String
    Dim s As String
    On Error Resume Next
    s = cel.Style.Name
    On Error GoTo 0
    If Left$(s, Len(ћ≈“ ј)) = ћ≈“ ј Then –ольячейки = Mid$(s, Len(ћ≈“ ј) + 1)
End Function

'  ласс строки -> метка: уровни глубже 4-го крас€тс€ как 4-й
Private Function ћетка—троки(ByVal класс As String) As String
    Dim d As Long
    If класс = "G" Then
        ћетка—троки = "RG"
    ElseIf Left$(класс, 1) = "S" Then
        d = CLng(Mid$(класс, 2))
        If d > 4 Then d = 4
        ћетка—троки = "RS" & d
    Else
        ћетка—троки = "RD"
    End If
End Function


'==========================================================================
'  —ћ≈Ќј ѕјЋ»“–џ √ќ“ќ¬ќ√ќ ќ“„®“ј Ч отдельные макросы
'==========================================================================
Public Sub ѕалитраќтчЄта()
    Dim s As String, i As Long, otvet As String, тек As Long
    тек = “екуща€ѕалитра(ActiveSheet)
    If тек = 0 Then
        MsgBox "Ќа активном листе не найден отчЄт, построенный макросом.", vbExclamation
        Exit Sub
    End If
    For i = 1 To „ислоѕалитр()
        s = s & i & " Ч " & »м€ѕалитры(i)
        If i = тек Then s = s & "   (сейчас)"
        s = s & vbLf
    Next i
    otvet = InputBox("Ќомер палитры дл€ отчЄта на активном листе:" & vbLf & vbLf & s, _
                     "ѕалитра отчЄта", CStr(тек Mod „ислоѕалитр() + 1))
    If Len(otvet) = 0 Or Not IsNumeric(otvet) Then Exit Sub
    ѕалитраЌомер CLng(otvet)
End Sub

Private Sub ѕалитраЌомер(ByVal n As Long)
    If n < 1 Or n > „ислоѕалитр() Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Not ѕерекраситьЋист(ActiveSheet, n) Then
        MsgBox "Ќа активном листе не найден отчЄт, построенный макросом.", vbExclamation
    End If
End Sub

' Ќомер палитры, в которой сейчас окрашен отчЄт (по цвету шапки); 0 Ч отчЄта нет
Private Function “екуща€ѕалитра(ws As Object) As Long
    Dim r As Long, p As Long, c As Long
    If TypeName(ws) <> "Worksheet" Then Exit Function
    For r = 1 To 60
        If –ольячейки(ws.Cells(r, 1)) = "H" Then
            c = ws.Cells(r, 1).Interior.Color
            For p = 1 To „ислоѕалитр()
                InitPalette p
                If cHead = c Then “екуща€ѕалитра = p: Exit Function
            Next p
            “екуща€ѕалитра = 1      ' отчЄт есть, цвет правили вручную
            Exit Function
        End If
    Next r
End Function

' √еометри€ и классы Ч из меток: столбец A хранит роли строк,
' строка нумерации Ч роли колонок.
Private Function ѕерекраситьЋист(ws As Worksheet, ByVal n As Long) As Boolean

    Dim r As Long, c As Long, посл—тр As Long, посл ол As Long, роль As String
    Dim rHdr As Long, rHdrEnd As Long, rNum As Long, rData As Long, rLast As Long
    Dim cLbl As Long, cData As Long, cLast As Long

    посл—тр = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    посл ол = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1

    For r = 1 To посл—тр
        роль = –ольячейки(ws.Cells(r, 1))
        If роль = "H" Then
            If rHdr = 0 Then rHdr = r
            rHdrEnd = r
        ElseIf Left$(роль, 1) = "R" Then
            If rData = 0 Then rData = r
            rLast = r
        End If
    Next r
    If rHdr = 0 Or rData = 0 Then Exit Function
    rNum = rHdrEnd + 1

    cLbl = 2
    For c = 2 To посл ол
        роль = –ольячейки(ws.Cells(rNum, c))
        If роль = "C" Or роль = "CG" Or роль = "CS" Then
            If cData = 0 Then cData = c
            cLast = c
        End If
    Next c
    If cData = 0 Then Exit Function

    ReDim m ласс—тр(1 To rLast - rData + 1)
    For r = rData To rLast
        роль = –ольячейки(ws.Cells(r, 1))
        Select Case роль
            Case "RG": m ласс—тр(r - rData + 1) = "G"
            Case "RS1", "RS2", "RS3", "RS4": m ласс—тр(r - rData + 1) = "S" & Right$(роль, 1)
            Case Else: m ласс—тр(r - rData + 1) = ""
        End Select
    Next r
    ReDim m ласс ол(1 To cLast - cData + 1)
    For c = cData To cLast
        m ласс ол(c - cData + 1) = Mid$(–ольячейки(ws.Cells(rNum, c)), 2)
    Next c

    InitPalette n
    Application.ScreenUpdating = False
    ѕокрасить ws, rHdr, rNum, rData, rLast, cLbl, cData, cLast
    Application.ScreenUpdating = True
    ѕерекраситьЋист = True

End Function

' ¬сЄ, что зависит от палитры, Ч в одном месте: и дл€ нового отчЄта,
' и дл€ перекраски готового.
Private Sub ѕокрасить(ws As Worksheet, ByVal rHdr As Long, ByVal rNum As Long, _
                      ByVal rData As Long, ByVal rLast As Long, ByVal cLbl As Long, _
                      ByVal cData As Long, ByVal cLast As Long)
    Dim r As Long
    For r = 1 To rHdr - 1
        Select Case –ольячейки(ws.Cells(r, cLbl))
            Case "T": ws.Cells(r, cLbl).Font.Color = cTitle
            Case "P": ws.Cells(r, cLbl).Font.Color = cBordo
        End Select
    Next r
    With ws.Range(ws.Cells(rHdr, 1), ws.Cells(rNum - 1, cLast))
        .Interior.Color = cHead
        .Font.Color = cHeadFont
    End With
    ws.Range(ws.Cells(rNum, cLbl), ws.Cells(rNum, cLast)).Font.Color = cNum
    ws.Range(ws.Cells(rData, 1), ws.Cells(rLast, 1)).Font.Color = cNum
    «акрасить“ело ws, rData, rLast, cLbl, cData, cLast
    √раницы ws.Range(ws.Cells(rHdr, 1), ws.Cells(rLast, cLast))
End Sub


'==========================================================================
'  ѕќ»—  —¬ќƒЌќ…
'==========================================================================
Private Function ѕолучить—водную() As PivotTable

    Dim ws As Worksheet, i As Long, n As Long
    Dim spisok As String, otvet As String

    If TypeName(ActiveSheet) <> "Worksheet" Then
        —ообщить "јктивный лист не €вл€етс€ рабочим листом."
        Exit Function
    End If
    Set ws = ActiveSheet

    ' 1) курсор внутри сводной
    On Error Resume Next
    Set ѕолучить—водную = ActiveCell.PivotTable
    On Error GoTo 0
    If Not ѕолучить—водную Is Nothing Then Exit Function

    ' 2) единственна€ сводна€ на листе
    n = ws.PivotTables.Count
    If n = 0 Then
        —ообщить "Ќа листе """ & ws.Name & """ нет сводных таблиц." & vbLf & _
                 "ќткройте лист со сводной и запустите макрос ещЄ раз."
        Exit Function
    End If
    If n = 1 Then
        Set ѕолучить—водную = ws.PivotTables(1)
        Exit Function
    End If

    ' 3) несколько Ч спрашиваем
    If mQuiet Then
        Set ѕолучить—водную = ws.PivotTables(1)
        Exit Function
    End If
    For i = 1 To n
        spisok = spisok & i & " Ч " & ws.PivotTables(i).Name & _
                 "   (" & ws.PivotTables(i).TableRange1.Address(False, False) & ")" & vbLf
    Next i
    otvet = InputBox("Ќа листе несколько сводных таблиц. ¬ведите номер нужной:" & _
                     vbLf & vbLf & spisok, "¬ыбор сводной таблицы", "1")
    If Not IsNumeric(otvet) Then Exit Function
    i = CLng(otvet)
    If i < 1 Or i > n Then Exit Function
    Set ѕолучить—водную = ws.PivotTables(i)

End Function


'==========================================================================
'  –ј«Ѕќ– —¬ќƒЌќ… Ч сн€тие данных с листа
'==========================================================================
Private Function –азобрать—водную(pt As PivotTable, _
                                  ByRef lbl() As Variant, _
                                  ByRef vals() As Variant) As Boolean

    Dim rr As Range, db As Range
    Dim i As Long, k As Long, j As Long, смещ—трок As Long

    On Error Resume Next
    Set db = pt.DataBodyRange
    Set rr = pt.RowRange
    On Error GoTo 0

    If db Is Nothing Then
        —ообщить "¬ сводной нет значений: добавьте хот€ бы одно поле в область Ђ«начени€ї."
        Exit Function
    End If

    nR = db.Rows.Count
    nC = db.Columns.Count
    If nR = 0 Or nC = 0 Then Exit Function

    Set m»сходник = db
    vals = „итатьƒиапазон(db, nR, nC)

    mGrandTxt = ""
    On Error Resume Next
    mGrandTxt = pt.GrandTotalName
    Err.Clear
    On Error GoTo 0

    If rr Is Nothing Then
        ' —водна€ без полей в строках: одна строка общего итога
        nLbl = 1
        ReDim lbl(1 To nR, 1 To 1)
        For i = 1 To nR
            lbl(i, 1) = “екстќбщего»тога()
        Next i
        ReDim mLblCap(1 To 1)
        mLblCap(1) = ""
        —н€ть—трокиЅезѕодписей
    Else
        nLbl = rr.Columns.Count
        ' RowRange включает сверху строку подписей полей Ч снимаем сдвиг,
        ' иначе подписи разъедутс€ с числами на одну строку.
        смещ—трок = db.Row - rr.Row
        If смещ—трок < 0 Then смещ—трок = 0
        If смещ—трок + nR > rr.Rows.Count Then смещ—трок = rr.Rows.Count - nR
        If смещ—трок < 0 Then смещ—трок = 0
        Set rr = rr.Offset(смещ—трок, 0).Resize(nR, nLbl)
        lbl = „итатьƒиапазон(rr, nR, nLbl)

        ' подписи столбцов уровней строк Ч в строке пр€мо над данными
        ReDim mLblCap(1 To nLbl)
        For k = 1 To nLbl
            mLblCap(k) = ""
            On Error Resume Next
            mLblCap(k) = ¬—троку(db.Worksheet.Cells(db.Row - 1, rr.Column + k - 1).Value)
            On Error GoTo 0
            If Len(Trim$(mLblCap(k))) = 0 Then mLblCap(k) = "”ровень " & k
        Next k

        —н€ть—троки pt, rr, lbl
    End If

    ReDim inHasVal(1 To nR)
    For i = 1 To nR
        For j = 1 To nC
            If Not ѕустое«начение(vals(i, j)) Then inHasVal(i) = True: Exit For
        Next j
    Next i

    —н€ть—толбцы pt
    —обратьЎапку

    —н€ть‘орматы db
    m‘ильтры = “екст‘ильтров(pt)

    –азобрать—водную = True

End Function

Private Sub —н€ть—трокиЅезѕодписей()
    Dim i As Long
    ReDim inTxt(1 To nR, 1 To 1)
    ReDim inCT(1 To nR, 1 To 1)
    ReDim inLev(1 To nR, 1 To 1)
    ReDim inInd(1 To nR, 1 To 1)
    ReDim inLT(1 To nR)
    For i = 1 To nR
        inTxt(i, 1) = “екстќбщего»тога()
        inCT(i, 1) = CT_GRAND
        inLT(i) = LT_GRAND
    Next i
End Sub

' ƒл€ каждой €чейки подписей: текст, отступ и Ч главное Ч что о ней
' знает сама сводна€: тип €чейки и поле, к которому она относитс€.
' ѕоле однозначно задаЄт уровень в любом макете (сжатом, структурном,
' табличном, смешанном), а отступ и столбец Ч запасной вариант.
Private Sub —н€ть—троки(pt As PivotTable, rr As Range, ByRef lbl() As Variant)

    Dim i As Long, k As Long, n As Long, f As Long
    Dim пол€() As String, им€ƒанных As String, им€ As String
    Dim cel As Range, ct As Long, пусто As Boolean
    Dim ax As Object, ok As Boolean

    ReDim inTxt(1 To nR, 1 To nLbl)
    ReDim inCT(1 To nR, 1 To nLbl)
    ReDim inLev(1 To nR, 1 To nLbl)
    ReDim inInd(1 To nR, 1 To nLbl)
    ReDim inLT(1 To nR)

    n = 0
    On Error Resume Next
    n = pt.RowFields.Count
    If n > 0 Then
        ReDim пол€(1 To n)
        For f = 1 To n
            пол€(f) = pt.RowFields(f).Name
        Next f
    End If
    им€ƒанных = ""
    им€ƒанных = pt.DataPivotField.Name
    Err.Clear
    On Error GoTo 0

    ' типы линий из оси сводной Ч самый надЄжный признак итога
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

    ' —водную опрашиваем по каждой непустой €чейке подписей: уровень
    ' должен быть точным в любом макете, а не угаданным по образцу.
    For i = 1 To nR
        If i Mod 500 = 0 Then
            Application.StatusBar = "ќтчЄт из сводной: разбор строк " & i & " из " & nR
        End If
        пусто = True
        For k = 1 To nLbl
            inTxt(i, k) = ¬—троку(lbl(i, k))
            If Len(Trim$(inTxt(i, k))) > 0 Then пусто = False
        Next k
        For k = 1 To nLbl
            inCT(i, k) = -1
            ' по пустым €чейкам спрашиваем сводную только в строке без подписей
            ' (пустой элемент) Ч иначе на больших отчЄтах это долго
            If Len(Trim$(inTxt(i, k))) > 0 Or пусто Then
                Set cel = rr.Cells(i, k)
                On Error Resume Next
                inInd(i, k) = cel.IndentLevel
                ct = -1
                ct = cel.PivotCell.PivotCellType
                If Err.Number <> 0 Then ct = -1: Err.Clear
                inCT(i, k) = ct
                им€ = ""
                If ct = CT_ITEM Or ct = CT_SUBTOTAL Or ct = CT_CUSTOMSUB Then
                    им€ = cel.PivotCell.PivotField.Name
                    If Err.Number <> 0 Then
                        Err.Clear
                        им€ = cel.PivotCell.PivotItem.Parent.Name
                        If Err.Number <> 0 Then им€ = "": Err.Clear
                    End If
                ElseIf ct = CT_DATAFIELD Then
                    им€ = им€ƒанных
                End If
                On Error GoTo 0
                If Len(им€) > 0 Then
                    For f = 1 To n
                        If StrComp(пол€(f), им€, vbTextCompare) = 0 Then inLev(i, k) = f: Exit For
                    Next f
                End If
            End If
        Next k
    Next i

    Application.StatusBar = False
End Sub

' ѕустое значение сводной: пусто, ошибка или пуста€ строка
Private Function ѕустое«начение(v As Variant) As Boolean
    If IsEmpty(v) Or IsError(v) Then ѕустое«начение = True: Exit Function
    If Not IsNumeric(v) Then
        If Len(Trim$(¬—троку(v))) = 0 Then ѕустое«начение = True
    End If
End Function

' „исловые форматы источника: обычно у каждой меры (колонки) свой,
' а если меры сто€т в строках Ч у каждой строки.
Private Sub —н€ть‘орматы(db As Range)
    Dim j As Long, i As Long, v As Variant, по ол As Boolean
    ReDim m‘ормат ол(1 To nC)
    ReDim m‘ормат—тр(1 To nR)
    по ол = True
    On Error Resume Next
    For j = 1 To nC
        v = Null
        v = db.Columns(j).NumberFormat
        If IsNull(v) Then по ол = False Else m‘ормат ол(j) = CStr(v)
    Next j
    If Not по ол Then
        For j = 1 To nC
            m‘ормат ол(j) = ""
        Next j
        For i = 1 To nR
            v = Null
            v = db.Rows(i).NumberFormat
            If Not IsNull(v) Then m‘ормат—тр(i) = CStr(v)
        Next i
    End If
    On Error GoTo 0
End Sub

' "–егион: —евер;   √од: 2024" Ч фильтры сводной, как они видны на листе:
' подпись фильтра и значение в соседней клетке (ЂAllї, Ђ(несколько
' элементов)ї, выбранный элемент). ‘ильтры Ђвсеї не пишем.
Private Function “екст‘ильтров(pt As PivotTable) As String
    Dim pf As PivotField, s As String, t As String, им€ As String
    Dim стр() As Long, txt() As String, n As Long, i As Long, j As Long, rr As Long, tt As String
    On Error Resume Next
    n = pt.PageFields.Count
    If n = 0 Then Exit Function
    ReDim стр(1 To n)
    ReDim txt(1 To n)
    n = 0
    For Each pf In pt.PageFields
        t = "": им€ = "": rr = 0
        rr = pf.LabelRange.Row
        им€ = Trim$(¬—троку(pf.LabelRange.Value))
        If Len(им€) = 0 Then им€ = „истое»м€(pf.Caption)
        t = Trim$(pf.LabelRange.Offset(0, 1).Text)
        If Len(t) = 0 Then t = Trim$(„истое»м€(pf.CurrentPageName))
        Err.Clear
        Select Case LCase$(t)
            Case "", "all", "(all)", "(все)", "все"
            Case Else
                n = n + 1
                стр(n) = rr
                txt(n) = им€ & ": " & t
        End Select
    Next pf
    On Error GoTo 0
    ' в том пор€дке, в каком фильтры сто€т на листе
    For i = 2 To n
        For j = i To 2 Step -1
            If стр(j) < стр(j - 1) Then
                rr = стр(j): стр(j) = стр(j - 1): стр(j - 1) = rr
                tt = txt(j): txt(j) = txt(j - 1): txt(j - 1) = tt
            End If
        Next j
    Next i
    For i = 1 To n
        If i > 1 Then s = s & ";   "
        s = s & txt(i)
    Next i
    “екст‘ильтров = s
End Function

'--------------------------------------------------------------------------
'  —толбцы: шапка Ч с листа, тип столбца Ч из оси сводной
'--------------------------------------------------------------------------
' ѕодписи шапки берутс€ ровно такими, как они видны в сводной (с учЄтом
' переименований): у сводных на модели данных объектна€ модель отдаЄт
' подписи элементов пустыми или исходными. “ип столбца (обычный / итог /
' промежуточный итог) надЄжно знает ось сводной Ч его берЄм оттуда.
Private Sub —н€ть—толбцы(pt As PivotTable)

    Dim ax As Object, j As Long, r As Long, ok As Boolean, ct As Long, i As Long
    Dim cr As Range, db As Range, ws As Worksheet, верх As Long, cel As Range
    Dim сл As String

    ReDim cType(1 To nC)
    ReDim cCnt(1 To nC)
    ReDim cDF(1 To nC)
    ReDim inCLT(1 To nC)
    cMeasLev = 0

    nMeas = 0
    On Error Resume Next
    nMeas = pt.DataFields.Count
    If nMeas > 0 Then
        ReDim mMeas(1 To nMeas)
        For j = 1 To nMeas
            mMeas(j) = »м€ћеры(pt.DataFields(j))
        Next j
    Else
        ReDim mMeas(1 To 1)
    End If
    Err.Clear
    On Error GoTo 0

    ' тип каждого столбца Ч из оси
    ok = False
    On Error Resume Next
    Set ax = pt.PivotColumnAxis
    If Not ax Is Nothing Then ok = (ax.PivotLines.Count = nC)
    For j = 1 To nC
        inCLT(j) = -1
        If ok Then
            inCLT(j) = ax.PivotLines(j).LineType
            If Err.Number <> 0 Then inCLT(j) = -1: Err.Clear
        End If
    Next j
    Err.Clear
    On Error GoTo 0

    ' строки шапки над областью значений
    Set db = pt.DataBodyRange
    Set ws = db.Worksheet
    On Error Resume Next
    Set cr = pt.ColumnRange
    On Error GoTo 0
    If cr Is Nothing Then верх = db.Row - 1 Else верх = cr.Row
    If верх >= db.Row Then верх = db.Row - 1
    If верх < 1 Then верх = 1
    nHR = db.Row - верх
    If nHR < 1 Then nHR = 1
    ReDim hTxt(1 To nHR, 1 To nC)
    ReDim hCT(1 To nHR, 1 To nC)
    For r = 1 To nHR
        For j = 1 To nC
            hCT(r, j) = -1
            If верх + r - 1 < db.Row Then
                Set cel = ws.Cells(верх + r - 1, db.Column + j - 1)
                hTxt(r, j) = Trim$(¬—троку(cel.Value))
                ct = -1
                On Error Resume Next
                ct = cel.PivotCell.PivotCellType
                If Err.Number <> 0 Then ct = -1: Err.Clear
                On Error GoTo 0
                hCT(r, j) = ct
            End If
        Next j
    Next r

    ' надписи, которые уровнем шапки не €вл€ютс€: названи€ полей и Ђ«начени€ї
    сл = "|названи€ столбцов|column labels|значени€|values|"
    On Error Resume Next
    сл = сл & LCase$(pt.DataPivotField.Caption) & "|"
    For i = 1 To pt.ColumnFields.Count
        сл = сл & LCase$(pt.ColumnFields(i).Caption) & "|"
    Next i
    Err.Clear
    On Error GoTo 0
    m—лужебные = сл

    ЎапкаѕоЋисту

End Sub

' Ќазвание меры, как оно показано в сводной (с учЄтом переименовани€)
Private Function »м€ћеры(df As Object) As String
    Dim s As String
    On Error Resume Next
    s = df.Caption
    If Err.Number <> 0 Or Len(s) = 0 Then Err.Clear: s = df.Name
    Err.Clear
    On Error GoTo 0
    »м€ћеры = „истое»м€(s)
End Function


'==========================================================================
'  ѕ≈„ј“№ ‘»Ћ№“–ќ¬, ћ≈Ћќ„» –ј«Ѕќ–ј
'==========================================================================
' ”никальное им€ OLAP вида [ginek].[ —— ѕ новый].&[ —] или
' [Measures].[„исло элементов] сводим к тому, что в последних скобках.
' ќбычный текст не трогаем: требуем именно форму [..].[..]
Private Function „истое»м€(ByVal s As String) As String
    Dim p As Long
    „истое»м€ = s
    If Left$(s, 1) <> "[" Then Exit Function
    If Right$(s, 1) <> "]" Then Exit Function
    If InStr(s, "].[") = 0 And InStr(s, "].&[") = 0 Then Exit Function
    p = InStrRev(s, "&[")
    If p > 0 Then
        „истое»м€ = Mid$(s, p + 2, Len(s) - p - 2)
        Exit Function
    End If
    p = InStrRev(s, "[")
    If p > 0 Then „истое»м€ = Mid$(s, p + 1, Len(s) - p - 1)
End Function


'#яƒ–ќ-Ќј„јЋќ
'==========================================================================
'  яƒ–ќ: —“–” “”–ј —“–ќ 
'  Ќе обращаетс€ к Excel Ч работает только с массивами in*, поэтому
'  одинаково ведЄт себ€ на любом макете и провер€етс€ автотестами.
'==========================================================================
Private Sub ядро—троки()

    Dim i As Long, k As Long, q As Long, lv As Long, глуб As Long
    Dim посл As Long, s As String, база As String, преж√луб As Long
    Dim стек() As String, стекR() As Long, стекC() As Long

    ReDim mType(1 To nR)
    ReDim mDepth(1 To nR)
    ReDim mDrop(1 To nR)
    ReDim mOwnCustom(1 To nR)

    For i = 1 To nR
        mType(i) = “ип—троки(i)
    Next i

    ѕостроить”ровни

    ReDim mKey(1 To nR, 1 To mMaxD)
    ReDim mKeyR(1 To nR, 1 To mMaxD)
    ReDim mKeyC(1 To nR, 1 To mMaxD)
    ReDim mPath(1 To nR, 1 To mMaxD)
    ReDim стек(1 To mMaxD)
    ReDim стекR(1 To mMaxD)
    ReDim стекC(1 To mMaxD)

    преж√луб = 0
    For i = 1 To nR
        If mType(i) = LT_GRAND Or mType(i) = LT_BLANK Then
            mDepth(i) = 1
            GoTo —лед—трока
        End If

        посл = 0
        For k = nLbl To 1 Step -1
            If mCellLev(i, k) > 0 Then посл = k: Exit For
        Next k

        If посл = 0 Then
            ' —трока без подписи Ч пустой элемент. ”ровень знает сводна€;
            ' если нет Ч это ребЄнок предыдущей строки (в табличном макете
            ' предыдуща€ строка уже на самом глубоком уровне).
            lv = 0
            If mѕоѕол€м Then
                lv = ѕустой”ровень(i)
            ElseIf nLbl = 1 Then
                ' сжатый макет: у пустого элемента тот же отступ, что у соседей
                For q = 1 To mMaxD
                    If m—лот(q) = 1000 + ћин(ћакс(inInd(i, 1), 0), 999) Then lv = q: Exit For
                Next q
            End If
            If lv = 0 Then lv = преж√луб + 1
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
                        ' —трока-итог бывает двух видов:
                        '  * итоги внизу: Ђ—евер »тогї после детей Ч группа
                        '    Ђ—еверї уже в стеке, им€ берЄм из подписи без Ђ»тогї;
                        '  * итоги вверху: строка родител€ Ђ—еверї сама €вл€етс€
                        '    итогом (Excel так еЄ и помечает) Ч подпись и есть им€
                        '    Ќќ¬ќ… группы, старое им€ из стека брать нельз€.
                        ' ѕользовательские итоги (Ђ—евер —уммаї) относ€тс€
                        ' к группе в стеке.
                        база = —уффикс»тога(s)
                        If inCT(i, k) = CT_CUSTOMSUB Then
                            mOwnCustom(i) = True
                            If Len(стек(lv)) = 0 Then стек(lv) = s: стекR(lv) = i: стекC(lv) = k
                        ElseIf Len(база) > 0 And StrComp(база, стек(lv), vbTextCompare) = 0 Then
                            ' итог внизу: группа та же, что в стеке
                        Else
                            стек(lv) = s: стекR(lv) = i: стекC(lv) = k
                        End If
                        For q = lv + 1 To mMaxD
                            стек(q) = "": стекR(q) = 0: стекC(q) = 0
                        Next q
                    ElseIf StrComp(s, стек(lv), vbTextCompare) <> 0 Or k = посл Then
                        ' Ќова€ подпись уровн€ Ч нижние уровни начинаютс€ заново.
                        ' ѕовтор той же подписи предка (табличный макет с
                        ' Ђповтор€ть подписиї) группу не мен€ет и детей не сбрасывает.
                        стек(lv) = s: стекR(lv) = i: стекC(lv) = k
                        For q = lv + 1 To mMaxD
                            стек(q) = "": стекR(q) = 0: стекC(q) = 0
                        Next q
                    End If
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
        преж√луб = глуб
—лед—трока:
    Next i

    ѕоглотить«аголовки

End Sub

' “ип строки: из оси сводной, иначе из типов €чеек, иначе по тексту
Private Function “ип—троки(ByVal i As Long) As Long
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
                    If Ётоќбщий»тог(s) Then
                        t = LT_GRAND
                    ElseIf Len(—уффикс»тога(s)) > 0 Then
                        t = LT_SUBTOTAL
                    End If
                    Exit For
                End If
            Next k
        End If
        If t < 0 Then t = LT_REGULAR
    End If

    ' пуста€ строка-разделитель (Ђпуста€ строка после каждого элементаї)
    If t <> LT_GRAND And Not есть And Not inHasVal(i) Then t = LT_BLANK
    If t = LT_BLANK And есть Then t = LT_REGULAR
    “ип—троки = t
End Function

Private Function Ётоќбщий»тог(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    Ётоќбщий»тог = (s = "общий итог" Or s = "grand total" Or s = LCase$(TXT_GRAND) _
                    Or s = LCase$(Trim$(mGrandTxt)))
End Function

' ѕодпись общего итога: переименованна€ в сводной, иначе стандартна€
Private Function “екстќбщего»тога() As String
    If Len(Trim$(mGrandTxt)) > 0 Then “екстќбщего»тога = mGrandTxt Else “екстќбщего»тога = TXT_GRAND
End Function

' Ђ—евер »тогї -> Ђ—еверї; не итог -> ""
Private Function —уффикс»тога(ByVal s As String) As String
    Dim t As String, x As Variant
    t = Trim$(s)
    For Each x In Array(" итог", " total", " всего", " итого")
        If Len(t) > Len(x) Then
            If StrComp(Right$(t, Len(x)), x, vbTextCompare) = 0 Then
                —уффикс»тога = Trim$(Left$(t, Len(t) - Len(x)))
                Exit Function
            End If
        End If
    Next x
End Function

' ”ровень каждой €чейки подписей.
' √лавный признак Ч место на листе, которое видно глазами: столбец, а внутри
' столбца Ч отступ (сжатый макет держит несколько полей в одном столбце и
' различает их только отступом). “ак уровни определ€ютс€ одинаково в любом
' макете и на модели данных, где сведени€ о пол€х от сводной ненадЄжны.
' —ведени€ о пол€х берутс€ только когда место неоднозначно: в одном столбце
' с одним отступом сто€т разные пол€ (сжатый макет с нулевым отступом).
Private Sub ѕостроить”ровни()

    Dim i As Long, k As Long, j As Long, n As Long, v As Long, a As Long, м As Long
    Dim поѕол€м As Boolean
    Dim знач() As Long, сыр() As Long
    Dim место() As Long, поле() As Long, nм As Long

    ReDim mCellLev(1 To nR, 1 To nLbl)
    ReDim сыр(1 To nR, 1 To nLbl)

    ' есть ли место, где сводна€ называет два разных пол€
    ReDim место(1 To nR * nLbl + 1)
    ReDim поле(1 To nR * nLbl + 1)
    поѕол€м = False
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 And inLev(i, k) > 0 Then
                    м = k * 1000 + ћин(ћакс(inInd(i, k), 0), 999)
                    For a = 1 To nм
                        If место(a) = м Then Exit For
                    Next a
                    If a > nм Then
                        nм = nм + 1: место(nм) = м: поле(nм) = inLev(i, k)
                    ElseIf поле(a) <> inLev(i, k) Then
                        поѕол€м = True
                    End If
                End If
            Next k
        End If
    Next i

    ' по пол€м Ч только если сводна€ знает поле каждой непустой подписи
    If поѕол€м Then
        ƒозаполнить”ровни
        For i = 1 To nR
            If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
                For k = 1 To nLbl
                    If Len(Trim$(inTxt(i, k))) > 0 And inLev(i, k) <= 0 Then поѕол€м = False
                Next k
            End If
        Next i
    End If

    ' сырые координаты и их различные значени€ по возрастанию
    ReDim знач(1 To nR * nLbl + 1)
    n = 0
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 Then
                    If поѕол€м Then
                        v = k * 1000 + ћин(inLev(i, k), 999)
                    Else
                        v = k * 1000 + ћин(ћакс(inInd(i, k), 0), 999)
                    End If
                    сыр(i, k) = v
                    ¬ставить«нач знач, n, v
                End If
            Next k
        End If
    Next i
    If n = 0 Then n = 1: знач(1) = 1000
    mѕоѕол€м = поѕол€м
    ReDim m—лот(1 To n)
    For j = 1 To n
        m—лот(j) = знач(j)
    Next j

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
    ' отступ внутри столбца Ч сколько уровней выше живут в том же столбце
    For j = 1 To mMaxD
        mIndOf(j) = 0
        For k = 1 To j - 1
            If mColOf(k) = mColOf(j) Then mIndOf(j) = mIndOf(j) + 1
        Next k
    Next j

End Sub

' ячейка, о которой сводна€ не сказала поле (бывает у итогов и на модели
' данных), получает уровень по месту на листе Ч если на этом месте
' (столбец + отступ) у остальных €чеек встречаетс€ ровно одно поле.
Private Sub ƒозаполнить”ровни()
    Dim i As Long, k As Long, a As Long, n As Long, м As Long
    Dim место() As Long, ур() As Long, дво€ко() As Boolean
    ReDim место(1 To nR * nLbl + 1)
    ReDim ур(1 To nR * nLbl + 1)
    ReDim дво€ко(1 To nR * nLbl + 1)
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 And inLev(i, k) > 0 Then
                    м = k * 1000 + ћин(ћакс(inInd(i, k), 0), 999)
                    For a = 1 To n
                        If место(a) = м Then Exit For
                    Next a
                    If a > n Then
                        n = n + 1: место(n) = м: ур(n) = inLev(i, k)
                    ElseIf ур(a) <> inLev(i, k) Then
                        дво€ко(a) = True
                    End If
                End If
            Next k
        End If
    Next i
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK Then
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 And inLev(i, k) <= 0 Then
                    м = k * 1000 + ћин(ћакс(inInd(i, k), 0), 999)
                    For a = 1 To n
                        If место(a) = м Then
                            If Not дво€ко(a) Then inLev(i, k) = ур(a)
                            Exit For
                        End If
                    Next a
                End If
            Next k
        End If
    Next i
End Sub

Private Sub ¬ставить«нач(ByRef a() As Long, ByRef n As Long, ByVal v As Long)
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

' ”ровень пустого элемента, если его знает сводна€
Private Function ѕустой”ровень(ByVal i As Long) As Long
    Dim k As Long, j As Long, v As Long, kk As Long
    For k = nLbl To 1 Step -1
        If inLev(i, k) > 0 Then
            v = inLev(i, k)
            ' перевод номера пол€ в номер уровн€
            For j = 1 To nR
                For kk = 1 To nLbl
                    If mCellLev(j, kk) > 0 And inLev(j, kk) = v Then
                        ѕустой”ровень = mCellLev(j, kk)
                        Exit Function
                    End If
                Next kk
            Next j
        End If
    Next k
End Function

' —труктурный макет с итогами внизу даЄт дл€ группы ƒ¬≈ строки: пустой
' заголовок Ђ—еверї сверху и Ђ—евер »тогї снизу. ¬ отчЄте нужна одна Ч
' итогова€, подн€та€ наверх. «аголовок без значений поглощаетс€ итогом.
Private Sub ѕоглотить«аголовки()
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
                                ' подпись заголовка точнее, чем Ђ—еверї из Ђ—евер »тогї
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
'  яƒ–ќ: ѕќ–яƒќ  —“–ќ 
'  ѕор€док групп Ч как в сводной (по первому по€влению), итог уровн€ Ч
'  над своими детьми, общий итог Ч первой строкой. —ортировка устойчива€:
'  внутри группы строки не перемешиваютс€.
'==========================================================================
Private Sub ѕор€док—трок()

    Dim i As Long, k As Long, n As Long, r As Long, m As Long, a As Long
    Dim idx() As Long, ключ() As String, s As String
    Dim ранг() As Long, сп() As Long, пк() As String

    ReDim mOrd(1 To nR)
    nOrd = 0

    ' общий итог Ч первой строкой
    For i = 1 To nR
        If mType(i) = LT_GRAND Then nOrd = nOrd + 1: mOrd(nOrd) = i
    Next i

    ' –анг группы уровн€ k = номер строки, где группа по€вилась впервые.
    ' —читаетс€ сортировкой путей (nЈlog n), без словарей: так быстро
    ' и в Excel, и где угодно ещЄ, на любом объЄме.
    ReDim ранг(1 To nR, 1 To mMaxD)
    ReDim сп(1 To nR)
    ReDim пк(1 To nR)
    For k = 1 To mMaxD
        m = 0
        For i = 1 To nR
            If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK And Not mDrop(i) Then
                If mDepth(i) >= k Then m = m + 1: сп(m) = i: пк(m) = mPath(i, k)
            End If
        Next i
        If m > 0 Then
            —ортировать—ли€нием сп, пк, m
            a = 1
            For i = 2 To m + 1
                If i > m Then
                    s = Chr$(2)
                Else
                    s = пк(i)
                End If
                If i > m Or s <> пк(a) Then
                    ' сортировка устойчива€: первый в серии Ч первое по€вление
                    For r = a To i - 1
                        ранг(сп(r), k) = сп(a)
                    Next r
                    a = i
                End If
            Next i
        End If
    Next k

    ReDim idx(1 To nR)
    ReDim ключ(1 To nR)
    n = 0
    For i = 1 To nR
        If mType(i) <> LT_GRAND And mType(i) <> LT_BLANK And Not mDrop(i) Then
            n = n + 1
            idx(n) = i
            s = ""
            For k = 1 To mMaxD
                If k <= mDepth(i) Then
                    r = ранг(i, k)
                Else
                    r = 0          ' сво€ строка группы Ч раньше детей
                End If
                s = s & Format$(r, "0000000")
            Next k
            ключ(n) = s & Format$(i, "0000000")
        End If
    Next i

    If n > 0 Then —ортировать—ли€нием idx, ключ, n
    For i = 1 To n
        nOrd = nOrd + 1
        mOrd(nOrd) = idx(i)
    Next i

End Sub

' ”стойчива€ сортировка сли€нием индексов по строковым ключам
Private Sub —ортировать—ли€нием(ByRef idx() As Long, ByRef ключ() As String, ByVal n As Long)
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

' ѕосле упор€дочивани€: кто кому родитель, какие подписи показывать,
' уровень группировки строк.
Private Sub ядроѕослеѕор€дка()

    Dim p As Long, i As Long, k As Long, d As Long, q As Long, c As Long
    Dim свой() As String, пред As Long, приставка As Long, лв As Long
    Dim txt As String

    ReDim m–одитель(1 To nR)
    ReDim m≈сть–одитель(1 To nR, 1 To mMaxD)
    ReDim mOutLbl(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutR(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutC(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutInd(1 To nOrd + 1, 1 To nLbl)
    ReDim mOutLvl(1 To nOrd + 1)
    ReDim свой(1 To mMaxD)

    ' родитель Ч если следующей идЄт строка той же группы глубже
    For p = 1 To nOrd
        i = mOrd(p)
        If mType(i) = LT_SUBTOTAL Then m–одитель(i) = True
        If mType(i) <> LT_GRAND And p < nOrd Then
            q = mOrd(p + 1)
            d = mDepth(i)
            If mType(q) <> LT_GRAND And mDepth(q) > d Then
                If mPath(q, d) = mPath(i, d) Then m–одитель(i) = True
            End If
        End If
    Next p

    ' есть ли над строкой отдельна€ строка еЄ предка уровн€ k
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
            ' подпись Ч как в сводной: при нескольких мерах в строках итогов
            ' несколько (Ђ»тог —умма выручкиї, Ђ»тог  оличествої)
            txt = ""
            For k = 1 To nLbl
                If Len(Trim$(inTxt(i, k))) > 0 Then txt = inTxt(i, k): Exit For
            Next k
            If Len(txt) = 0 Then txt = “екстќбщего»тога()
            mOutLbl(p, 1) = txt
            mOutLvl(p) = 1
            пред = 0
            GoTo —лед
        End If

        d = mDepth(i)
        свой(d) = mPath(i, d)
        For k = d + 1 To mMaxD
            свой(k) = ""
        Next k
        For k = 1 To d
            m≈сть–одитель(i, k) = (свой(k) = mPath(i, k) And Len(свой(k)) > 0)
        Next k

        ' собственна€ подпись
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
            If Not m≈сть–одитель(i, k) Then
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
                        ' тот же столбец Ч дописываем путь перед подписью
                        mOutLbl(p, c) = mKey(i, k) & " / " & “екстѕодписи(p, c)
                        mOutR(p, c) = 0
                        mOutC(p, c) = 0
                    End If
                End If
            End If
        Next k

        ' уровень группировки: 1 + число предков, у которых есть сво€ строка
        лв = 1
        For k = 1 To d - 1
            If m≈сть–одитель(i, k) Then лв = лв + 1
        Next k
        If лв > 8 Then лв = 8
        mOutLvl(p) = лв
        пред = i
—лед:
    Next p

End Sub

Private Function “екстѕодписи(ByVal p As Long, ByVal c As Long) As String
    “екстѕодписи = mOutLbl(p, c)
End Function

Private Function ћин(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then ћин = a Else ћин = b
End Function

Private Function ћакс(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then ћакс = a Else ћакс = b
End Function


'==========================================================================
'  яƒ–ќ: Ўјѕ ј —“ќЋЅ÷ќ¬ ѕќ Ћ»—“”
'  ¬ход: hTxt/hCT (строки шапки над значени€ми), inCLT (тип столбца из оси),
'  m—лужебные, mMeas. ¬ыход: cItem/cCnt/cType/cDF/cMeasLev.
'  ѕодпись группы стоит только над первым столбцом группы Ч пуста€ клетка
'  верхних строк наследует подпись слева. ¬ самой нижней строке шапки
'  наследовани€ нет: пуста€ клетка там Ч пустой элемент, а не продолжение.
'==========================================================================
Private Sub ЎапкаѕоЋисту()

    Dim r As Long, j As Long, k As Long, q As Long, t As Long, ct As Long
    Dim нужна() As Boolean, ур() As Long, стек() As String, s As String
    Dim typ As Long, итогЌа As Long, база As String, лучш As Long, n As Long, m As Long

    ReDim нужна(1 To nHR)
    For r = 1 To nHR
        For j = 1 To nC
            ct = hCT(r, j)
            If ct = CT_ITEM Or ct = CT_SUBTOTAL Or ct = CT_GRAND Or ct = CT_DATAFIELD Or ct = CT_CUSTOMSUB Then
                нужна(r) = True
            ElseIf ct = -1 And Len(hTxt(r, j)) > 0 And Not —лужебна€(hTxt(r, j)) Then
                нужна(r) = True
            End If
        Next j
    Next r
    t = 0
    ReDim ур(1 To nHR)
    For r = 1 To nHR
        If нужна(r) Then t = t + 1: ур(t) = r
    Next r
    If t = 0 Then t = 1: ур(1) = nHR

    ReDim cItem(1 To nC, 1 To t)
    ReDim стек(1 To t)
    ReDim cCnt(1 To nC)
    ReDim cType(1 To nC)
    ReDim cDF(1 To nC)

    ' уровень мер Ч строка, где больше всего подписей совпадает с мерами
    cMeasLev = 0: лучш = 0
    For k = 1 To t
        n = 0
        For j = 1 To nC
            If Len(ћераѕо“ексту(hTxt(ур(k), j), True)) > 0 Then n = n + 1
        Next j
        If n > лучш Then лучш = n: cMeasLev = k
    Next k

    For j = 1 To nC
        ' тип столбца: ось, иначе тип €чейки, иначе текст
        typ = inCLT(j)
        If typ < 0 Then
            typ = LT_REGULAR
            For k = 1 To t
                ct = hCT(ур(k), j)
                s = hTxt(ур(k), j)
                If ct = CT_GRAND Then typ = LT_GRAND
                If (ct = CT_SUBTOTAL Or ct = CT_CUSTOMSUB) And typ <> LT_GRAND Then typ = LT_SUBTOTAL
                If ct = -1 And Len(s) > 0 Then
                    If Ётоќбщий»тог(s) Or Left$(LCase$(s), 5) = "итог " Or Left$(LCase$(s), 6) = "total " Then
                        typ = LT_GRAND
                    ElseIf Len(—уффикс»тога(s)) > 0 And typ <> LT_GRAND Then
                        typ = LT_SUBTOTAL
                    End If
                End If
            Next k
        End If
        If typ = LT_BLANK Then typ = LT_REGULAR
        cType(j) = typ

        ' клетка с подписью итога Ч сама€ нижн€€ непуста€
        итогЌа = 0
        If typ <> LT_REGULAR Then
            For k = t To 1 Step -1
                If Len(hTxt(ур(k), j)) > 0 And Not —лужебна€(hTxt(ур(k), j)) Then итогЌа = k: Exit For
            Next k
            If итогЌа = 0 Then итогЌа = 1
        End If

        For k = 1 To t
            s = hTxt(ур(k), j)
            If —лужебна€(s) Then s = ""
            If typ <> LT_REGULAR And k = итогЌа Then
                ' итог: путь Ч подписи групп над ним, сам уровень Ч им€ группы
                For q = 1 To k - 1
                    cItem(j, q) = стек(q)
                Next q
                If typ = LT_SUBTOTAL Then
                    база = —уффикс»тога(s)
                    If Len(база) = 0 Then база = стек(k)
                    If Len(база) = 0 Then база = s
                    cItem(j, k) = база
                Else
                    cItem(j, k) = s
                End If
                cCnt(j) = k
                Exit For
            End If
            If Len(s) > 0 Then
                стек(k) = s
                For q = k + 1 To t
                    стек(q) = ""
                Next q
            ElseIf k = t Then
                стек(k) = ""            ' пустой элемент на нижнем уровне
            End If
            cItem(j, k) = стек(k)
        Next k
        If typ = LT_REGULAR Then cCnt(j) = t

        ' мера столбца
        If cMeasLev > 0 And cMeasLev <= cCnt(j) Then cDF(j) = ћераѕо“ексту(cItem(j, cMeasLev), True)
        If Len(cDF(j)) = 0 Then
            For k = 1 To t
                If Len(hTxt(ур(k), j)) > 0 Then
                    m = 0
                    cDF(j) = ћераѕо“ексту(hTxt(ур(k), j), False)
                    If Len(cDF(j)) > 0 Then Exit For
                End If
            Next k
        End If
    Next j

End Sub

Private Function —лужебна€(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    —лужебна€ = (InStr(1, m—лужебные, "|" & LCase$(Trim$(s)) & "|", vbBinaryCompare) > 0)
End Function

' ћера, названна€ в тексте: точно = подпись равна имени меры,
' иначе Ч им€ меры входит в подпись (Ђ»тог дельта суммыї); берЄм самое длинное
Private Function ћераѕо“ексту(ByVal s As String, ByVal точно As Boolean) As String
    Dim q As Long, лучш As String
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For q = 1 To nMeas
        If Len(mMeas(q)) > 0 Then
            If точно Then
                If StrComp(s, mMeas(q), vbTextCompare) = 0 Then ћераѕо“ексту = mMeas(q): Exit Function
            ElseIf InStr(1, s, mMeas(q), vbTextCompare) > 0 Then
                If Len(mMeas(q)) > Len(лучш) Then лучш = mMeas(q)
            End If
        End If
    Next q
    If Not точно Then ћераѕо“ексту = лучш
End Function


'==========================================================================
'  яƒ–ќ: Ўјѕ ј —“ќЋЅ÷ќ¬
'  ¬ход: cItem/cCnt/cType/cDF/cMeasLev/mMeas. ¬ыход: dLev/dDepth/nLev.
'==========================================================================
Private Sub —обратьЎапку()

    Dim j As Long, k As Long, maxLev As Long, L As Long
    Dim мера() As String
    Dim счЄт As Long, пусто As Boolean

    maxLev = 1
    For j = 1 To nC
        If cCnt(j) > maxLev Then maxLev = cCnt(j)
    Next j
    If UBound(cItem, 2) > maxLev Then maxLev = UBound(cItem, 2)

    ¬осстановитьѕуть»тогов cItem, cCnt, cType, nC, maxLev

    ' --- мера каждого столбца (нужна столбцам итогов) ---
    ReDim мера(1 To nC)
    счЄт = 0
    For j = 1 To nC
        If Len(cDF(j)) > 0 Then
            мера(j) = cDF(j)
        ElseIf cType(j) <> LT_REGULAR And nMeas > 1 Then
            ' Excel выдаЄт по одному столбцу итога на каждую меру подр€д,
            ' в пор€дке мер Ч по позиции в серии и определ€ем меру
            If j > 1 Then
                If cType(j - 1) = cType(j) And ќдин люч(cItem, cCnt, j - 1, j) Then
                    счЄт = счЄт + 1
                Else
                    счЄт = 1
                End If
            Else
                счЄт = 1
            End If
            If счЄт >= 1 And счЄт <= nMeas Then мера(j) = mMeas(счЄт)
        ElseIf cMeasLev > 0 And cMeasLev <= cCnt(j) Then
            мера(j) = cItem(j, cMeasLev)
        End If
    Next j

    ' --- уровни шапки Ч в том же пор€дке, что в сводной ---
    ' (меры остаютс€ на своЄм месте: сверху, если в сводной Ђ«начени€ї
    ' сто€т первыми, или под элементами, если поле стоит выше Ђ«наченийї)
    nLev = maxLev
    ReDim dLev(1 To nC, 1 To nLev)
    ReDim dDepth(1 To nC)

    For j = 1 To nC
        L = cCnt(j)
        If L < 1 Then L = 1
        If L > nLev Then L = nLev
        If cType(j) = LT_REGULAR Then
            For k = 1 To L
                dLev(j, k) = cItem(j, k)
            Next k
        Else
            ' итог: подписи групп над ним, сам он Ч на своЄм уровне
            For k = 1 To L - 1
                dLev(j, k) = cItem(j, k)
            Next k
            If L = cMeasLev And Len(мера(j)) > 0 Then
                ' подпись итога стоит в строке мер: мера Ч на еЄ уровне,
                ' пометка итога Ч уровнем ниже
                dLev(j, L) = мера(j)
                If L + 1 <= nLev Then
                    If cType(j) = LT_GRAND Then dLev(j, L + 1) = “екстќбщего»тога() Else dLev(j, L + 1) = TXT_SUBTOT
                End If
            ElseIf cType(j) = LT_GRAND Then
                dLev(j, L) = “екстќбщего»тога()
            Else
                dLev(j, L) = cItem(j, L)
                If L + 1 <= nLev And L + 1 <> cMeasLev Then dLev(j, L + 1) = TXT_SUBTOT
            End If
            ' мера итога Ч на уровне мер, если он ниже
            If cMeasLev > L And cMeasLev <= nLev And Len(мера(j)) > 0 Then dLev(j, cMeasLev) = мера(j)
        End If
        dDepth(j) = L
    Next j

    ”братьѕустые”ровни

    ' ≈сли у сводной нет полей в столбцах, шапка окажетс€ пустой Ч
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
            If Len(Trim$(dLev(j, 1))) = 0 Then dLev(j, 1) = "«начение"
        End If
    Next j

End Sub

' ”ровень, пустой у всех столбцов (так бывает у безым€нных членов OLAP),
' занимает строку в шапке и ничего не сообщает Ч выкидываем его.
Private Sub ”братьѕустые”ровни()

    Dim k As Long, j As Long, n As Long, t As Long
    Dim держать() As Boolean
    Dim новый() As String, нов√луб() As Long

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
    ReDim нов√луб(1 To nC)
    For j = 1 To nC
        t = 0
        нов√луб(j) = 0
        For k = 1 To nLev
            If держать(k) Then
                t = t + 1
                новый(j, t) = dLev(j, k)
                If k <= dDepth(j) Then нов√луб(j) = t
            End If
        Next k
        If нов√луб(j) < 1 Then нов√луб(j) = 1
    Next j

    nLev = n
    dLev = новый
    dDepth = нов√луб

End Sub

' ƒл€ линии промежуточного итога Excel отдаЄт не путь предков, а сам
' итожимый элемент (на модели данных Ч ещЄ и продублированный). »з-за этого
' итог оказываетс€ в отдельной группе и уезжает вправо. ¬осстанавливаем путь
' по ЅЋ»∆ј…Ў≈… обычной линии с этим элементом: итог стоит вплотную к своей
' группе, а одноимЄнные элементы других родителей Ч дальше.
Private Sub ¬осстановитьѕуть»тогов(ByRef item() As String, ByRef cnt() As Long, _
                                  ByRef тип() As Long, ByVal n As Long, ByVal maxLev As Long)
    Dim j As Long, q As Long, k As Long, шаг As Long
    Dim посл As String, строка As Long, уровень As Long

    For j = 1 To n
        If тип(j) = LT_SUBTOTAL Then
            ' ќбычна€ сводна€ уже отдаЄт полный путь предков Ч трогать нельз€,
            ' иначе одноимЄнные элементы разных родителей скле€тс€ в одну группу.
            If ѕуть—огласован(item, cnt, тип, n, j) Then GoTo ƒальше
            посл = item(j, cnt(j))
            строка = 0: уровень = 0
            For шаг = 1 To 2 * n
                ' сначала ближайша€ слева, потом справа, расшир€€ окно
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
ƒальше:
    Next j
End Sub

' ≈сть ли обычна€ лини€, у которой первые cnt(j) уровней совпадают с путЄм
' линии итога. ≈сли есть Ч путь насто€щий, восстанавливать нечего.
Private Function ѕуть—огласован(ByRef item() As String, ByRef cnt() As Long, _
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
            If совпало Then ѕуть—огласован = True: Exit Function
        End If
    Next q
End Function

Private Function ќдин люч(ByRef item() As String, ByRef cnt() As Long, _
                          ByVal a As Long, ByVal b As Long) As Boolean
    Dim k As Long
    If cnt(a) <> cnt(b) Then Exit Function
    For k = 1 To cnt(a)
        If StrComp(item(a, k), item(b, k), vbTextCompare) <> 0 Then Exit Function
    Next k
    ќдин люч = True
End Function


'==========================================================================
'  яƒ–ќ: ѕќ–яƒќ  —“ќЋЅ÷ќ¬ Ч итог группы левее еЄ столбцов
'==========================================================================
Private Sub ѕор€док—толбцов()
    Dim j As Long, ост() As Long, n As Long
    ReDim cOrd(1 To nC)
    nCOrd = 0

    ReDim ост(1 To nC)
    n = 0
    For j = 1 To nC
        If cType(j) <> LT_BLANK Then n = n + 1: ост(n) = j
    Next j
    If n > 0 Then ѕор€док—толбцов–ек 1, ост, n
End Sub

Private Sub ѕор€док—толбцов–ек(ByVal d As Long, ByRef idx() As Long, ByVal n As Long)

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

    ' различные ключи уровн€ в пор€дке первого по€влени€
    ReDim ключи(1 To n)
    nk = 0
    For i = 1 To n
        найден = False
        For k = 1 To nk
            If StrComp(dLev(idx(i), d), ключи(k), vbTextCompare) = 0 Then найден = True: Exit For
        Next k
        If Not найден Then nk = nk + 1: ключи(nk) = dLev(idx(i), d)
    Next i

    ' группы, целиком состо€щие из общего итога, выносим в начало уровн€
    пер = 0
    For k = 1 To nk
        If “олько»тог(idx, n, d, ключи(k)) Then
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
                    nс = nс + 1: свои(nс) = idx(i)      ' итог уровн€ Ч впереди
                Else
                    nд = nд + 1: дети(nд) = idx(i)
                End If
            End If
        Next i
        For i = 1 To nс
            nCOrd = nCOrd + 1: cOrd(nCOrd) = свои(i)
        Next i
        If nд > 0 Then ѕор€док—толбцов–ек d + 1, дети, nд
    Next k

End Sub

Private Function “олько»тог(ByRef idx() As Long, ByVal n As Long, _
                            ByVal d As Long, ByVal ключ As String) As Boolean
    Dim i As Long, есть As Boolean
    For i = 1 To n
        If StrComp(dLev(idx(i), d), ключ, vbTextCompare) = 0 Then
            есть = True
            If cType(idx(i)) <> LT_GRAND Then Exit Function
        End If
    Next i
    “олько»тог = есть
End Function

' ѕлан объединени€ шапки: пр€моугольники (уровень1, столбец1)-(уровень2, столбец2)
' в координатах шапки. ѕо горизонтали Ч соседние столбцы с одинаковыми
' подпис€ми на всех уровн€х до k. ѕо вертикали Ч столбец, у которого нижние
' уровни пустые (итоги, пустой элемент), но только если его подпись не
' объединена по горизонтали с сосед€ми: иначе объединени€ пересекутс€,
' Excel сольЄт их в один блок и сотрЄт подписи нижнего уровн€.
Private Sub ѕланЎапки()

    Dim k As Long, ci As Long, a As Long, п As Long, кк As Long
    Dim шир() As Long          ' (уровень, позици€) Ч ширина горизонтального объединени€

    nќбъед = 0
    ReDim mќб 1(1 To 2 * nLev * nCOrd + 1)
    ReDim mќб—1(1 To 2 * nLev * nCOrd + 1)
    ReDim mќб 2(1 To 2 * nLev * nCOrd + 1)
    ReDim mќб—2(1 To 2 * nLev * nCOrd + 1)
    ReDim шир(1 To nLev, 1 To nCOrd)

    For k = 1 To nLev
        ci = 1
        Do While ci <= nCOrd
            a = ci
            Do While ci + 1 <= nCOrd
                If Not ќдинаковыƒо(cOrd(a), cOrd(ci + 1), k) Then Exit Do
                If Len(Trim$(dLev(cOrd(a), k))) = 0 Then Exit Do
                ci = ci + 1
            Loop
            For п = a To ci
                шир(k, п) = ci - a + 1
            Next п
            If ci > a Then
                nќбъед = nќбъед + 1
                mќб 1(nќбъед) = k: mќб—1(nќбъед) = a
                mќб 2(nќбъед) = k: mќб—2(nќбъед) = ci
            End If
            ci = ci + 1
        Loop
    Next k

    For ci = 1 To nCOrd
        кк = 0
        For п = nLev To 1 Step -1
            If Len(Trim$(dLev(cOrd(ci), п))) > 0 Then кк = п: Exit For
        Next п
        If кк >= 1 And кк < nLev Then
            If шир(кк, ci) = 1 Then
                nќбъед = nќбъед + 1
                mќб 1(nќбъед) = кк: mќб—1(nќбъед) = ci
                mќб 2(nќбъед) = nLev: mќб—2(nќбъед) = ci
            End If
        End If
    Next ci

End Sub

Private Function ќдинаковыƒо(ByVal a As Long, ByVal b As Long, ByVal k As Long) As Boolean
    Dim i As Long
    For i = 1 To k
        If StrComp(dLev(a, i), dLev(b, i), vbTextCompare) <> 0 Then Exit Function
    Next i
    ќдинаковыƒо = True
End Function

'  ласс строки дл€ раскраски Ч по ”–ќ¬Ќё, а не по наличию детей:
' все строки одного уровн€ выгл€д€т одинаково (так в дизайне: итог,
' уровень 1, 2, 3...). Ќижний уровень сводной Ч детализаци€, без заливки.
' "G" Ч общий итог, "S<n>" Ч уровень n, "" Ч детализаци€.
Private Function  ласс—троки(ByVal i As Long) As String
    If mType(i) = LT_GRAND Then
         ласс—троки = "G"
    ElseIf mType(i) = LT_SUBTOTAL Or mDepth(i) < mMaxD Then
         ласс—троки = "S" & mDepth(i)
    Else
         ласс—троки = ""
    End If
End Function
'#яƒ–ќ- ќЌ≈÷


'==========================================================================
'  Ћ»—“-ѕ–»®ћЌ» 
'==========================================================================
Private Function —оздатьЋист(pt As PivotTable) As Worksheet

    Dim otvet As VbMsgBoxResult
    Dim wb As Workbook
    Dim baza As String, imya As String, i As Long

    If mQuiet Then
        otvet = vbYes
    Else
        otvet = MsgBox(" уда положить отчЄт?" & vbLf & vbLf & _
                       "ƒа Ч новый лист в этой книге" & vbLf & _
                       "Ќет Ч нова€ книга", _
                       vbYesNoCancel + vbQuestion, "–азмещение отчЄта")
    End If
    If otvet = vbCancel Then Exit Function

    If otvet = vbYes Then
        Set wb = pt.Parent.Parent
    Else
        Set wb = Workbooks.Add(xlWBATWorksheet)
    End If

    baza = Left$(„истое»м€Ћиста(pt.Parent.Name), 24) & "_отчет"
    imya = baza
    i = 1
    Do While Ћист≈сть(wb, imya)
        i = i + 1
        imya = baza & CStr(i)
    Loop

    If otvet = vbNo Then
        Set —оздатьЋист = wb.Worksheets(1)
    Else
        Set —оздатьЋист = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    End If
    —оздатьЋист.Name = imya

End Function

Private Function „истое»м€Ћиста(ByVal s As String) As String
    Dim x As Variant
    For Each x In Array(":", "\", "/", "?", "*", "[", "]")
        s = Replace(s, x, "_")
    Next x
    „истое»м€Ћиста = s
End Function

Private Function Ћист≈сть(wb As Workbook, ByVal imya As String) As Boolean
    Dim ws As Object
    For Each ws In wb.Sheets
        If StrComp(ws.Name, imya, vbTextCompare) = 0 Then Ћист≈сть = True: Exit Function
    Next ws
End Function


'==========================================================================
'  ¬џ¬ќƒ » ќ‘ќ–ћЋ≈Ќ»≈
'==========================================================================
Private Sub ¬ывестиќтчЄт(ws As Worksheet, ByRef lbl() As Variant, ByRef vals() As Variant)

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

    ' ---------- классы строк и колонок ----------
    ReDim m ласс—тр(1 To nOrd)
    ReDim m ласс ол(1 To nCOrd)
    For ci = 1 To nCOrd
        j = cOrd(ci)
        If cType(j) = LT_GRAND Then
            m ласс ол(ci) = "G"
        ElseIf cType(j) = LT_SUBTOTAL Then
            m ласс ол(ci) = "S"
        Else
            m ласс ол(ci) = ""
        End If
    Next ci
    For ri = 1 To nOrd
        i = mOrd(ri)
        If mType(i) = LT_GRAND Then
            m ласс—тр(ri) = "G"
        Else
            m ласс—тр(ri) =  ласс—троки(i)
        End If
    Next ri

    ' ---------- метки ролей: до любого форматировани€ ----------
    ѕометить ws.Cells(1, cLbl), "T"
    ѕометить ws.Range(ws.Cells(2, cLbl), ws.Cells(3, cLbl)), "P"
    ѕометить ws.Range(ws.Cells(rHdr, 1), ws.Cells(rNum - 1, cLast)), "H"
    ѕометить ws.Range(ws.Cells(rNum, 1), ws.Cells(rNum, cData - 1)), "CL"
    ѕометить уски ws, rNum, cData, nCOrd, False
    ѕометить уски ws, rData, 1, nOrd, True

    ws.Cells.Font.Name = FNT

    ' ---------- шапка листа ----------
    With ws.Cells(1, cLbl)
        .Value = TXT_TITLE
        .Font.Size = FS_TITLE: .Font.Bold = True
    End With
    With ws.Cells(2, cLbl)
        .Value = TXT_SUB
        .Font.Size = 10: .Font.Bold = True
    End With
    With ws.Cells(3, cLbl)               ' фильтры сводной
        .Value = “екстячейки(m‘ильтры)
        .Font.Size = 9
    End With

    ' ---------- шапка таблицы ----------
    ReDim a(1 To nLev, 1 To nCOrd)
    For k = 1 To nLev
        For ci = 1 To nCOrd
            a(k, ci) = “екстячейки(dLev(cOrd(ci), k))
        Next ci
    Next k
    ws.Range(ws.Cells(rHdr, cData), ws.Cells(rNum - 1, cLast)).Value = a
    ' подписи слева Ч в верхнюю строку шапки: ниже они объедин€ютс€ по высоте,
    ' а при объединении остаЄтс€ значение верхней клетки
    For k = 1 To nLbl
        ws.Cells(rHdr, cLbl + k - 1).Value = “екстячейки(mLblCap(k))
    Next k
    ws.Cells(rHdr, 1).Value = "є"

    With ws.Range(ws.Cells(rHdr, 1), ws.Cells(rNum - 1, cLast))
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

    ќбъединитьЎапку ws, rHdr, cData

    ' ---------- нумераци€ столбцов ----------
    ReDim a(1 To 1, 1 To nLbl + nCOrd)
    For j = 1 To nLbl + nCOrd
        a(1, j) = j
    Next j
    With ws.Range(ws.Cells(rNum, cLbl), ws.Cells(rNum, cLast))
        .Value = a
        .Font.Size = FS_NUM
        .HorizontalAlignment = xlCenter
    End With

    ' ---------- подписи строк Ч одним массивом ----------
    ReDim a(1 To nOrd, 1 To nLbl)
    For ri = 1 To nOrd
        For k = 1 To nLbl
            If mOutR(ri, k) > 0 Then
                a(ri, k) = «начениеѕодписи(lbl(mOutR(ri, k), mOutC(ri, k)))
            Else
                a(ri, k) = “екстячейки(mOutLbl(ri, k))
            End If
        Next k
    Next ri
    With ws.Range(ws.Cells(rData, cLbl), ws.Cells(rLast, cLbl + nLbl - 1))
        .Value = a
        .HorizontalAlignment = xlLeft
        .Font.Size = FS_BODY
        .Font.Color = C_“≈ —“
    End With
    ' отступы Ч сплошными кусками одного отступа, не по €чейке
    Dim нач As Long, отс As Long
    For k = 1 To nLbl
        нач = 1
        For ri = 2 To nOrd + 1
            If ri > nOrd Then отс = -99 Else отс = mOutInd(ri, k)
            If отс <> mOutInd(нач, k) Then
                If mOutInd(нач, k) > 0 Then
                    ws.Range(ws.Cells(rData + нач - 1, cLbl + k - 1), _
                             ws.Cells(rData + ri - 2, cLbl + k - 1)).IndentLevel = ћин(mOutInd(нач, k), 15)
                End If
                нач = ri
            End If
        Next ri
    Next k

    ' ---------- значени€ Ч одним массивом ----------
    ReDim a(1 To nOrd, 1 To nCOrd)
    For ri = 1 To nOrd
        i = mOrd(ri)
        For ci = 1 To nCOrd
            v = vals(i, cOrd(ci))
            If ѕустое«начение(v) Then a(ri, ci) = "Ц" Else a(ri, ci) = v
        Next ci
    Next ri
    Set rng = ws.Range(ws.Cells(rData, cData), ws.Cells(rLast, cLast))
    rng.Value = a
    rng.HorizontalAlignment = xlCenter
    rng.Font.Size = FS_BODY
    rng.Font.Color = C_“≈ —“
    ‘орматы„исел ws, rData, cData, vals

    ' прочерки вместо пустых Ч серым
    ' (SpecialCells на одной €чейке берЄт весь лист Ч еЄ провер€ем сами)
    If rng.Cells.Count = 1 Then
        If VarType(rng.Value) = vbString Then rng.Font.Color = C_ѕ–ќ„≈– 
    Else
        On Error Resume Next
        rng.SpecialCells(xlCellTypeConstants, xlTextValues).Font.Color = C_ѕ–ќ„≈– 
        On Error GoTo 0
    End If

    ' нумераци€ строк
    ReDim a(1 To nOrd, 1 To 1)
    For ri = 1 To nOrd
        a(ri, 1) = ri
    Next ri
    With ws.Range(ws.Cells(rData, 1), ws.Cells(rLast, 1))
        .Value = a
        .Font.Size = FS_NUM
        .HorizontalAlignment = xlCenter
    End With

    ' ---------- итоги жирным, цвета Ч по палитре ----------
    For ri = 1 To nOrd
        If Len(m ласс—тр(ri)) > 0 Then
            ws.Range(ws.Cells(rData + ri - 1, cLbl), ws.Cells(rData + ri - 1, cLast)).Font.Bold = True
        End If
    Next ri
    For ci = 1 To nCOrd
        If Len(m ласс ол(ci)) > 0 Then
            ws.Range(ws.Cells(rData, cData + ci - 1), ws.Cells(rLast, cData + ci - 1)).Font.Bold = True
        End If
    Next ci
    ѕокрасить ws, rHdr, rNum, rData, rLast, cLbl, cData, cLast

    ' ---------- геометри€ ----------
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
    ' строки таблицы Ч чуть выше стандартных, текст по центру по вертикали
    ws.Range(ws.Rows(rData), ws.Rows(rLast)).RowHeight = 15
    ws.Range(ws.Cells(rHdr, 1), ws.Cells(rLast, cLast)).VerticalAlignment = xlCenter

    ' ---------- группировка строк по уровн€м ----------
    If √–”ѕѕ»–ќ¬ј“№ Then —группировать ws, rData

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Cells(rData, cData).Select
    ActiveWindow.FreezePanes = True

    ѕеренести”словное‘орматирование ws, rData, cData

    Ќастроитьѕечать ws, rHdr, rNum, rLast, cLast

End Sub

' —трока, похожа€ на число или дату, но бывша€ в сводной текстом,
' должна остатьс€ текстом Ч иначе Ђ1-2ї превратитс€ в 1 феврал€.
Private Function “екстячейки(ByVal s As String) As Variant
    If Len(s) > 0 Then
        If IsNumeric(s) Or IsDate(s) Or Left$(s, 1) = "=" Then
            “екстячейки = "'" & s
            Exit Function
        End If
    End If
    “екстячейки = s
End Function

Private Function «начениеѕодписи(v As Variant) As Variant
    If VarType(v) = vbString Then
        «начениеѕодписи = “екстячейки(CStr(v))
    ElseIf IsError(v) Then
        «начениеѕодписи = ""
    Else
        «начениеѕодписи = v
    End If
End Function

' ‘ормат каждой колонки берЄм у меры в сводной; Ђќбщийї замен€ем
' на разделители разр€дов, дроби Ч только если они есть в данных.
Private Sub ‘орматы„исел(ws As Worksheet, ByVal rData As Long, ByVal cData As Long, _
                         ByRef vals() As Variant)
    Dim ci As Long, ri As Long, j As Long, i As Long, f As String
    Dim по—трокам As Boolean
    Dim rLast As Long
    rLast = rData + nOrd - 1

    For j = 1 To nC
        If Len(m‘ормат ол(j)) = 0 Then по—трокам = True
    Next j

    If Not по—трокам Then
        For ci = 1 To nCOrd
            j = cOrd(ci)
            f = m‘ормат ол(j)
            If ќбщий‘ормат(f) Then f = јвто‘ормат(vals, 0, j)
            ws.Range(ws.Cells(rData, cData + ci - 1), ws.Cells(rLast, cData + ci - 1)).NumberFormat = f
        Next ci
    Else
        For ri = 1 To nOrd
            i = mOrd(ri)
            f = m‘ормат—тр(i)
            If ќбщий‘ормат(f) Then f = јвто‘ормат(vals, i, 0)
            ws.Range(ws.Cells(rData + ri - 1, cData), ws.Cells(rData + ri - 1, cData + nCOrd - 1)).NumberFormat = f
        Next ri
    End If
End Sub

Private Function ќбщий‘ормат(ByVal f As String) As Boolean
    f = LCase$(Trim$(f))
    ќбщий‘ормат = (f = "" Or f = "general" Or f = "общий" Or f = "основной" Or f = "@")
End Function

' i = 0 Ч по колонке j, j = 0 Ч по строке i
Private Function јвто‘ормат(ByRef vals() As Variant, ByVal i As Long, ByVal j As Long) As String
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
        јвто‘ормат = "#,##0.00;-#,##0.00;""Ц"""
    Else
        јвто‘ормат = "#,##0;-#,##0;""Ц"""
    End If
End Function

' ”ровни строк -> структура листа: кнопки 1-2-3... сворачивают отчЄт
' до нужной глубины. »тог над группой, поэтому SummaryRow = над.
Private Sub —группировать(ws As Worksheet, ByVal rData As Long)
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
'  «ј –ј— ј, √–јЌ»÷џ, ”—Ћќ¬Ќќ≈ ‘ќ–ћј“»–ќ¬јЌ»≈
'==========================================================================
' ”словное форматирование сводной (цветовые шкалы, гистограммы, наборы
' значков, правила по значению) переноситс€ на соответствующие €чейки
' отчЄта с учЄтом того, что строки и колонки переставлены.
Private Sub ѕеренести”словное‘орматирование(ws As Worksheet, _
                                            ByVal rData As Long, ByVal cData As Long)

    Dim fc As Object, обл As Range, целевой As Range, кус As Range
    Dim строка≈сть() As Boolean, столбец≈сть() As Boolean
    Dim k As Long, i As Long, j As Long, ri As Long, ci As Long
    Dim нач As Long, попал As Boolean, всего As Long

    If m»сходник Is Nothing Then Exit Sub
    On Error Resume Next
    всего = m»сходник.FormatConditions.Count
    On Error GoTo 0
    If всего = 0 Then Exit Sub

    ReDim строка≈сть(1 To nR)
    ReDim столбец≈сть(1 To nC)

    For k = 1 To всего

        Set fc = Nothing
        Set обл = Nothing
        On Error Resume Next
        Set fc = m»сходник.FormatConditions(k)
        Set обл = Application.Intersect(fc.AppliesTo, m»сходник)
        On Error GoTo 0
        If обл Is Nothing Then GoTo ƒальше

        For i = 1 To nR
            строка≈сть(i) = Not (Application.Intersect(обл, m»сходник.Rows(i)) Is Nothing)
        Next i
        For j = 1 To nC
            столбец≈сть(j) = Not (Application.Intersect(обл, m»сходник.Columns(j)) Is Nothing)
        Next j

        ' собираем адресат сплошными кусками Ч по €чейке было бы слишком долго
        Set целевой = Nothing
        For ci = 1 To nCOrd
            If столбец≈сть(cOrd(ci)) Then
                нач = 0
                For ri = 1 To nOrd + 1
                    If ri <= nOrd Then попал = строка≈сть(mOrd(ri)) Else попал = False
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

        If Not целевой Is Nothing Then —оздать”словие fc, целевой
ƒальше:
    Next k

End Sub

Private Sub —оздать”словие(fc As Object, rng As Range)

    Dim n As Object, i As Long

    On Error Resume Next
    Select Case fc.Type

        Case 3      ' цветова€ шкала
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

        Case 1      ' правило по значению €чейки
            If fc.Operator = xlBetween Or fc.Operator = xlNotBetween Then
                Set n = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=fc.Operator, _
                                                 Formula1:=fc.Formula1, Formula2:=fc.Formula2)
            Else
                Set n = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=fc.Operator, _
                                                 Formula1:=fc.Formula1)
            End If
             опировать‘ормат”слови€ fc, n

        Case 2      ' правило-формула
            Set n = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fc.Formula1)
             опировать‘ормат”слови€ fc, n

        Case Else
            ' остальные типы переносить нечем Ч пропускаем молча

    End Select
    Err.Clear
    On Error GoTo 0

End Sub

Private Sub  опировать‘ормат”слови€(fc As Object, n As Object)
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

' —троки итогов Ч целиком, колонки итогов Ч поверх; всЄ сплошными кусками
' одного цвета, а не по €чейке.
Private Sub «акрасить“ело(ws As Worksheet, ByVal rData As Long, ByVal rLast As Long, _
                          ByVal cLbl As Long, ByVal cData As Long, ByVal cLast As Long)

    Dim ri As Long, ci As Long, нач As Long, n As Long
    Dim база() As Long, цв() As Long

    n = rLast - rData + 1
    ReDim база(1 To n + 1)
    ReDim цв(1 To n + 1)

    For ri = 1 To n
        база(ri) = ÷вет–оли(m ласс—тр(ri), "")
    Next ri
    база(n + 1) = -2
    нач = 1
    For ri = 2 To n + 1
        If база(ri) <> база(нач) Then
            ' строка итога заливаетс€ целиком, вместе с номером Ђєї Ч как в дизайне
            With ws.Range(ws.Cells(rData + нач - 1, 1), ws.Cells(rData + ri - 2, cLast)).Interior
                If база(нач) >= 0 Then .Color = база(нач) Else .Pattern = xlNone
            End With
            нач = ri
        End If
    Next ri

    For ci = 1 To UBound(m ласс ол)
        If Len(m ласс ол(ci)) > 0 Then
            For ri = 1 To n
                цв(ri) = ÷вет–оли(m ласс—тр(ri), m ласс ол(ci))
            Next ri
            цв(n + 1) = -2
            нач = 1
            For ri = 2 To n + 1
                If цв(ri) <> цв(нач) Then
                    If цв(нач) >= 0 Then
                        ws.Range(ws.Cells(rData + нач - 1, cData + ci - 1), _
                                 ws.Cells(rData + ri - 2, cData + ci - 1)).Interior.Color = цв(нач)
                    End If
                    нач = ri
                End If
            Next ri
        End If
    Next ci

End Sub

' ћетки ролей кусками: строки (столбец A) или колонки (строка нумерации)
Private Sub ѕометить уски(ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long, _
                          ByVal n As Long, ByVal строки As Boolean)
    Dim i As Long, нач As Long, м() As String
    ReDim м(1 To n + 1)
    For i = 1 To n
        If строки Then м(i) = ћетка—троки(m ласс—тр(i)) Else м(i) = "C" & m ласс ол(i)
    Next i
    м(n + 1) = "#"
    нач = 1
    For i = 2 To n + 1
        If м(i) <> м(нач) Then
            If строки Then
                ѕометить ws.Range(ws.Cells(r0 + нач - 1, c0), ws.Cells(r0 + i - 2, c0)), м(нач)
            Else
                ѕометить ws.Range(ws.Cells(r0, c0 + нач - 1), ws.Cells(r0, c0 + i - 2)), м(нач)
            End If
            нач = i
        End If
    Next i
End Sub

' ќбъединение заголовков по плану €дра (ѕланЎапки): одинаковые подписи Ч
' по горизонтали, Ђвис€щиеї Ч по вертикали. ѕлан заранее исключает
' пересечени€: Excel сливает пересекающиес€ объединени€ в один блок
' и стирает подписи внутри него.
Private Sub ќбъединитьЎапку(ws As Worksheet, ByVal rHdr As Long, ByVal cData As Long)

    Dim i As Long
    Dim былиѕредупреждени€ As Boolean
    былиѕредупреждени€ = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ѕланЎапки
    For i = 1 To nќбъед
        On Error Resume Next
        ws.Range(ws.Cells(rHdr + mќб 1(i) - 1, cData + mќб—1(i) - 1), _
                 ws.Cells(rHdr + mќб 2(i) - 1, cData + mќб—2(i) - 1)).Merge
        On Error GoTo 0
    Next i

    Application.DisplayAlerts = былиѕредупреждени€
End Sub

Private Sub √раницы(rng As Range)
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

Private Sub Ќастроитьѕечать(ws As Worksheet, ByVal rHdr As Long, ByVal rNum As Long, _
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

    ' ј3 ставим ќ“ƒ≈Ћ№Ќќ: если драйвер принтера его не поддерживает,
    ' присвоение падает и в отложенном режиме обнул€ет всю пачку настроек.
    Err.Clear
    ws.PageSetup.PaperSize = xlPaperA3
    Err.Clear
    On Error GoTo 0

End Sub


'==========================================================================
'  ћ≈Ћќ„»
'==========================================================================
Private Function „итатьƒиапазон(rng As Range, ByVal nrow As Long, ByVal ncol As Long) As Variant
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
    „итатьƒиапазон = a
End Function

Private Sub —ообщить(ByVal текст As String)
    If Not mQuiet Then MsgBox текст, vbExclamation, "ќтчЄт из сводной"
End Sub

Private Function ¬—троку(v As Variant) As String
    On Error Resume Next
    If IsError(v) Then ¬—троку = "" Else ¬—троку = CStr(v)
    On Error GoTo 0
End Function
