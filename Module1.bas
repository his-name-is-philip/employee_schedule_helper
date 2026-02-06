Attribute VB_Name = "Module1"
' --------------------------------------------------------------
' Copyright (c) 2026, Филипп Сергеевич Соколов
' Все права защищены. Запрещено воспроизводство, модификация или
' распространение этого кода в коммерческих целяхбез письменной
' лицензии от владельца.
' Contact: hisnameisphilip@gmail.com
' --------------------------------------------------------------
' === Константы конфигурации ===
Private Const SETTINGS_SHEET As String = "Настройки"
Private Const BREAKS_SHEET As String = "Перерывы"
' номер столбца I (первый столбец с разметкой времени, он же старт графика)
Private Const FIRST_BREAK_COL As Long = 9
' номер столбца EO (конец графика)
Private Const LAST_COL As Long = FIRST_BREAK_COL + 136
' номер строки, с которой начинается поиск ФИО в столбце C
Private Const FIRST_SEARCH_ROW As Long = 12
' максимальное число перерывов, которое учитывается
Private Const MAX_BREAKS As Long = 8
' максимальное количество листов для обработки при sheetIndex = 0
Private Const MAX_SHEETS_TO_PROCESS As Long = 4

Private Sub Workbook_Open()
    Call RefreshMyBreaks
End Sub

Sub RefreshMyBreaks()
    Dim fName As String
    Dim wbSrc  As Workbook
    Dim wsSrc  As Worksheet
    Dim wsSet  As Worksheet
    Dim dataRow As Long
    
    Dim myName As String
    Dim filePattern As String
    Dim sourceFolder As String
    Dim breakText As String
    Dim t As Double
    Dim lastResCol As Long
    
    Dim myTimeZone As Double
    Dim tzMinutes As Long
    Dim sheetIndex As Long
    Dim i As Long
    Dim foundCount As Long
    Dim outputCol As Long
    Dim wsBreaks As Worksheet
    Dim hasAnyNonEmptyBreaks As Boolean

    ' 1) Читаем параметры из листа настроек
    If Not LoadSettings(myName, myTimeZone, filePattern, sheetIndex, sourceFolder, wsSet) Then
        Exit Sub
    End If

    ' 2) Ищем файл супервайзеров
    fName = ChooseFile(sourceFolder, filePattern)
    If fName = "" Then
        MsgBox "В папке " & sourceFolder & " не найден файл с '" & filePattern & "' в имени.", vbExclamation
        Exit Sub
    End If

    ' 3) Выводим имя файла
    Set wsBreaks = ThisWorkbook.Sheets(BREAKS_SHEET)
    wsBreaks.Range("B6").Value = Left(Dir(fName), InStrRev(Dir(fName), ".") - 1)

    ' 4) Открываем файл-источник
    Set wbSrc = Workbooks.Open(fName, ReadOnly:=True)

    ' 5) Считаем приращение в минутах для часового пояса
    tzMinutes = CLng(myTimeZone * 60)

    ' === Обработка всех листов при sheetIndex = 0 ===
    If sheetIndex = 0 Then
        foundCount = 0
        outputCol = 2 ' начинаем с колонки B
        hasAnyNonEmptyBreaks = False
        
        With wsBreaks
            lastResCol = 2 + 3 * MAX_SHEETS_TO_PROCESS
        
            ' Очищаем строки 3 и 4 в нужном диапазоне (от B до lastResCol)
            .Range(.Cells(3, 2), .Cells(4, lastResCol)).ClearContents
        
            ' Ставим всем столбцам такую же ширину, как у столбца №(lastResCol+1)
            .Range(.Columns(3), .Columns(lastResCol)).ColumnWidth = .Columns(lastResCol + 1).ColumnWidth
        End With

        
        ' Проходим по всем листам
        For i = 1 To wbSrc.Sheets.Count
            If foundCount >= MAX_SHEETS_TO_PROCESS Then Exit For
            
            Set wsSrc = wbSrc.Sheets(i)
            
            ' Поиск строки с ФИО на текущем листе
            dataRow = FindPerson(wsSrc, myName)
            
            If dataRow > 0 Then
                ' Формирование текста с перерывами
                breakText = BuildText(wsSrc, dataRow, tzMinutes)
                
                If breakText <> "" Then
                    ' Выводим пользователю результат
                    wsBreaks.Cells(3, outputCol).Value = breakText
                    ' Ставим ту же ширину, что у B
                    wsBreaks.Columns(outputCol).ColumnWidth = wsBreaks.Columns(2).ColumnWidth
                    ' Значение намеренно выводится только для sheetIndex = 0
                    wsBreaks.Cells(4, outputCol).Value = wsSrc.Name
                    ' Пишим в колонки B, E, H и т.д.
                    outputCol = outputCol + 3
                    hasAnyNonEmptyBreaks = True
                End If
                
                foundCount = foundCount + 1
            End If
        Next i
        
        ' Закрываем файл-источник
        wbSrc.Close False
        
        ' Проверка результатов
        If foundCount = 0 Then
            MsgBox "Не найдена строка с ФИО '" & myName & "' ни на одном листе.", vbExclamation
            Exit Sub
        End If
        
        If Not hasAnyNonEmptyBreaks Then
            MsgBox "Перерывов не найдено :(", vbInformation
        Else
            ' Активируем лист с перерывами
            On Error Resume Next
            wsBreaks.Activate
            On Error GoTo 0
            
            Application.StatusBar = "Перерывы обновлены. Надеюсь, сегодня у друзей такие же!"
            
            t = Timer + 3
            Do While Timer < t
                DoEvents
            Loop
            
            Application.StatusBar = False
            
            t = Timer + 0.01
            Do While Timer < t
                DoEvents
            Loop
            
            ThisWorkbook.Saved = True
        End If
        
    Else
        ' === Работа с одним конкретным листом ===
        
        ' Проверка корректности номера листа
        If sheetIndex > wbSrc.Sheets.Count Then
            MsgBox "Указан неверный номер листа: " & sheetIndex & ". В файле только " & _
                wbSrc.Sheets.Count & " лист(а/ов).", vbExclamation
            wbSrc.Close False
            Exit Sub
        End If
        
        Set wsSrc = wbSrc.Sheets(sheetIndex)

        ' 6) Поиск строки с ФИО (столбец C)
        dataRow = FindPerson(wsSrc, myName)
        If dataRow = 0 Then
            MsgBox "Не найдена строка с ФИО '" & myName & "'.", vbExclamation
            wbSrc.Close False
            Exit Sub
        End If

        ' 7) Формирование строк с интервалами перерывов
        breakText = BuildText(wsSrc, dataRow, tzMinutes)

        ' Закрываем файл-источник
        wbSrc.Close False

        If breakText = "" Then
            MsgBox "Перерывов не найдено :(", vbInformation
        Else
            wsBreaks.Range("B3").Value = breakText
            
            On Error Resume Next
            wsBreaks.Activate
            On Error GoTo 0
            
            Application.StatusBar = "Перерывы обновлены. Надеюсь, сегодня у друзей такие же!"

            t = Timer + 3
            Do While Timer < t
                DoEvents
            Loop
            
            Application.StatusBar = False
            
            t = Timer + 0.01
            Do While Timer < t
                DoEvents
            Loop
            
            ThisWorkbook.Saved = True
        End If
    End If
End Sub

' Функция возвращает номер строки с ФИО (dataRow)
Public Function FindPerson(ByVal wsSrc As Worksheet, ByVal myName As String) As Long
    Dim r As Long
    Dim lastRow As Long

    FindPerson = 0

    ' Сохраняем исходную
    lastRow = FIRST_SEARCH_ROW + wsSrc.Cells(wsSrc.Rows.Count, "C").End(xlUp).Row

    For r = FIRST_SEARCH_ROW To lastRow
        If Trim(wsSrc.Cells(r, "C").Value) = myName Then
            FindPerson = r
            Exit For
        End If
    Next r
End Function

' Функция возвращает готовый текст для ячейки или "", если cntBreaks = 0
Public Function BuildText(ByVal wsSrc As Worksheet, ByVal dataRow As Long, ByVal tzMinutes As Long) As String
    Dim col As Long
    Dim cntBreaks As Long
    Dim cellVal As String
    Dim baseCellValue As Variant
    Dim tStart As Date, tEnd As Date
    Dim breakText As String

    breakText = ""
    cntBreaks = 0
    col = FIRST_BREAK_COL

    Do While col <= LAST_COL And cntBreaks < MAX_BREAKS
        cellVal = Trim(wsSrc.Cells(dataRow, col).Value)

        ' Берём базовое значение времени из ячейки I1 (столбец FIRST_BREAK_COL)
        baseCellValue = wsSrc.Cells(1, FIRST_BREAK_COL).Value
        ' Коррекция базового времени (приведение к МСК) — вычитаем 4 часа
        baseCellValue = DateAdd("h", -4, wsSrc.Cells(1, 9).Value)

        Select Case cellVal
            Case "п"
                tStart = DateAdd("n", 15 * (col - FIRST_BREAK_COL) + tzMinutes, baseCellValue)
                tEnd = DateAdd("n", 15, tStart)
                cntBreaks = cntBreaks + 1
                breakText = breakText & Format(tStart, "HH:mm") & " - " & Format(tEnd, "HH:mm") & vbCrLf

            Case "п/10"
                tStart = DateAdd("n", 15 * (col - FIRST_BREAK_COL) + tzMinutes, baseCellValue)
                tEnd = DateAdd("n", 10, tStart)
                cntBreaks = cntBreaks + 1
                breakText = breakText & Format(tStart, "HH:mm") & " - " & Format(tEnd, "HH:mm") & vbCrLf

            Case "о"
                tStart = DateAdd("n", 15 * (col - FIRST_BREAK_COL) + tzMinutes, baseCellValue)
                tEnd = DateAdd("n", 30, tStart)
                cntBreaks = cntBreaks + 1
                breakText = breakText & Format(tStart, "HH:mm") & " - " & Format(tEnd, "HH:mm") & vbCrLf
                col = col + 1
        End Select

        col = col + 1
    Loop

    If cntBreaks = 0 Then
        BuildText = ""
        Exit Function
    End If

    ' Удалим последний перевод строки, если есть
    If Right(breakText, 2) = vbCrLf Then
        breakText = Left(breakText, Len(breakText) - 2)
    End If

    BuildText = breakText
End Function

' Загрузка настроек с листа Настройки
Private Function LoadSettings( _
    ByRef myName As String, _
    ByRef myTimeZone As Double, _
    ByRef filePattern As String, _
    ByRef sheetIndex As Long, _
    ByRef sourceFolder As String, _
    ByRef wsSet As Worksheet) _
As Boolean

    Dim sheetIndexCell As String

    On Error Resume Next
    Set wsSet = ThisWorkbook.Sheets(SETTINGS_SHEET)
    On Error GoTo 0

    If wsSet Is Nothing Then
        MsgBox "Не найден лист '" & SETTINGS_SHEET & "' с настройками.", vbExclamation
        LoadSettings = False
        Exit Function
    End If

    myName = Trim(wsSet.Range("C6").Value)
    myTimeZone = Val(wsSet.Range("C7").Value)
    filePattern = Trim(wsSet.Range("C8").Value)
    sheetIndexCell = wsSet.Range("C9").Value
    sourceFolder = Trim(wsSet.Range("C10").Value)
    
    
    ' Читаем номер листа из C9 (по умолчанию 1, если пусто или некорректно)
    If Trim(sheetIndexCell) = "" Then
        sheetIndex = 1
    Else
        sheetIndex = CLng(Val(sheetIndexCell))
        If sheetIndex < 0 Then sheetIndex = 1
    End If

    If myName = "" Or filePattern = "" Or sourceFolder = "" Then ' По умолчанию мск пояс
        MsgBox "Пожалуйста, заполните C6, C7, C8 и C10 на листе '" & SETTINGS_SHEET & "'.", vbExclamation
        LoadSettings = False
        Exit Function
    End If

    LoadSettings = True
End Function

' Ищем файл по маске в папке или выводим диалог выбора
Public Function ChooseFile(ByVal sourceFolder As String, filePattern As String) As String
    Dim baseFolder As String
    Dim fName As String
    Dim files As Collection
    Dim dlg As FileDialog
    Dim mask As String

    ' Нормализуем папку
    baseFolder = sourceFolder
    If Len(baseFolder) = 0 Then baseFolder = CurDir
    If Right$(baseFolder, 1) <> "\" And Right$(baseFolder, 1) <> "/" Then
        baseFolder = baseFolder & "\"
    End If

    ' Собираем файлы по маске в указанной папке
    Set files = New Collection
    fName = Dir(baseFolder & "*" & filePattern & "*.xls*")
    Do While Len(fName) > 0
        files.Add baseFolder & fName
        fName = Dir()
    Loop

    ' Нет файлов
    If files.Count = 0 Then
        ChooseFile = ""
        Set files = Nothing
        Exit Function
    End If

    ' Если найден только 1 - возвращаем сразу
    If files.Count = 1 Then
        ChooseFile = files(1)
        Set files = Nothing
        Exit Function
    End If

    ' Несколько вариантов - показываем FileDialog.
    mask = "*" & filePattern & "*.xls*"

    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .InitialFileName = baseFolder & mask ' папка + маска
        .Title = "Выберите файл"
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls;*.xlsx;*.xlsm"
        .Filters.Add "All files", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            ChooseFile = .SelectedItems(1)
        Else
            ChooseFile = ""
        End If
    End With

    Set dlg = Nothing
    Set files = Nothing
End Function
