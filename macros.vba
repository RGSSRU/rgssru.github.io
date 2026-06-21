Option Explicit
Option Base 1

' ==========================================
' НАСТРОЙКИ (меняете под себя!)
' ==========================================

' Папка с файлами для слепка (внутри будет папка с датой)
Const SOURCE_FOLDER As String = "C:\Data\"

' Папка с целевыми файлами
Const TARGET_FOLDER As String = "C:\Data\Targets\"

' Целевые файлы (4 штуки)
Const FILE1 As String = ".xlsx"
Const FILE2 As String = ".xlsx"
Const FILE3 As String = ".xlsx"
Const FILE4 As String = ".xlsx"

' Столбцы для копирования (буквы)
' Для файлов 2020-2024
Const COLUMNS_2020_2024 As String = "O,W-AL,AO-BO,BY,CH"
' Для файлов 2025-2026
Const COLUMNS_2025_2026 As String = "H,Q-AE,AG-AY,BA"

' ==========================================
' ГЛАВНЫЙ МАКРОС
' ==========================================

Sub ProcessFiles()
    Dim startTime As Double
    startTime = Timer
    
    ' Отключаем всё лишнее для скорости
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' ШАГ 1: Создаем СЛЕПОК в памяти
    ' ==========================================
    Dim fso As Object
    Dim folderPath As String
    Dim folder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim allFilesData As Collection
    Dim fileData As Variant
    Dim totalRows As Long
    Dim maxCols As Long
    Dim snapshotData As Variant
    Dim currentRow As Long
    Dim i As Long, j As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set allFilesData = New Collection
    
    ' Формируем путь к папке с текущей датой
    folderPath = SOURCE_FOLDER & Format(Date, "YYYY-MM-DD") & "\"
    
    ' Проверяем, существует ли папка
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Папка не найдена: " & folderPath, vbCritical, "Ошибка"
        GoTo CleanUp
    End If
    
    Set folder = fso.GetFolder(folderPath)
    
    ' Читаем ВСЕ xlsx файлы из папки
    For Each file In folder.Files
        If LCase(Right(file.Name, 5)) = ".xlsx" Then
            ' Открываем файл только для чтения (быстро)
            Set wb = Workbooks.Open(file.Path, ReadOnly:=True, UpdateLinks:=False)
            Set ws = wb.Sheets(1)
            
            ' Читаем данные в массив (ОДНА операция!)
            fileData = ws.UsedRange.Value
            
            ' Сохраняем в коллекцию
            allFilesData.Add fileData
            
            ' Считаем строки и столбцы
            totalRows = totalRows + UBound(fileData, 1)
            If UBound(fileData, 2) > maxCols Then maxCols = UBound(fileData, 2)
            
            ' Закрываем файл
            wb.Close SaveChanges:=False
        End If
    Next file
    
    ' Проверяем, есть ли данные
    If allFilesData.Count = 0 Then
        MsgBox "В папке нет xlsx файлов!", vbExclamation, "Предупреждение"
        GoTo CleanUp
    End If
    
    ' Объединяем все файлы в один массив (СЛЕПОК)
    ReDim snapshotData(1 To totalRows, 1 To maxCols)
    currentRow = 1
    
    For Each fileData In allFilesData
        For i = 1 To UBound(fileData, 1)
            For j = 1 To UBound(fileData, 2)
                snapshotData(currentRow, j) = fileData(i, j)
            Next j
            currentRow = currentRow + 1
        Next i
    Next fileData
    
    ' Освобождаем память
    Set allFilesData = Nothing
    
    ' ==========================================
    ' ШАГ 2: Определяем уникальные годы из слепка
    ' ==========================================
    Dim yearsDict As Object
    Dim yearStr As String
    
    Set yearsDict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(snapshotData, 1)
        If Not IsEmpty(snapshotData(i, 1)) Then
            yearStr = ExtractYear(CStr(snapshotData(i, 1)))
            If yearStr <> "" Then
                If Not yearsDict.Exists(yearStr) Then
                    yearsDict.Add yearStr, True
                End If
            End If
        End If
    Next i
    
    ' ==========================================
    ' ШАГ 3: Формируем список файлов для открытия
    ' ==========================================
    Dim fileList As Collection
    Dim targetFiles As Variant
    Dim targetFile As Variant
    Dim targetFilePath As String
    Dim fileYears As Collection
    Dim needOpen As Boolean
    Dim yearKey As Variant
    
    Set fileList = New Collection
    targetFiles = Array(FILE1, FILE2, FILE3, FILE4)
    
    For Each targetFile In targetFiles
        targetFilePath = TARGET_FOLDER & targetFile
        
        If fso.FileExists(targetFilePath) Then
            Set fileYears = GetYearsFromFileName(targetFile)
            needOpen = False
            
            For Each yearKey In fileYears
                If yearsDict.Exists(yearKey) Then
                    needOpen = True
                    Exit For
                End If
            Next yearKey
            
            If needOpen Then
                fileList.Add targetFilePath
            End If
        End If
    Next targetFile
    
    ' Проверяем, есть ли файлы для обработки
    If fileList.Count = 0 Then
        MsgBox "Нет файлов для обработки!", vbInformation, "Готово"
        GoTo CleanUp
    End If
    
    ' ==========================================
    ' ШАГ 4: Создаем индекс для слепка (для быстрого поиска)
    ' ==========================================
    Dim snapshotIndex As Object
    Dim prefix As String
    
    Set snapshotIndex = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(snapshotData, 1)
        If Not IsEmpty(snapshotData(i, 1)) Then
            prefix = GetPrefix(CStr(snapshotData(i, 1)))
            If prefix <> "" Then
                If Not snapshotIndex.Exists(prefix) Then
                    snapshotIndex.Add prefix, New Collection
                End If
                snapshotIndex(prefix).Add i
            End If
        End If
    Next i
    
    ' ==========================================
    ' ШАГ 5: Сохраняем исходное состояние файлов
    ' ==========================================
    Dim originalData As Object
    Set originalData = CreateObject("Scripting.Dictionary")
    
    ' ==========================================
    ' ШАГ 6: Обрабатываем КАЖДЫЙ файл ПО ОЧЕРЕДИ
    ' ==========================================
    Dim processedFiles As Long
    Dim totalAdded As Long
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetData As Variant
    Dim targetRows As Long
    Dim targetCols As Long
    Dim targetIndex As Object
    Dim rowIdx As Variant
    Dim suff As Long
    Dim colsToCopy As Collection
    Dim col As Variant
    Dim sourceRow As Long
    Dim maxSuffix As Long
    Dim currentSuffix As Long
    Dim cellValue As String
    Dim rowsToAdd As Long
    Dim s As Long
    Dim newIndex As Long
    
    processedFiles = 0
    totalAdded = 0
    
    For Each targetFilePath In fileList
        ' --- Открываем файл для редактирования ---
        Set targetWb = Workbooks.Open(targetFilePath)
        Set targetWs = targetWb.Sheets(1)
        
        ' --- Читаем данные в массив (ОДНА операция!) ---
        targetData = targetWs.UsedRange.Value
        targetRows = UBound(targetData, 1)
        targetCols = UBound(targetData, 2)
        
        ' --- Сохраняем исходные данные для сравнения ---
        originalData.Add targetFilePath, targetData
        
        ' --- Создаем индекс для целевого файла ---
        Set targetIndex = CreateObject("Scripting.Dictionary")
        
        For i = 1 To targetRows
            If Not IsEmpty(targetData(i, 1)) Then
                prefix = GetPrefix(CStr(targetData(i, 1)))
                If prefix <> "" Then
                    If Not targetIndex.Exists(prefix) Then
                        targetIndex.Add prefix, New Collection
                    End If
                    targetIndex(prefix).Add i
                End If
            End If
        Next i
        
        ' --- Определяем столбцы для копирования ---
        If InStr(targetFilePath, "2020") > 0 Or _
           InStr(targetFilePath, "2022") > 0 Or _
           InStr(targetFilePath, "2023") > 0 Or _
           InStr(targetFilePath, "2024") > 0 Then
            Set colsToCopy = ParseColumns(COLUMNS_2020_2024)
        Else
            Set colsToCopy = ParseColumns(COLUMNS_2025_2026)
        End If
        
        ' ==========================================
        ' НОВАЯ ЛОГИКА: КОПИРУЕМ И ВСТАВЛЯЕМ СТРОКИ
        ' ==========================================
        
        ' Проходим по КАЖДОМУ префиксу в целевом файле
        For Each prefix In targetIndex.Keys
            
            ' Проверяем, есть ли такой префикс в слепке
            If snapshotIndex.Exists(prefix) Then
                
                ' --- Находим максимальный индекс в слепке ---
                maxSuffix = 0
                For Each rowIdx In snapshotIndex(prefix)
                    suff = GetSuffix(CStr(snapshotData(rowIdx, 1)))
                    If suff > maxSuffix Then maxSuffix = suff
                Next rowIdx
                
                ' --- Находим текущий индекс в целевом файле ---
                currentSuffix = 0
                sourceRow = 0
                
                For Each rowIdx In targetIndex(prefix)
                    cellValue = CStr(targetData(rowIdx, 1))
                    
                    If InStr(cellValue, "/") > 0 Then
                        currentSuffix = GetSuffix(cellValue)
                        sourceRow = rowIdx
                        Exit For
                    Else
                        ' Если нет индекса, берем как есть
                        currentSuffix = 0
                        sourceRow = rowIdx
                        Exit For
                    End If
                Next rowIdx
                
                ' --- Проверяем, есть ли строка без индекса ---
                Dim hasWithoutIndex As Boolean
                hasWithoutIndex = False
                
                For Each rowIdx In targetIndex(prefix)
                    cellValue = CStr(targetData(rowIdx, 1))
                    If InStr(cellValue, "/") = 0 Then
                        hasWithoutIndex = True
                        Exit For
                    End If
                Next rowIdx
                
                ' ==========================================
                ' СЛУЧАЙ 1: Есть строка без индекса
                ' ==========================================
                If hasWithoutIndex Then
                    ' Подсчитываем количество индексов в слепке
                    Dim indexCount As Long
                    indexCount = snapshotIndex(prefix).Count
                    
                    ' Если в слепке 2+ индекса
                    If indexCount >= 2 Then
                        ' Удаляем строку без индекса
                        For Each rowIdx In targetIndex(prefix)
                            cellValue = CStr(targetData(rowIdx, 1))
                            If InStr(cellValue, "/") = 0 Then
                                targetWs.Rows(rowIdx).Delete
                                Exit For
                            End If
                        Next rowIdx
                        
                        ' Обновляем targetIndex после удаления
                        Set targetIndex = Nothing
                        Set targetIndex = CreateObject("Scripting.Dictionary")
                        
                        ' Перестраиваем индекс
                        targetData = targetWs.UsedRange.Value
                        targetRows = UBound(targetData, 1)
                        
                        For i = 1 To targetRows
                            If Not IsEmpty(targetData(i, 1)) Then
                                prefix = GetPrefix(CStr(targetData(i, 1)))
                                If prefix <> "" Then
                                    If Not targetIndex.Exists(prefix) Then
                                        targetIndex.Add prefix, New Collection
                                    End If
                                    targetIndex(prefix).Add i
                                End If
                            End If
                        Next i
                        
                        ' Теперь добавляем все строки из слепка через копирование
                        ' Находим первую строку с этим префиксом в целевом файле
                        sourceRow = 0
                        For Each rowIdx In targetIndex(prefix)
                            sourceRow = rowIdx
                            Exit For
                        Next rowIdx
                        
                        If sourceRow > 0 Then
                            ' Добавляем все строки из слепка
                            For s = 1 To maxSuffix
                                ' Проверяем, есть ли уже такой индекс
                                Dim exists As Boolean
                                exists = False
                                For Each rowIdx In targetIndex(prefix)
                                    If GetSuffix(CStr(targetData(rowIdx, 1))) = s Then
                                        exists = True
                                        Exit For
                                    End If
                                Next rowIdx
                                
                                If Not exists Then
                                    ' Копируем строку
                                    targetWs.Rows(sourceRow).Copy
                                    targetWs.Rows(sourceRow + 1).Insert Shift:=xlDown
                                    
                                    ' Меняем индекс
                                    targetWs.Cells(sourceRow + 1, 1).Value = prefix & "/" & s
                                    
                                    ' Обновляем данные из слепка
                                    For Each rowIdx In snapshotIndex(prefix)
                                        If GetSuffix(CStr(snapshotData(rowIdx, 1))) = s Then
                                            For Each col In colsToCopy
                                                If col <= targetCols And col <= UBound(snapshotData, 2) Then
                                                    targetWs.Cells(sourceRow + 1, col).Value = snapshotData(rowIdx, col)
                                                End If
                                            Next col
                                            Exit For
                                        End If
                                    Next rowIdx
                                    
                                    sourceRow = sourceRow + 1
                                    totalAdded = totalAdded + 1
                                End If
                            Next s
                        End If
                    End If
                    ' Если indexCount = 1 → ничего не делаем
                    
                Else
                    ' ==========================================
                    ' СЛУЧАЙ 2: Нет строки без индекса
                    ' ==========================================
                    
                    ' Если нужно добавить строки
                    If maxSuffix > currentSuffix Then
                        ' Определяем, сколько строк нужно добавить
                        rowsToAdd = maxSuffix - currentSuffix
                        
                        ' Для каждой новой строки
                        For s = 1 To rowsToAdd
                            ' --- 1. КОПИРУЕМ строку-источник ---
                            targetWs.Rows(sourceRow).Copy
                            
                            ' --- 2. ВСТАВЛЯЕМ копию НИЖЕ ---
                            targetWs.Rows(sourceRow + 1).Insert Shift:=xlDown
                            
                            ' --- 3. МЕНЯЕМ индекс в скопированной строке ---
                            newIndex = currentSuffix + s
                            targetWs.Cells(sourceRow + 1, 1).Value = prefix & "/" & newIndex
                            
                            ' --- 4. ОБНОВЛЯЕМ данные из слепка ---
                            For Each rowIdx In snapshotIndex(prefix)
                                If GetSuffix(CStr(snapshotData(rowIdx, 1))) = newIndex Then
                                    For Each col In colsToCopy
                                        If col <= targetCols And col <= UBound(snapshotData, 2) Then
                                            targetWs.Cells(sourceRow + 1, col).Value = snapshotData(rowIdx, col)
                                        End If
                                    Next col
                                    Exit For
                                End If
                            Next rowIdx
                            
                            ' Сдвигаем sourceRow для следующей вставки
                            sourceRow = sourceRow + 1
                            totalAdded = totalAdded + 1
                            
                        Next s
                    End If
                End If
            End If
        Next prefix
        
        ' --- Сохраняем изменения ---
        If totalAdded > 0 Then
            targetWb.Save
            processedFiles = processedFiles + 1
        Else
            targetWb.Close SaveChanges:=False
        End If
    Next targetFilePath
    
    ' ==========================================
    ' ШАГ 7: Сравниваем массивы и выводим статистику
    ' ==========================================
    Call CompareAndShowStatistics(originalData, fileList)
    
    ' ==========================================
    ' РЕЗУЛЬТАТ
    ' ==========================================
    MsgBox "ГОТОВО!" & vbCrLf & _
           "Обработано файлов: " & processedFiles & vbCrLf & _
           "Добавлено строк: " & totalAdded & vbCrLf & _
           "Время выполнения: " & Format(Timer - startTime, "0.00") & " сек.", _
           vbInformation, "Успешно!"
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка!"

CleanUp:
    ' Включаем всё обратно
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Освобождаем память
    Set fso = Nothing
    Set snapshotData = Nothing
    Set snapshotIndex = Nothing
    Set yearsDict = Nothing
    Set fileList = Nothing
    Set originalData = Nothing
End Sub

' ==========================================
' ФУНКЦИЯ ДЛЯ СРАВНЕНИЯ И СТАТИСТИКИ
' ==========================================

Sub CompareAndShowStatistics(originalData As Object, fileList As Collection)
    Dim filePath As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentData As Variant
    Dim originalArray As Variant
    Dim i As Long, j As Long
    Dim row As Long, col As Long
    Dim totalChanges As Long
    Dim changesByFile As String
    Dim maxRows As Long
    Dim maxCols As Long
    
    totalChanges = 0
    changesByFile = ""
    
    For Each filePath In fileList
        ' Открываем файл для чтения
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        Set ws = wb.Sheets(1)
        
        ' Получаем текущие данные
        currentData = ws.UsedRange.Value
        
        ' Получаем исходные данные из словаря
        If originalData.Exists(filePath) Then
            originalArray = originalData(filePath)
        Else
            GoTo NextFile
        End If
        
        ' Определяем максимальные размеры для сравнения
        maxRows = UBound(currentData, 1)
        If UBound(originalArray, 1) > maxRows Then maxRows = UBound(originalArray, 1)
        
        maxCols = UBound(currentData, 2)
        If UBound(originalArray, 2) > maxCols Then maxCols = UBound(originalArray, 2)
        
        ' Сравниваем ячейки
        Dim fileChanges As Long
        fileChanges = 0
        
        For i = 1 To maxRows
            For j = 1 To maxCols
                Dim val1 As Variant
                Dim val2 As Variant
                
                ' Получаем значения из массивов (если есть)
                If i <= UBound(originalArray, 1) And j <= UBound(originalArray, 2) Then
                    val1 = originalArray(i, j)
                Else
                    val1 = Empty
                End If
                
                If i <= UBound(currentData, 1) And j <= UBound(currentData, 2) Then
                    val2 = currentData(i, j)
                Else
                    val2 = Empty
                End If
                
                ' Сравниваем (игнорируем пустые ячейки)
                If Not (IsEmpty(val1) And IsEmpty(val2)) Then
                    If IsEmpty(val1) And Not IsEmpty(val2) Then
                        fileChanges = fileChanges + 1
                    ElseIf Not IsEmpty(val1) And IsEmpty(val2) Then
                        fileChanges = fileChanges + 1
                    ElseIf val1 <> val2 Then
                        fileChanges = fileChanges + 1
                    End If
                End If
            Next j
        Next i
        
        totalChanges = totalChanges + fileChanges
        
        ' Формируем статистику по файлу
        Dim fileName As String
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        changesByFile = changesByFile & fileName & ": " & fileChanges & " изменений" & vbCrLf
        
NextFile:
        wb.Close SaveChanges:=False
    Next filePath
    
    ' Показываем статистику
    Dim statsMsg As String
    statsMsg = "СТАТИСТИКА ИЗМЕНЕНИЙ:" & vbCrLf & vbCrLf
    statsMsg = statsMsg & changesByFile & vbCrLf
    statsMsg = statsMsg & "ВСЕГО ИЗМЕНЕНИЙ: " & totalChanges & vbCrLf
    
    If totalChanges > 0 Then
        statsMsg = statsMsg & vbCrLf & "✔ Изменения успешно применены"
    Else
        statsMsg = statsMsg & vbCrLf & "ℹ Изменений не обнаружено"
    End If
    
    MsgBox statsMsg, vbInformation, "Статистика изменений"
End Sub

' ==========================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ==========================================

' --- Извлечение года из строки ---
Function ExtractYear(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b\d{4}\b"
    
    If regex.Test(text) Then
        ExtractYear = regex.Execute(text)(0).Value
    Else
        ExtractYear = ""
    End If
End Function

' --- Получение части до "/" ---
Function GetPrefix(text As String) As String
    If InStr(text, "/") > 0 Then
        GetPrefix = Left(text, InStr(text, "/") - 1)
    Else
        GetPrefix = text
    End If
End Function

' --- Получение числа после "/" ---
Function GetSuffix(text As String) As Long
    If InStr(text, "/") > 0 Then
        Dim suffixStr As String
        suffixStr = Mid(text, InStr(text, "/") + 1)
        If IsNumeric(suffixStr) Then
            GetSuffix = CLng(suffixStr)
        Else
            GetSuffix = 0
        End If
    Else
        GetSuffix = 0
    End If
End Function

' --- Извлечение годов из названия файла ---
Function GetYearsFromFileName(fileName As String) As Collection
    Dim result As New Collection
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d{4}"
    regex.Global = True
    
    Set matches = regex.Execute(fileName)
    
    For Each match In matches
        result.Add match.Value
    Next match
    
    Set GetYearsFromFileName = result
End Function

' --- Парсинг столбцов (например, "O,W-AL,AO-BO,BY,CH") ---
Function ParseColumns(colRange As String) As Collection
    Dim result As New Collection
    Dim parts As Variant
    Dim part As Variant
    Dim rangeParts As Variant
    Dim startCol As Long, endCol As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    parts = Split(colRange, ",")
    For Each part In parts
        part = Trim(part)
        If InStr(part, "-") > 0 Then
            ' Диапазон (например, "W-AL")
            rangeParts = Split(part, "-")
            startCol = ColumnLetterToNumber(rangeParts(0))
            endCol = ColumnLetterToNumber(rangeParts(1))
            For i = startCol To endCol
                result.Add i
            Next i
        Else
            ' Одиночный столбец (например, "O")
            result.Add ColumnLetterToNumber(part)
        End If
    Next part
    
    Set ParseColumns = result
    Exit Function
    
ErrorHandler:
    MsgBox "Ошибка в константе столбцов: " & colRange & vbCrLf & Err.Description, vbCritical
    Set ParseColumns = New Collection
End Function

' --- Преобразование буквы столбца в номер ---
Function ColumnLetterToNumber(colLetter As String) As Long
    Dim result As Long
    Dim i As Long
    Dim letter As String
    Dim text As String
    
    text = UCase(Trim(colLetter))
    result = 0
    For i = 1 To Len(text)
        letter = Mid(text, i, 1)
        result = result * 26 + (Asc(letter) - 64)
    Next i
    
    ColumnLetterToNumber = result
End Function
                                         
