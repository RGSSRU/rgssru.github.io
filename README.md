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
    Set
