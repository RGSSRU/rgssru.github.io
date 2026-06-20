
' ==========================================
' Открываем два Excel файла
' ==========================================

Dim excelApp1, excelApp2

' --- ПУТИ К ФАЙЛАМ (ЗАМЕНИТЕ НА СВОИ!) ---
Dim file1, file2
file1 = "C:\YourFolder\.xlsx"
file2 = "C:\YourFolder\.xlsx"

' --- ОТКРЫВАЕМ ПЕРВЫЙ ФАЙЛ ---
Set excelApp1 = CreateObject("Excel.Application")
excelApp1.Visible = True
excelApp1.Workbooks.Open file1

' --- ОТКРЫВАЕМ ВТОРОЙ ФАЙЛ ---
Set excelApp2 = CreateObject("Excel.Application")
excelApp2.Visible = True
excelApp2.Workbooks.Open file2

' --- СООБЩЕНИЕ ---
MsgBox "Оба файла открыты!", vbInformation, "Готово!"
