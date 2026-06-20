# VBS скрипт для открытия двух файлов Excel

```vbscript
' ==========================================
' Открываем ДВА файла с АВТООБНОВЛЕНИЕМ связей
' ==========================================

Dim excelApp1, excelApp2, wb1, wb2

' ПУТИ К ФАЙЛАМ
Dim file1, file2
file1 = "C:.xlsx"
file2 = "C:\.xlsx"

' --- ПЕРВЫЙ ФАЙЛ ---
Set excelApp1 = CreateObject("Excel.Application")
excelApp1.Visible = True
excelApp1.AskToUpdateLinks = False
excelApp1.DisplayAlerts = False
Set wb1 = excelApp1.Workbooks.Open(file1, UpdateLinks:=3)

' --- ВТОРОЙ ФАЙЛ ---
Set excelApp2 = CreateObject("Excel.Application")
excelApp2.Visible = True
excelApp2.AskToUpdateLinks = False
excelApp2.DisplayAlerts = False
Set wb2 = excelApp2.Workbooks.Open(file2, UpdateLinks:=3)

' Если хотите сразу сохранить после обновления
wb1.Save
wb2.Save

MsgBox "Файлы открыты, связи обновлены!", vbInformation, "Готово!"
```
