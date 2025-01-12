using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using wordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using wordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using wordRun = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace PrPr_exportDataToDoc_WinApp
{
    /// <summary>
    /// Класс копирует данные из xlsx файла в docx.
    /// </summary>

    public class copyXlDataTableToDoc
    {
        /// <param name="xlsxFileName">Имя xlsx файла с исходными данными</param>
        /// <param name="sheetName">Имя листа о старта таблицы</param>
        /// <param name="startCell">Номер стартовой ячейки (например, "A1")</param>
        /// <param name="templateDirectory">Каталог с документом-шаблоном docx</param>
        /// <param name="templateDocxName">Имя документа-шаблона docx</param>
        /// <param name="newDocxName">Имя нового файла docx</param>
        /// <param name="tableIndex">Номер таблицы в документе docx</param>
        /// <param name="startRowInDocx">Номер первой строки для вставки данных</param> 
        public void Execute(
            string xlsxFileName,       // Имя xlsx файла с исходными данными
            string sheetName,          // Имя листа о старта таблицы
            string startCell,          // Номер стартовой ячейки (например, "A1")
            string templateDirectory,  // Каталог с документом-шаблоном docx
            string templateDocxName,   // Имя документа-шаблона docx
            string newDocxName,        // Имя нового файла docx
            int tableIndex,            // Номер таблицы в документе docx
            int startRowInDocx         // Номер первой строки для вставки данных
        )

        {
            // Получение полного пути к исходному Excel-файлу и Word-шаблону
            string xlsxFilePath = xlsxFileName;
            string templateDocxPath = Path.Combine(templateDirectory, templateDocxName);
            string newDocxPath = newDocxName;

            try
            {
                // 1. Чтение данных из Excel
                Console.WriteLine("Чтение данных из Excel-файла...");
                var excelData = ReadExcelData(xlsxFilePath, sheetName, startCell);

                // 2. Открытие шаблона Word и запись данных
                Console.WriteLine("Запись данных в Word-документ...");
                WriteDataToDocx(templateDocxPath, newDocxPath, tableIndex, startRowInDocx, excelData);

                // 3. Уведомление об успехе
                Console.WriteLine($"Новый документ сохранен: {newDocxPath}");
            }
            catch (Exception ex)
            {
                // Обработка ошибок
                Console.WriteLine($"Ошибка: {ex.Message}");
            }
        }

        // Чтение данных из Excel
        private string[,] ReadExcelData(string filePath, string sheetName, string startCell)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                // Находим указанный лист
                Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);
                if (sheet == null)
                    throw new Exception($"Лист с именем '{sheetName}' не найден.");

                // Получаем часть листа
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Примерная организационная обработка: немного захардкодим размеры
                List<List<string>> rows = new();
                foreach (Row row in sheetData.Elements<Row>())
                {
                    List<string> cellValues = new();
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(document, cell);
                        cellValues.Add(cellValue);
                    }
                    rows.Add(cellValues);
                }

                // Преобразуем List<List<string>> в двумерный массив
                int maxRowLength = rows.Max(r => r.Count);
                string[,] result = new string[rows.Count, maxRowLength];

                for (int i = 0; i < rows.Count; i++)
                {
                    for (int j = 0; j < rows[i].Count; j++)
                    {
                        result[i, j] = rows[i][j];
                    }
                }
                return result;
            }
        }

        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            // Получаем строковое значение ячейки
            if (cell == null || cell.CellValue == null) return string.Empty;

            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return document.WorkbookPart
                    .SharedStringTablePart
                    .SharedStringTable
                    .ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }

        // Запись данных в документ Word
        private void WriteDataToDocx(
            string templateDocxPath,
            string newDocxPath,
            int tableIndex,
            int startRowInDocx,
            string[,] excelData)
        {
            // Копируем шаблон в новый файл, чтобы работать с ним
            File.Copy(templateDocxPath, newDocxPath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(newDocxPath, true))
            {
                // Получаем таблицу из документа
                wordTable table = doc.MainDocumentPart.Document.Body.Elements<wordTable>().ElementAtOrDefault(tableIndex);
                if (table == null)
                    throw new Exception($"Таблица с индексом {tableIndex} не найдена в документе.");

                // Вставляем данные из Excel в таблицу
                for (int i = 0; i < excelData.GetLength(0); i++)
                {
                    TableRow row = new TableRow();

                    for (int j = 0; j < excelData.GetLength(1); j++)
                    {
                        TableCell cell = new TableCell(new Paragraph(new wordRun(new wordText(excelData[i, j]))));
                        row.Append(cell);
                    }

                    // Вставляем строку в таблицу после указанной стартовой строки
                    table.Append(row);
                }

                // Сохраняем изменения
                doc.MainDocumentPart.Document.Save();

                MessageBox.Show("Файл docx успешно создан, данные скопированы!", "ok");
            }
        }


    }
}
