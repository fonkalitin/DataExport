using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Packaging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PrPr_exportDataToDoc
{
    public class CreateXlDataTable
    {
        /// <summary>
        /// /// Что делает код:
        /// 1. Создает новый файл `.xlsx` с использованием библиотеки **OpenXML**.
        /// 2. Использует переданный двумерный массив для генерации строк и столбцов в таблице.
        /// 3. Учитывает стартовый адрес ячейки (например, `B2`) при записи.
        /// 4. Генерирует строки (`Row`) и ячейки(`Cell`) построчно из массива.
        /// 5. Сохраняет файл в указанном каталоге с заданным именем.
        /// </summary>
        /// <param name="data">Двумерный массив строк для записи в таблицу</param>
        /// <param name="fileName">Имя создаваемого файла (с расширением .xlsx)</param>
        /// <param name="sheetName">Имя листа таблицы</param>
        /// <param name="startCellAddress">Адрес начальной ячейки для записи (например, "A1")</param>
        public void CreateXlsxFromData(string[,] data, string fileName, string sheetName, string startCellAddress)
        {

 
            // Создаем новый xlsx файл
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // Создаем части для файла: Workbook и Worksheet
                WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Создаем коллекцию листов и добавляем туда наш лист
                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);

                // Получаем SheetData, куда мы будем добавлять строки и ячейки
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Конвертация начального адреса ячейки (например, "B2")
                int startRow = int.Parse(startCellAddress.Substring(1)); // Извлекаем номер строки
                int startColumn = GetColumnIndexFromLetter(startCellAddress[0].ToString()); // Преобразуем букву в индекс

                // Построчно обходим массив данных
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    Row newRow = new Row() { RowIndex = (uint)(startRow + i) }; // Устанавливаем индекс строки

                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        Cell newCell = new Cell()
                        {
                            CellReference = GetCellAddress(startRow + i, startColumn + j), // Адрес ячейки
                            DataType = CellValues.String,                                  // Тип данных - строка
                            CellValue = new CellValue(data[i, j])                          // Значение ячейки
                        };
                        newRow.Append(newCell); // Добавляем ячейку в строку
                    }

                    sheetData.Append(newRow); // Добавляем строку в лист данных
                }

                // Сохраняем изменения
                workbookPart.Workbook.Save();
            }
        }

        /// <summary>
        /// Преобразует индекс строки (число) и индекс колонки (число) в адрес ячейки (например, "A1")
        /// </summary>
        private string GetCellAddress(int rowIndex, int columnIndex)
        {
            string columnLetter = GetColumnLetterFromIndex(columnIndex);
            return $"{columnLetter}{rowIndex}";
        }

        /// <summary>
        /// Преобразует индекс колонки (число) в букву колонки (например, 1 -> "A", 2 -> "B")
        /// </summary>
        private string GetColumnLetterFromIndex(int columnIndex)
        {
            string columnLetter = string.Empty;
            while (columnIndex > 0)
            {
                columnIndex--;
                columnLetter = (char)('A' + columnIndex % 26) + columnLetter;
                columnIndex /= 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// Преобразует букву колонки (например, "A", "B", "AA") в индекс (например, "A" -> 1, "B" -> 2)
        /// </summary>
        private int GetColumnIndexFromLetter(string columnLetter)
        {
            int index = 0;
            foreach (char c in columnLetter)
            {
                index *= 26;
                index += (c - 'A' + 1);
            }
            return index;
        }

    }
}
