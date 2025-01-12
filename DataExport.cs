
using App = HostMgd.ApplicationServices;
    using Db = Teigha.DatabaseServices;
    using Ed = HostMgd.EditorInput;
    using Rtm = Teigha.Runtime;
using System.Diagnostics;
using PRPR_METHODS;
using System.Linq.Expressions;
using Multicad;
using System.Collections.Generic;
using PrPr_exportDataToDoc_WinApp;

[assembly: Rtm.CommandClass(typeof(Tools.CadCommand))]

namespace Tools
    {
        /// <summary> 
        /// Комманды
        /// </summary>
        class CadCommand : Rtm.IExtensionApplication
        {

            #region INIT
            public void Initialize()
            {
                //think добавить проверку есть ли doc

                App.DocumentCollection dm = App.Application.DocumentManager;


                Ed.Editor ed = dm.MdiActiveDocument.Editor;
                    //ed.WriteMessage("\nStart list of commands: \n");
                string sCom =
                    "dataexport" + "\tэкспорт данных из объектов nanoCAD";
                ed.WriteMessage(sCom);

#if DEBUG
                //для отладки список команд
#endif
            }

            public void Terminate()
            {
                // throw new System.NotImplementedException();
            }

            #endregion


            /// <summary>
            /// запуск ссылок
            /// </summary>
            [Rtm.CommandMethod("dataexport", Rtm.CommandFlags.Session)]
            public void dataexport()
            {
                Db.Database db = Db.HostApplicationServices.WorkingDatabase;
                App.Document doc = App.Application.DocumentManager.MdiActiveDocument;
                Ed.Editor ed = doc.Editor;

            // ---- Имена и пути оригинального dwg-файла (открытого)
            string dwgName = doc.Name; // метод получения полного пути и имени текущего dwg-файла (db.Filename; // Альтернативный метод)
            string dwgFileDirPath = Path.GetDirectoryName(dwgName); // Путь до каталога dwg файла (без имени файла) 

            List<string> atribList = new List<string>();
                // Заполняем данными
                atribList.Add("pos_tag");
                atribList.Add("media_name");
                atribList.Add("media_wtemp");
                atribList.Add("media_wpress");

                List<List<ExValue>> atribDataList = new List<List<ExValue>>();
                atribDataList = PRPR_METHODS.PRPR_METHODS.CollectParObjData("Позиция трубопровода/оборудования v0.6", atribList, 1);

            // Преобразуем данные в таблицу
            string[,] dataTable = GetTable(atribDataList);

            string fullFileName_xlsx = Path.Combine(dwgFileDirPath, "ТИП.xlsx");
            string fullFileName_docx_tmp = Path.Combine(dwgFileDirPath, "шаблоны");
            string fullFileName_docx = Path.Combine(dwgFileDirPath, "Номер-РД.ТИП.docx");
            string templatePath = Path.Combine(dwgFileDirPath, "шаблоны");

            // Создание xlsx файла с данными по выбранным объектам
            var xlTableCreator = new CreateXlDataTable();
            xlTableCreator.CreateXlsxFromData(dataTable, fullFileName_xlsx, "Лист1", "B2");
            MessageBox.Show("Файл xlsx успешно создан!", "ok");

            copyXlDataTableToDoc tableDataCopy = new copyXlDataTableToDoc();
            

            // Аргументы вызова метода
            tableDataCopy.Execute(
                fullFileName_xlsx,               // Имя файла Excel
                "Лист1",                        // Имя листа
                "B2",                           // Начальная ячейка
                templatePath,                  // Каталог шаблонов
                "Шаблон.ТИП.docx",            // Имя файла документа-шаблона
                fullFileName_docx,          // Имя нового документа
                1,                            // Номер таблицы в документе
                2                             // Номер стартовой строки
            );

        }


        // Метод для преобразования списка объектов и их параметров в двумерный массив строк
        public static string[,] GetTable(List<List<ExValue>> atribDataList)
        {
            int rows = atribDataList.Count; // Количество строк равно количеству объектов
            int cols = 4; // Количество столбцов (параметров)

            // Инициализируем двумерный массив для хранения таблицы
            string[,] table = new string[rows, cols];

            // Заполняем таблицу значениями параметров для каждого объекта
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    // Получаем строковое представление значения параметра
                    table[i, j] = atribDataList[i][j].AsString;
                }
            }

            return table; // Возвращаем заполненную таблицу
        }
    }

}


