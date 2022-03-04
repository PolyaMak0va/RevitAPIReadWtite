using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPIReadWtite
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // загрузим библиотеку NPOI для работы с файлами Excel: нажимаем правой кнопкой мыши на "Ссылки" в "Обозревателе решений"; в браузере ищем NPOI, скачиваем и устанавливаем
            // записываем данные в файл Excel
            string roomInfo = string.Empty;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            // определим путь, по кот-му будет сохраняться файл Excel
            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");

            // создаём файл
            // excelPath - путь к нахождению файла
            // FileMode.Create - создаём файл
            // FileAccess.Write - записываем в него данные
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                // виртуально создали книгу, файл Excel
                IWorkbook workbook = new XSSFWorkbook();

                // создадим там же лист
                ISheet sheet = workbook.CreateSheet("Лист1");

                // заполнить данными файл Excel, но перед этим надо создать метод (SheetExts), кот. позволит это делать и поозволит делать код читаемым
                int rowIndex = 0;
                foreach (var room in rooms) // проходимя по каждому помещению, кот-е есть в модели
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name); // записываем 1-е значение в 1-ю строку, 1-й столбец
                    sheet.SetCellValue(rowIndex, columnIndex: 1, room.Number);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, room.Area);
                    rowIndex++;
                }

                // после всех действий, нужно будет файл Excel закрыть
                workbook.Write(stream);
                workbook.Close();
            }

            // в самом конце, собрали все помещения и собрали все данные, можно запустить файл Excel автоматически
            System.Diagnostics.Process.Start(excelPath);

            return Result.Succeeded;
        }
    }
}
