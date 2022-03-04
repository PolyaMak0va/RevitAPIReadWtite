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

            // научимся считывать данные из файла Excel и записывать высчитанные значения параметров в наши экземпляры семейств
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Excel files(*.xlsx)|*.xlsx"
            };

            string filePath = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }

            // если путь не указан
            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            // должны открыть по указ-му пути файл Excel и просчитать из него данные
            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // создаём переменную для чтения файла Excel
                IWorkbook workbook = new XSSFWorkbook(filePath);

                //далее берём лист, с кот-го будем считывать данные
                ISheet sheet = workbook.GetSheetAt(index: 0); // берём самый первый лист

                //далее проходим построчно файл в Excel и собирать данные
                int rowIndex = 0;
                while (sheet.GetRow(rowIndex) != null)  // до тех пор, пока строка есть, считываем данные, когда не будет - заканчиваем данный цикл
                {
                    // если в указ-й строке в указ-м столбце ячейка пустая либо ячейка пустая во 2-м столбце текущей строке, переходим к след-й итерации
                    if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                        sheet.GetRow(rowIndex).GetCell(1) == null)
                    {
                        rowIndex++;
                        continue;
                    }

                    // считываем данные из файла Excel
                    string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                    string number = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

                    // далее из всех собранных помещений в модели ищем помещение с указ-м номером
                    var room = rooms.FirstOrDefault(r => r.Number.Equals(number));

                    if (room == null)
                    {
                        rowIndex++;
                        continue;
                    }

                    // создаём транзакцию, в кот-ю записываем данные
                    using (var ts = new Transaction(doc, "Set parameters"))
                    {
                        ts.Start();
                        // берём параметр помещения, его имя и вписываем его Set(name)
                        room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);
                        ts.Commit();
                    }
                    // для того, чтобы переходить каждый раз к новой строке, если всё прошло удачно
                    rowIndex++; 
                }
            }
            return Result.Succeeded;
        }
    }
}
