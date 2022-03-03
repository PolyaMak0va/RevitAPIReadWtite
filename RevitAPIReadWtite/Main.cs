using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
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

            // научимся задавать путь сохранения нашего файла
            string roomInfo = string.Empty;

            // собираем сами помещения
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            foreach (Room room in rooms)
            {
                // извлекаем имя помещения
                string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                // заполняем переменную roomInfo новыми данными
                roomInfo += $"{roomName}\t{room.Number}\t{room.Area}{Environment.NewLine}";
            }

            // для того, чтобы вызвать путь к сохранению файла
            var saveDialog = new SaveFileDialog
            {
                // свойство OverwritePrompt, если файл уже существует, выдавать запрос на его перезапись
                OverwritePrompt = true,                                                              // выдавать
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),     // с какой папки начинается сравнение файла
                Filter = "All files (*.*)|(*.*)",                                                    // будут отображаться файлы всех форматов
                FileName = "roomInfo.csv",                                                           // при этом можно будет изменить наименование файла
                DefaultExt = ".csv"                                                                  // расширение по умолчанию
            };

            // создадим переменную, которую сохранит выбранный пользователем путь
            string selectedFilePath = string.Empty;
            if (saveDialog.ShowDialog() == DialogResult.OK)                                          //если сохранение файла прошло успешно
            {
                selectedFilePath = saveDialog.FileName;                                              // тогда я забираю введённый путь и сохраняю его в selectedFilePath
            }

            // если путь не был указан, то стоит выйти из данного метода и завершить выполнение программы
            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            // если всё в порядке, вызываю статический метод
            File.WriteAllText(selectedFilePath, roomInfo);

            return Result.Succeeded;
        }
    }
}
