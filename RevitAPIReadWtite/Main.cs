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

            // определим переменную, кот-я будет собирать данные по помещениям: имя, номер помещения и площадь
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

            // теперь нужно эти данные где-то сохранить, напр., на раб. столе
            string decktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // получаем путь к папке раб. стола

            // указываем полный путь к новому файлу с данными roomInfo. Для этого соединяем путь к папке и roomInfo.csv 
            string csvPath = Path.Combine(decktopPath, "roomInfo.csv");

            File.WriteAllText(csvPath, roomInfo);
            return Result.Succeeded;
        }
    }
}
