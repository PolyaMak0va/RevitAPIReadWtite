using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
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

            // научимся записывать данные из файла ревита в формат .json
            // сначала надо загрузить библиотеку JSON
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            var roomDataList = new List<RoomData>();

            // далее будем записывать определённые данные по помещениям
            // пройдёмся по каждому помещению
            foreach (var room in rooms)
            {
                roomDataList.Add(new RoomData
                {
                    Name = room.Name,
                    Number = room.Number
                });
            }

            // преобразуем roomDataList в формат .json
            string json = JsonConvert.SerializeObject(roomDataList, Formatting.Indented);

            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "data.json"), json);

            return Result.Succeeded;
        }
    }
}
