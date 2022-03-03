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

            // рассмотрим запись данных с текстового файла в значения параметров модели

            // в первую очередь, когда будем записывать какие-то данные с текстового файла, надо указать с какого именно текстового файла

            // для этого воспользуемся диалоговым окном
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "All files (*.*)|*.*";

            // надо создать переменную, которую сохранит путь к файлу
            string filePath = string.Empty;

            // если путь к сохранению файла указан, тогда мы забираем этот путь
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }

            // Однако если строка null или пустая, то возвращаем рузельтат и таким образом, заканчиваем выполнение нашей программы
            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;

            // забираем все строки из текстового файла
            var lines = File.ReadAllLines(filePath).ToList();

            // когда считали все данные, надо как-то их сохранить и удобнее работать
            // создаём переменную списка RoomData
            List<RoomData> roomDataList = new List<RoomData>();
            foreach (var line in lines)
            {
                List<string> values = line.Split(';').ToList();         //разделяем каждую строку по значению разделителя
                roomDataList.Add(new RoomData
                {
                    Name = values[0],
                    Number = values[1]
                });
            }

            // теперь остаётся записать собранные данные в наше помещение, кот-е есть в модели.
            // условия: изменяем имя помещения, согласно номеру

            string roomInfo = string.Empty;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            using (var ts = new Transaction(doc, "Set parameters"))
            {
                ts.Start();

                foreach (RoomData roomData in roomDataList)
                {
                    // находим то помещение с номером, которое указано в RoomData
                    Room room = rooms.FirstOrDefault(r => r.Number.Equals(roomData.Number));  // находим помещение, кот-е совпадает по номеру с данными из текстового файла

                    // если помещение не найдено
                    if (room == null)
                        continue;

                    // обращаемся к параметру имени
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(roomData.Name);
                }
                ts.Commit();
            }
            return Result.Succeeded;
        }
    }
}
