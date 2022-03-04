using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIReadWtite
{
    // сделаем метод расширения для работы с листом: для записи в ячейки значения
    public static class SheetExts
    {
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)
        {
            // ссылка на ячейку
            var cellReference = new CellReference(rowIndex, columnIndex);

            // проверяем, если ячейки не существует, то надо её создать, для этого надо проверить наличие строки и наличие столбца
            var row = sheet.GetRow(cellReference.Row);

            // если строки с таким индексом нет, то надо её создать
            if (row == null)
                row = sheet.CreateRow(cellReference.Row);

            // создаём ссылку на ячейку и в выбранной строке указываем конкретную ячейку
            var cell = row.GetCell(cellReference.Col);
            if (cell == null)
                cell = row.CreateCell(cellReference.Col);

            // нужно проверить, что явл-ся значением, кот-е нужно записать в ячейку
            if (value is string)
            {
                cell.SetCellValue((string)(object)value); // если значение - текст, записываем как текст
            }
            else if (value is double)
            {
                cell.SetCellValue((double)(object)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)(object)value);
            }
        }
    }
}
