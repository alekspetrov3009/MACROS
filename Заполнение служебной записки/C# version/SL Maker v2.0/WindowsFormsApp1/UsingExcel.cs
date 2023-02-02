using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1;


namespace SL_Maker
{
    public class UsingExcel : MainForm
    {
        //выгрузка шаблона
        

        public void OpenExcel(string templatePath)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Workbooks.Open(templatePath);
            app.Visible = false;
            Console.WriteLine(templatePath);
            FillExcel(app);

            string baloonText = "Служебная записка создана";
            Notify(baloonText);
        }

        private void FillExcel(Microsoft.Office.Interop.Excel.Application app)
        {
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1)
            Worksheet sheet = (Worksheet)app.Worksheets.get_Item(1);

            //Номер служебной записки
            sheet.Range["B11"].Value = MainForm.noteNumber;
            //Дата составления
            sheet.Range["C11"].Value = MainForm.date;
            //Тип трансформатора
            sheet.Range["D11"].Value = MainForm.type;
            //Номер заказа
            sheet.Range["E11"].Value = MainForm.order;

            //Заполнение массивов
            int i = 11; //строка
            int lastCell = 0; //номер последней ячейки
            int stringNumber = 1; // Номер позиции в таблице
            //string withoutBTLI = oboznachenie.Remove("БТЛИ.");
            for (int a = 0; paths.Count > a;)
            {
                sheet.Cells[i, 9].Value = "БТЛИ.";
                sheet.Cells[i, 10].Value = oboznachenie[a].Remove(0, 5); //удаляет БТЛИ.
                sheet.Cells[i, 11].Value = naimenovanie[a];
                sheet.Cells[i, 12].Value = countOfSheets[a];
                sheet.Cells[i, 13].Value = Formats[a];
                sheet.Cells[i, 1].Value = stringNumber;
                i++;
                a++;
                stringNumber++;
                lastCell = 10 + a;
            }

            //Руководитель проекта
            sheet.Range[$"D{lastCell + 4}"].Value = "Руководитель проекта";
            sheet.Range[$"J{lastCell + 4}"].Value = "Уфрутов Р.С.";

            //Исполнитель
            sheet.Range[$"D{lastCell + 6}"].Value = "Исполнитель";
            sheet.Range[$"J{lastCell + 6}"].Value = MainForm.performer;

            FormatCells(sheet, lastCell);
            SaveExcel(app);
        }

        private void FormatCells(Worksheet sheet, int lastcell)
        {
            //Выбор диапазона
            Range cellsRange = sheet.get_Range("A11", $"M{lastcell}");

            //Шрифт для диапазона
            cellsRange.Cells.Font.Name = "Times New Roman";

            //Размер шрифта для диапазона
            cellsRange.Cells.Font.Size = 11;

            //Обводка ячеек
            cellsRange.Borders.Color = ColorTranslator.ToOle(Color.Black);

            // Выравнивание текста в ячейках по центру
            cellsRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            cellsRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            //Объединение ячеек 
            for (int i = 0; i <= 3;)
            {
                string[] Columns = { "B", "C", "D", "E" };
                Range mergeCells = sheet.get_Range($"{Columns[i]}11", $"{Columns[i]}{lastcell}");
                mergeCells.Merge(Type.Missing);
                i++;
            }
        }

        private void SaveExcel(Microsoft.Office.Interop.Excel.Application app)
        {
            app.Application.ActiveWorkbook.SaveAs($@"{folderName}\Cлужебная записка на обработку и размножение чертежей {noteNumber}.xlsx");

            //закрытие приложения
            app.Quit();
            Marshal.ReleaseComObject(app);
        }
    }
}
