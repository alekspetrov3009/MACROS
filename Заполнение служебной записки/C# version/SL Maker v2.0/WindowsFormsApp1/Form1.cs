using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private KompasObject kompas;
        public List<string> paths = new List<string>();
        public List<string> Formats = new List<string>();
        List<string> countOfSheets = new List<string>();
        List<string> CountOfSpec = new List<string>();
        List<string> oboznachenie = new List<string>();
        List<string> naimenovanie = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] surnames = { "Круглов П.А.", "Петров А.И.", "Сорокин А.А.", "Уфрутов Р.С." };
            string[] orderNumbers = { "340434/1" };
            string[] transformerType = { "ЭТЦНР-10500/35-У3" };
            comboBox1.Items.AddRange(surnames);
            comboBox2.Items.AddRange(orderNumbers);
            comboBox3.Items.AddRange(transformerType);
        }

        private void StartKompas()
        {
            if (kompas == null)
            {
#if __LIGHT_VERSION__
			            Type t = Type.GetTypeFromProgID("KOMPASLT.Application.5");
#else
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
#endif
                Console.WriteLine("Запускаем Компас в невидимом режиме..");
                try { kompas = (KompasObject)Activator.CreateInstance(t); kompas.Visible = false; }
                catch (Exception ee1) { MessageBox.Show(this, "Не могу запустить Компас!", "Сообщение"); }
            }

            else
                MessageBox.Show(this, "Не найден активный объект", "Сообщение");
        }

        public void InterfaceConnectionAPI5()
        {
            string progId = "KOMPAS.Application.5";
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();

        }

        public void InterfaceConnectionAPI7()
        {
        }

        public void OpenDrawing()
        {
            //используем API  - 7 версии
            KompasAPI7._Application My7Komp = (KompasAPI7._Application)kompas.ksGetApplication7();

            for (int i = 0; i < paths.Count; i++)
            {
                //Пропустить сообщения
                My7Komp.HideMessage = ksHideMessageEnum.ksHideMessageNo;

                IKompasDocument docOpen = My7Komp.Documents.Open(paths[i], false, true);

                string fileExtension = Path.GetExtension(paths[i]);

                if (fileExtension == ".spw")
                {
                    ReadSpecification();
                }
                else
                {
                    ReadDrawings(docOpen);
                }
                My7Komp.HideMessage = ksHideMessageEnum.ksShowMessage;

                ReadDrawShtamp(docOpen);
                
                ProgressBar(i);
            }
            string baloonText = "Все данные считаны";
            Notify(baloonText);   
        }

        public void ReadDrawShtamp(IKompasDocument docOpen)
        {
            LayoutSheets _ls = docOpen.LayoutSheets;
            LayoutSheet LS = _ls.ItemByNumber[1];

            IStamp istamp = LS.Stamp;

            IText naimen = istamp.Text[1];
            IText obozn = istamp.Text[2];
            naimenovanie.Add(naimen.Str);
            oboznachenie.Add(obozn.Str);
        }

        public void ReadDrawings(IKompasDocument docOpen)
        {

            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();

            // Количество листов в документе
            int CountPages = doc.ksGetDocumentPagesCount();

            countOfSheets.Add(CountPages.ToString());

            // здесь искомый номер листа
            int numb = 1;

            // итерация по количеству листов в чертеже
            //while (numb <= CountPages)   //итерация по всем листам в документе
            while (numb <= 1)       // итерация только по первому листу
            {
                ILayoutSheets _ls = docOpen.LayoutSheets;
                ILayoutSheet LS = _ls.ItemByNumber[numb];

                ISheetFormat ShFormat = LS.Format;

                // получаем формат в ввиде перечисления
                ksDocumentFormatEnum YesFormat = ShFormat.Format;

                // для перевода в человеческий вид	
                int fgg = YesFormat.GetHashCode();
                string drawingFormat = "А" + fgg.ToString();
                // добавляем в список все форматы листов в чертеже
                Formats.Add(drawingFormat.ToString());
                string countOfPages = CountPages.ToString();
                numb++;

                Console.WriteLine(drawingFormat);
            }
            Console.WriteLine(CountPages);


        }

        public void ReadSpecification()
        {
            ksSpcDocument spec = (ksSpcDocument)kompas.SpcActiveDocument();
            // Количество листов в документе
            int CountSpec = spec.ksGetSpcDocumentPagesCount();
            //CountOfSpec.Add(CountSpec.ToString());
            countOfSheets.Add(CountSpec.ToString());
            Formats.Add("A4");
            Console.WriteLine(CountSpec.ToString());
        }

        public void CloseKompas()
        {
            kompas.Quit();
        }

        public void ProgressBar(int i)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = paths.Count();
            progressBar1.Step = paths.Count / paths.Count;
            progressBar1.PerformStep();
        }

   
        public void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                e.Effect = DragDropEffects.Copy;
            }
        }


        public void textBox1_DragDrop(object sender, DragEventArgs e)
        {

            foreach (string obj in (string[])e.Data.GetData(DataFormats.FileDrop))
                if (Directory.Exists(obj))
                    paths.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".cdw") || s.EndsWith(".spw")));
                else
                    
                    paths.Add(obj);
            //textBox1.Text = paths.EndsWith(".cdw") || paths.EndsWith(".spw");

            // Обработка повторяющихся путей в списке
            for (int i = 0; i < paths.Count; i++)
            {
                for (int j = i + 1; j < paths.Count; j++)
                {
                    if (paths[i] == paths[j])
                    {
                        paths.RemoveAt(j);
                    }
                    textBox1.Clear();
                }
            }
            foreach (string path in paths)
                textBox1.Text += path + "\r\n";
            int numbersOfSheets = paths.Count();
            label1.Text = $"Добавлено файлов: {numbersOfSheets}";

        }


        private void button2_Click(object sender, EventArgs e)
        {
            Formats.Clear();
            paths.Clear();
            textBox1.Clear();
            label1.Text = "Добавлено файлов: 0";
            System.Windows.Forms.Application.Restart();
        }


        private void button3_Click(object sender, EventArgs e)
        {
            StartKompas();
            OpenDrawing();
            CloseKompas();




            foreach (var item in Formats)
            {
                Console.WriteLine(item);
            }

            foreach (var sheets in countOfSheets)
            {
                Console.WriteLine(sheets);
            }


            //MessageBox.Show("ok");

            //ReadDrawings();
            //Notify(string baloonText);
        }


        private void button4_Click(object sender, EventArgs e)
        {
            TemplateUpload();
            //OpenExcel();

        }


        private void сhooseFolderButton_Click(object sender, EventArgs e)
        {
            DialogResult dialogresult = folderBrowserDialog1.ShowDialog();
            //Надпись выше окна контрола
            folderBrowserDialog1.Description = "Выбор папки";
            string folderName = "";
            if (dialogresult == DialogResult.OK)
            {
                //Извлечение имени папки
                folderName = folderBrowserDialog1.SelectedPath;
            }
            textBox2.Text = folderName;
        }


        //выгрузка шаблона
        public void TemplateUpload()
        {
            string noteNumber = noteNumberTextbox.Text.Replace("/", "-").Replace(@"\", "-");
            File.WriteAllBytes($@"{textBox2.Text}\Cлужебная записка на обработку и размножение чертежей {noteNumber}.xltx", SL_Maker.Properties.Resources.Excel_Template);
            string templatePath = $@"{textBox2.Text}\Cлужебная записка на обработку и размножение чертежей {noteNumber}.xltx";
            OpenExcel(templatePath);
            //Удалить шаблон после заполнения
            File.Delete(templatePath);
        }

        //всплывающее уведомление
        public void Notify(string baloonText)
        {
            notifyIcon1.Icon = Icon;
            notifyIcon1.ShowBalloonTip(10000, "Выполнено", baloonText, ToolTipIcon.Info);
        }
        
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

        public void FillExcel(Microsoft.Office.Interop.Excel.Application app)
        {
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1)
            Worksheet sheet = (Worksheet)app.Worksheets.get_Item(1);

            
            //Номер служебной записки
            sheet.Range["B11"].Value = noteNumberTextbox.Text.ToString();
            //Дата составления
            sheet.Range["C11"].Value = dateTimePicker1.Text.ToString();
            //Тип трансформатора
            sheet.Range["D11"].Value = comboBox3.Text.ToString();
            //Номер заказа
            sheet.Range["E11"].Value = comboBox2.Text.ToString();


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
                lastCell= 10 + a;    
            }

            //Руководитель проекта
            sheet.Range[$"D{lastCell + 4}"].Value = "Руководитель проекта";
            sheet.Range[$"J{lastCell + 4}"].Value = "Уфрутов Р.С.";
            //Исполнитель
            sheet.Range[$"D{lastCell + 6}"].Value = "Исполнитель";
            sheet.Range[$"J{lastCell + 6}"].Value = comboBox1.Text.ToString();

            FormatCells(sheet, lastCell);
            SaveExcel(app);
        }

        public void FormatCells(Worksheet sheet, int lastcell)
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
            for (int i = 0; i <= 3; )
            {
                string[] Columns = { "B", "C", "D", "E" };
                Range mergeCells = sheet.get_Range($"{Columns[i]}11", $"{Columns[i]}{lastcell}");
                mergeCells.Merge(Type.Missing);
                i++;
            }
        }

        private void SaveExcel(Microsoft.Office.Interop.Excel.Application app)
        {
            string noteNumber = noteNumberTextbox.Text.Replace("/", "-").Replace(@"\", "-");
            app.Application.ActiveWorkbook.SaveAs($@"{textBox2.Text}\Cлужебная записка на обработку и размножение чертежей {noteNumber}.xlsx");
            //закрытие приложения
            app.Quit();
            Marshal.ReleaseComObject(app);
        }
    }
}

