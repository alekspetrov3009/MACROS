using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Data;
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
        private static IApplication _kompas7;
        public List<string> Formats = new List<string>();
        List<string> countOfSheets = new List<string>();
        public Form1()
        {
            InitializeComponent();
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

        public void InterfaceConnection()
        {
            string progId = "KOMPAS.Application.5";
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
        }


        public void CloseKompas()
        {
            kompas.Quit();
        }


        public void OpenDrawing()
        {
            InterfaceConnection();
            //используем API  - 7 версии
            KompasAPI7._Application My7Komp = (_Application)kompas.ksGetApplication7();
            IKompasAPIObject retw = My7Komp.ActiveDocument;

            for (int i = 0; i < paths.Count; i++)
            {
                IKompasDocument docOpen = My7Komp.Documents.Open(paths[i], false, true);

                // вызов метода чтения форматов
                ReadDrawings();

                Console.WriteLine(paths[i]);
                Console.WriteLine();
            }
        }


        public void ReadShtamp()
        {
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
            KompasAPI7._Application My7Komp = (_Application)kompas.ksGetApplication7();
            ILayoutSheet MyLSheet = My7Komp.ActiveDocument.LayoutSheets.get_ItemByNumber(1);
            ksStamp stamp = (ksStamp)doc.GetStamp();
            IStamp istamp = MyLSheet.Stamp;

            IText naimenovanie = istamp.Text[1];
            IText oboznachenie = istamp.Text[2];
            Console.WriteLine(naimenovanie.Str);
            Console.WriteLine(oboznachenie.Str);
        }


        public void ReadDrawings()
        {
            

            InterfaceConnection();
            //string progId = "KOMPAS.Application.5";
            //KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();


            if (kompas == null)
            {
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                kompas = (KompasObject)Activator.CreateInstance(t);
            }

            //используем API  - 7 версии
            KompasAPI7._Application My7Komp = (_Application)kompas.ksGetApplication7();
            //IKompasAPIObject retw = My7Komp.ActiveDocument;

            // Количество листов в документе
            int CountPages = doc.ksGetDocumentPagesCount();

            countOfSheets.Add(CountPages.ToString());

            // здесь искомый номер листа
            int numb = 1;

            // итерация по количеству листов в чертеже
            while (numb <= CountPages)
            {
                ILayoutSheet MyLSheet = My7Komp.ActiveDocument.LayoutSheets.get_ItemByNumber(numb);

                // обращаемся к искомому объекту SheetFormat в котром и хранятся все данные о листе
                // формат, ориентация, кратность, высота, ширина,
                ISheetFormat ShFormat = MyLSheet.Format;

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
            //Console.WriteLine(countOfSheets);
            // чтение штампа
            ReadShtamp();
        }


        private void button1_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < paths.Count; i++)
            {
                Console.WriteLine(paths[i]);
            }

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
            foreach (string path in paths)
                textBox1.Text += path + "\r\n";
            int numbersOfSheets = paths.Count();
            label4.Text = numbersOfSheets.ToString();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Formats.Clear();
            paths.Clear();
            textBox1.Clear();
            label4.Text = "0";

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

            //ReadDrawings();
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

  



        //protected void newdocument()
        //{
        //    var application = new Microsoft.Office.Interop.Excel.Application();
        //    var workbook = application.Workbooks.add(template: geturi());
        //    var worksheet = (Worksheet)workbook.sheets[1];
        //}

        //private string geturi()
        //{
        //    var resource = new { name = "служебная_записка_на_обработку_и_размножение_чертежей.xltx ", buff = resources.template };

        //    var tempdirectory = path.getdirectoryname(path.gettempfilename());

        //    var path = string.format("{0}\\{1}", tempdirectory, resource.name);

        //    if (!file.exists(path) || file.readallbytes(path).length.equals(0))
        //    {
        //        var stream = new memorystream(resource.buff);

        //        using (var file = new filestream(path, filemode.create))
        //        {
        //            var buffer = new byte[4096];
        //            int bytesread;

        //            while ((bytesread = stream.read(buffer, 0, buffer.length)) > 0)
        //            {
        //                file.write(buffer, 0, bytesread);
        //            }
        //        }
        //    }

        //    return path;
        //}



    }

}
