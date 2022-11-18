using KAPITypes;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private KompasObject kompas;
        public List<string> paths = new List<string>();
        private static IApplication _kompas7;
        public List<string> Formats = new List<string>();
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


        public void CloseKompas()
        {
            kompas.Quit();
        }

        public void openDrawing()
        {

            string progId = "KOMPAS.Application.5";
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
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


        public void readShtamp()
        {
            string progId = "KOMPAS.Application.5";
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();
            ksStamp stamp = (ksStamp)doc.GetStamp();
            if (stamp != null)
            {
                if (stamp.ksOpenStamp() == 1)
                {
                    stamp.ksOpenStamp();
                    string[] stampInfo = stamp.ksGetStampColumnText(1);
                }
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            
            for (int i = 0; i < paths.Count; i++)
            {
                Console.WriteLine(paths[i]);
            }
            
        }
          

        public void ReadDrawings()
        {
            //var Formats = new List<string>();

            string progId = "KOMPAS.Application.5";
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();

            
            if (kompas == null)
            {
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                kompas = (KompasObject)Activator.CreateInstance(t);
            }
            
            //используем API  - 7 версии
            KompasAPI7._Application My7Komp = (_Application)kompas.ksGetApplication7();
            IKompasAPIObject retw = My7Komp.ActiveDocument;

            int CountPages = doc.ksGetDocumentPagesCount();

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
                Console.WriteLine(countOfPages);
            }
            
            /*
            foreach (var item in Formats)
            {
                Console.WriteLine(item);
            }
            */

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
        }


        private void button2_Click(object sender, EventArgs e)
        {
            Formats.Clear();
            paths.Clear();
            textBox1.Clear();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            StartKompas();
            openDrawing();
            CloseKompas();

            
            foreach (var item in Formats)
            {
                Console.WriteLine(item);
            }
            
            //ReadDrawings();
        }
    }
    
}


        /*public void textBox1_DragEnter(object sender, DragEventArgs e)
       {
           if (e.Data.GetDataPresent(DataFormats.FileDrop))
           {
               string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
               e.Effect = DragDropEffects.Copy;
           }
       }

       public void textBox1_DragDrop(object sender, DragEventArgs e)
       {
           if (e.Data.GetDataPresent(DataFormats.FileDrop))
           {
               string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
               foreach (string path in paths)
                   textBox1.Text += path + "\r\n";*/

 