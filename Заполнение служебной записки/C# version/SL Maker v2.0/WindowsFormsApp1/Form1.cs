using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private KompasObject kompas;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Console.WriteLine(paths[0]);
        }
/*            var Formats = new List<string>();
            // подключаемся к текущей сессии
            string progId = "KOMPAS.Application.5";
            
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject(progId);
            try
            {
                kompas = (KompasObject)Marshal.GetActiveObject(progId);
            }
            catch (Exception)
            {
                MessageBox.Show("Подключение к Компасу не прошло");
            }
            if (kompas != null)
            {

                kompas.Visible = true;
                kompas.ActivateControllerAPI();
                ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();

                // Узнаём количество листов в документе,
                // соответственно можно организовать цикл по прохождениям по всем листам
                // важно помнить что нумерация начинается с 1 - как и указана в штампе чертежа
                int CountPages = doc.ksGetDocumentPagesCount();

                //используем API  - 7 версии
                KompasAPI7._Application My7Komp = (_Application)kompas.ksGetApplication7();
                IKompasAPIObject retw = My7Komp.ActiveDocument;


                // здесь искомый номер листа
                int numb = 1;

                // итерация по количеству листов в чертеже
                while (numb <= CountPages)
                {

                    // правильней наверное было идти по шагам постепенно
                    // сначало ActiveDocument, затем LayoutSheets, затем get_ItemByNumber
                    // т.е. мы пропустили коллекцию листов - LayoutSheets - и сразу обратились по номеру
                    ILayoutSheet MyLSheet = My7Komp.ActiveDocument.LayoutSheets.get_ItemByNumber(numb);


                    // обращаемся к искомому объекту SheetFormat в котром и хранятся все данные о листе
                    // формат, ориентация, кратность, высота, ширина,
                    ISheetFormat ShFormat = MyLSheet.Format;
                    //ISheetFormat count = MyLSheet.Count;


                    // получаем формат в ввиде перечисления

                    ksDocumentFormatEnum YesFormat = ShFormat.Format;

                    // радуемся но не совсем,
                    //label1.Text = YesFormat.ToString();//тк получим значение = ksFormatA3

                    // для перевода в человеческий вид	
                    int fgg = YesFormat.GetHashCode();
                    label2.Text = "А" + fgg.ToString();
                    // добавляем в список все форматы листов в чертеже
                    Formats.Add(label2.Text.ToString());
                    label3.Text = CountPages.ToString();
                    numb++;

                }
                foreach (var item in Formats)
                {
                    Console.WriteLine(item);
                }
            }
        }*/
        public void readDrawings()
        {
            var Formats = new List<string>();
            if (kompas == null)
            {
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                kompas = (KompasObject)Activator.CreateInstance(t);
            }

            if (kompas != null)
            {
                kompas.Visible = true;
                kompas.ActivateControllerAPI();
            }

            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();

            // Узнаём количество листов в документе,
            // соответственно можно организовать цикл по прохождениям по всем листам
            // важно помнить что нумерация начинается с 1 - как и указана в штампе чертежа
            int CountPages = doc.ksGetDocumentPagesCount();

            //используем API  - 7 версии
            KompasAPI7._Application My7Komp = (_Application)kompas.ksGetApplication7();
            IKompasAPIObject retw = My7Komp.ActiveDocument;


            // здесь искомый номер листа
            int numb = 1;

            // итерация по количеству листов в чертеже
            while (numb <= CountPages)
            {

                // правильней наверное было идти по шагам постепенно
                // сначало ActiveDocument, затем LayoutSheets, затем get_ItemByNumber
                // т.е. мы пропустили коллекцию листов - LayoutSheets - и сразу обратились по номеру
                ILayoutSheet MyLSheet = My7Komp.ActiveDocument.LayoutSheets.get_ItemByNumber(numb);


                // обращаемся к искомому объекту SheetFormat в котром и хранятся все данные о листе
                // формат, ориентация, кратность, высота, ширина,
                ISheetFormat ShFormat = MyLSheet.Format;
                //ISheetFormat count = MyLSheet.Count;


                // получаем формат в ввиде перечисления

                ksDocumentFormatEnum YesFormat = ShFormat.Format;

                // радуемся но не совсем,
                //label1.Text = YesFormat.ToString();//тк получим значение = ksFormatA3

                // для перевода в человеческий вид	
                int fgg = YesFormat.GetHashCode();
                label2.Text = "А" + fgg.ToString();
                // добавляем в список все форматы листов в чертеже
                Formats.Add(label2.Text.ToString());
                label3.Text = CountPages.ToString();
                numb++;

            }
            foreach (var item in Formats)
            {
                Console.WriteLine(item);
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
            List<string> paths = new List<string>();
            foreach (string obj in (string[])e.Data.GetData(DataFormats.FileDrop))
                if (Directory.Exists(obj))
                    paths.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".cdw") || s.EndsWith(".spw")));

                else
                    paths.Add(obj);
            foreach (string path in paths)
                textBox1.Text += path + "\r\n";


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

 