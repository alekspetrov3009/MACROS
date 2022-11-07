using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

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
                int numberOfSheets = CountPages - numb;

                for (int i = 0; i <= CountPages; i++)
                    Console.WriteLine("sdfsdf");

                // правильней наверное было идти по шагам постепенно
                // сначало ActiveDocument, затем LayoutSheets, затем get_ItemByNumber
                // т.е. мы пропустили коллекцию листов - LayoutSheets - и сразу обратились по номеру
                LayoutSheet MyLSheet = My7Komp.ActiveDocument.LayoutSheets.get_ItemByNumber(numberOfSheets);

                
                // обращаемся к искомому объекту SheetFormat в котром и хранятся все данные о листе
                // формат, ориентация, кратность, высота, ширина,
                SheetFormat ShFormat = MyLSheet.Format;

                // получаем формат в ввиде перечисления
                
                ksDocumentFormatEnum YesFormat = ShFormat.Format;

                
                // радуемся но не совсем,
                //label1.Text = YesFormat.ToString();//тк получим значение = ksFormatA3

                // для перевода в человеческий вид	
                int fgg = YesFormat.GetHashCode();
                label2.Text = "А" + fgg.ToString();
                label3.Text = CountPages.ToString();
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}