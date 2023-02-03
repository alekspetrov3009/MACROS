using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1;

namespace SL_Maker
{
    public class UsingKompas
    {
        public UsingKompas(string Formats, string countOfSheets, string CountOfSpec, string oboznachenie, string naimenovanie) 
        { 
            _Formats = Formats;
            _countOfSheets= countOfSheets;
            _CountOfSpec = CountOfSpec;
            _oboznachenie = oboznachenie;
            _naimenovanie = naimenovanie;
        }

        private KompasObject kompas;
        private string _Formats ;
        private string _countOfSheets;
        private string _CountOfSpec;
        private string _oboznachenie;
        private string _naimenovanie;
        public UsingKompas()
        {

        }
        public void StartKompas()
        {
            if (kompas == null)
            {
#if __LIGHT_VERSION__
			            Type t = Type.GetTypeFromProgID("KOMPASLT.Application.5");
#else
                Type t = Type.GetTypeFromProgID("KOMPAS.Application.5");
#endif
                Console.WriteLine("Запускаем Компас в невидимом режиме..");
                kompas = (KompasObject)Activator.CreateInstance(t); kompas.Visible = false;
                try { kompas = (KompasObject)Activator.CreateInstance(t); kompas.Visible = false; }
                catch (Exception ee1) { MessageBox.Show("Не могу запустить Компас!", "Сообщение"); }
            }
                else
                    MessageBox.Show("Не найден активный объект", "Сообщение");

            OpenDrawing();
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

        private void OpenDrawing()
        {
            //используем API  - 7 версии
            KompasAPI7._Application My7Komp = (KompasAPI7._Application)kompas.ksGetApplication7();

            for (int i = 0; i < MainForm.paths.Count; i++)
            {
                //Пропустить сообщения
                My7Komp.HideMessage = ksHideMessageEnum.ksHideMessageNo;

                IKompasDocument docOpen = My7Komp.Documents.Open(MainForm.paths[i], false, true);

                string fileExtension = Path.GetExtension(MainForm.paths[i]);

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
                //ProgressBar(i);
            }

            string baloonText = "Все данные считаны";
            MainForm mainForm = new MainForm();
            mainForm.Notify(baloonText);
        }

        public void ReadDrawShtamp(IKompasDocument docOpen)
        {
            LayoutSheets _ls = docOpen.LayoutSheets;
            LayoutSheet LS = _ls.ItemByNumber[1];
            IStamp istamp = LS.Stamp;
            IText naimen = istamp.Text[1];
            IText obozn = istamp.Text[2];
            MainForm.naimenovanie.Add(naimen.Str);
            MainForm.oboznachenie.Add(obozn.Str);
        }

        public void ReadDrawings(IKompasDocument docOpen)
        {

            ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();

            // Количество листов в документе
            int CountPages = doc.ksGetDocumentPagesCount();

            MainForm.countOfSheets.Add(CountPages.ToString());

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
                MainForm.Formats.Add(drawingFormat.ToString());
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
            MainForm.countOfSheets.Add(CountSpec.ToString());
            MainForm.Formats.Add("A4");
            Console.WriteLine(CountSpec.ToString());
        }

        public void CloseKompas()
        {
            kompas.Quit();
        }
    }
}