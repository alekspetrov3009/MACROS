using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using Microsoft.Office.Interop.Excel;
using SL_Maker;
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
    public partial class MainForm : Form
    {
        public static List<string> paths = new List<string>();
        public static List<string> Formats = new List<string>();
        public static List<string> countOfSheets = new List<string>();
        public static List<string> CountOfSpec = new List<string>();
        public static List<string> oboznachenie = new List<string>();
        public static List<string> naimenovanie = new List<string>();

        public static string noteNumber;
        public static string date;
        public static string type;
        public static string order;
        public static string performer;
        public static string folderName;

        public MainForm()
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

        public void ProgressBar(int i)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = paths.Count();
            progressBar1.Step = paths.Count / paths.Count;
            progressBar1.PerformStep();
        }

        //всплывающее уведомление
        public void Notify(string baloonText)
        {
            notifyIcon1.Icon = Icon;
            notifyIcon1.ShowBalloonTip(10000, "Выполнено", baloonText, ToolTipIcon.Info);
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
            textBox1.Clear();
            foreach (string obj in (string[])e.Data.GetData(DataFormats.FileDrop))
                if (Directory.Exists(obj))
                    paths.AddRange(Directory.GetFiles(obj, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".cdw") || s.EndsWith(".spw")));
                else
                    
                    paths.Add(obj);

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

        public void TemplateUpload()
        {
            UsingExcel usingExcel = new UsingExcel();
            string noteNumber = noteNumberTextbox.Text.Replace("/", "-").Replace(@"\", "-");
            File.WriteAllBytes($@"{textBox2.Text}\Cлужебная записка на обработку и размножение чертежей {noteNumber}.xltx", SL_Maker.Properties.Resources.Excel_Template);
            string templatePath = $@"{textBox2.Text}\Cлужебная записка на обработку и размножение чертежей {noteNumber}.xltx";
            usingExcel.OpenExcel(templatePath);
            //Удалить шаблон после заполнения
            File.Delete(templatePath);
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
            UsingKompas usingKompas = new UsingKompas();
            usingKompas.StartKompas();
            usingKompas.CloseKompas();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            TemplateUpload();
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

            //noteNumber = noteNumberTextbox.Text.ToString();
            date = dateTimePicker1.Text.ToString();
            type = comboBox3.Text.ToString();
            order = comboBox2.Text.ToString();
            performer = comboBox1.Text.ToString();
            noteNumber = noteNumberTextbox.Text.Replace("/", "-").Replace(@"\", "-");
            MainForm.folderName = textBox2.Text;
        }
    }
}

