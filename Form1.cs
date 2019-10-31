using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using AutoItX3Lib;
using System.IO;
namespace SPY_PostWin
{
    public partial class Form1 : Form
    {
        int time = 0;   //назначаем закрытую переменную time

        public Form1()
        {
            TopMost = true;//поверх всех окон
            InitializeComponent();
            button2.Enabled = false; //кнопка стоп
        }

        //---------------Ракета ----------------------------------------------------->>>>>>>>>>>
        private void button1_Click(object sender, EventArgs e) 
        {
            bool status = true; // переменная для хранения статуса
            timer1.Enabled = true;  //активируем таймер
            pictureBox1.Visible = true;
            progressBar1.Visible = true;
            label3.Visible = true;
            button2.Enabled = true;
            button1.Enabled = false;
            button4.Enabled = false;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            tabPage3.Enabled = false;
            if (status) { timer2.Enabled = true; } else { timer2.Enabled = false; }               
        }
        //---------------Ракета -----------------------------------------------------<<<<<<<<<<<<<


        //_____________________Таймер прогресс бара_____________________________________
        private void timer1_Tick(object sender, EventArgs e) 
        {
            time += 1; 
            progressBar1.Value = time;   
            if (progressBar1.Value == 100)   
            {
                progressBar1.Value = 0;  
                time = 0;  
            }

        }
        //_____________________Таймер прогресс бара Конец_____________________________________


        //_____________________Кнопка остановки программы_____________________________________
        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            button4.Enabled = true;
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            pictureBox1.Visible = false;
            progressBar1.Visible = false;
            label3.Visible = false;
            timer2.Enabled = false;
            timer1.Enabled = false;
            tabPage3.Enabled = true;
        }
        //_____________________Кнопка остановки программы конец_____________________________________

        //Диалог открытия .EXE PostWin
        private void button4_Click(object sender, EventArgs e) 
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        // Блок
        private void timer2_Tick(object sender, EventArgs e)
        {
            AutoItX3 autoIt = new AutoItX3();
            if (autoIt.WinExists("ОШИБКА") == 1)
            {

                string t = autoIt.ControlGetText("ОШИБКА", "", "TStaticText1");
                string pos = "ПЕРЕПОЛНЕНИЕ";
                if (t.Contains(pos))
                {
                    autoIt.WinActivate("ОШИБКА");
                    autoIt.WinGetState("ОШИБКА");
                    autoIt.ControlClick("ОШИБКА", "", "[CLASS:TBitBtn; INSTANCE:1]", "main", 1);
                    autoIt.Sleep(5000);
                    autoIt.WinClose(textBox2.Text);
                    autoIt.Sleep(4000);
                    autoIt.WinActivate("ЗАВЕРШЕНИЕ РАБОТЫ");
                    autoIt.WinGetState("ЗАВЕРШЕНИЕ РАБОТЫ");
                    autoIt.Sleep(2000);
                    autoIt.ControlClick("ЗАВЕРШЕНИЕ РАБОТЫ", "", "[CLASS:Button; INSTANCE:1]", "main", 1);
                    autoIt.Sleep(Convert.ToInt16(textBox5.Text));
                    autoIt.Run(textBox1.Text);
                    autoIt.Sleep(7000);
                    autoIt.WinActivate("ВЫБОР БИБЛИОТЕК");
                    autoIt.WinGetState("ВЫБОР БИБЛИОТЕК");
                    autoIt.Sleep(3000);
                    //Выбор библиотеки поцвету пикселя начало
                    object coord = autoIt.PixelSearch(0, 0, Screen.GetWorkingArea(this).Width, Screen.GetWorkingArea(this).Height, Convert.ToInt32(Convert.ToString(textBox3.Text), 16));//edit texbox3 hex
                    object[] pixelcoord = (object[])coord;
                    autoIt.MouseClick("Left", (int)pixelcoord[0], (int)pixelcoord[1], 1, 10);
                    //Выбор библиотеки поцвету пикселя конец
                    autoIt.Sleep(Convert.ToInt16(textBox6.Text));
                    //Запуск PostWin поцвету пикселя начало
                    autoIt.WinActivate(textBox2.Text);
                    autoIt.WinGetState(textBox2.Text);
                    autoIt.Sleep(2000);
                    object coord1 = autoIt.PixelSearch(0, 0, Screen.GetWorkingArea(this).Width, Screen.GetWorkingArea(this).Height, Convert.ToInt32(Convert.ToString(textBox4.Text), 16));//edit textbox4 hex
                    object[] pixelcoord1 = (object[])coord1;
                    autoIt.MouseClick("Left", (int)pixelcoord1[0], (int)pixelcoord1[1], 1, 10);
                    //Запуск PostWin поцвету пикселя конец
                    autoIt.Sleep(2000);
                //*** Запись лога
                    string date = DateTime.Now.ToString("dd MMMM yyyy,HH:mm:ss");
                    System.IO.StreamWriter writer = new System.IO.StreamWriter("slib.dll", true);
                    writer.WriteLine(date);
                    writer.Close();
                //*** Запись лога конец

                }
                else timer2.Interval = 1; // Выбираем из массива случайное значение для таймера
            }
        }
        private void tabPage2_Click(object sender, EventArgs e)
        {

            StreamReader rd = new StreamReader("slib.dll");
            DataSet ds = new DataSet();
            ds.Tables.Add("log");
            string header = rd.ReadLine();
            string[] col = System.Text.RegularExpressions.Regex.Split(header, ",");
                for (int c = 0; c < col.Length; c++)
                {
                    ds.Tables[0].Columns.Add(col[c]);
                }
                string row = rd.ReadLine();
                while (row != null)
                {
                    string[] rvalue = System.Text.RegularExpressions.Regex.Split(row, ",");
                    ds.Tables[0].Rows.Add(rvalue);
                    row = rd.ReadLine();

                }
                dataGridView1.DataSource = ds.Tables[0];
                dataGridView1.Update();
                rd.Close();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            File.Delete("slib.dll");
            string date = DateTime.Now.ToString("Дата,Время");
            System.IO.StreamWriter writer = new System.IO.StreamWriter("slib.dll", true);
            writer.WriteLine(date);
            writer.Close();
            dataGridView1.Update();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            StreamReader rd = new StreamReader("slib.dll");
            DataSet ds = new DataSet();
            ds.Tables.Add("log");
            string header = rd.ReadLine();
            string[] col = System.Text.RegularExpressions.Regex.Split(header, ",");
            for (int c = 0; c < col.Length; c++)
            {
                ds.Tables[0].Columns.Add(col[c]);
            }
            string row = rd.ReadLine();
            while (row != null)
            {
                string[] rvalue = System.Text.RegularExpressions.Regex.Split(row, ",");
                ds.Tables[0].Rows.Add(rvalue);
                row = rd.ReadLine();

            }
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Update();
            rd.Close();
        }
      }

    }
