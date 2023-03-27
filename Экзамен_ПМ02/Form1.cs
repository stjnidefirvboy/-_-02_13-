using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Экзамен_ПМ02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        double main_cost;
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void calculation_Click(object sender, EventArgs e)
        {   
            //проверка
            try
            {
                string txt = comboBox1.Text;
                double cost;
                double quantity = Convert.ToDouble(textBox1.Text);

                //расчёт стоимости билета
                if (txt == "Лебединое озеро")
                {
                    cost = 3000;
                    //проверка на выбор сектора
                    if (radioButton1.Checked)
                    {
                        cost *= 1.5;
                        //проверка на ввод количества билетов
                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton2.Checked)
                    {
                        cost *= 1.07;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton3.Checked)
                    {
                        cost *= 1.2;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else MessageBox.Show("Выберите сектор");
                }
                else if (txt == "Красная шапочка")
                {
                    cost = 2900;

                    if (radioButton1.Checked)
                    {
                        cost *= 1.5;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton2.Checked)
                    {
                        cost *= 1.07;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton3.Checked)
                    {
                        cost *= 1.2;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else MessageBox.Show("Выберите сектор");
                }
                else if (txt == "Летучий корабль")
                {
                    cost = 2700;

                    if (radioButton1.Checked)
                    {
                        cost *= 1.5;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton2.Checked)
                    {
                        cost *= 1.07;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton3.Checked)
                    {
                        cost *= 1.2;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else MessageBox.Show("Выберите сектор");
                }
                else if (txt == "Донкихот")
                {
                    cost = 2800;

                    if (radioButton1.Checked)
                    {
                        cost *= 1.5;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton2.Checked)
                    {
                        cost *= 1.07;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton3.Checked)
                    {
                        cost *= 1.2;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else MessageBox.Show("Выберите сектор");
                }
                else if (txt == "Алые паруса")
                {
                    cost = 3500;

                    if (radioButton1.Checked)
                    {
                        cost *= 1.5;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton2.Checked)
                    {
                        cost *= 1.07;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton3.Checked)
                    {
                        cost *= 1.2;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else MessageBox.Show("Выберите сектор");
                }
                else if (txt == "Щелкунчик")
                {
                    cost = 4000;

                    if (radioButton1.Checked)
                    {
                        cost *= 1.5;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton2.Checked)
                    {
                        cost *= 1.07;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else if (radioButton3.Checked)
                    {
                        cost *= 1.2;

                        if (quantity > 10)
                        {
                            cost = (cost / 100) * 95;
                        }
                        else if (quantity > 15)
                        {
                            cost = (cost / 100) * 93;
                        }
                        else if (quantity > 20)
                        {
                            cost = (cost / 100) * 90;
                        }
                        else if (quantity > 30)
                        {
                            cost = (cost / 100) * 75;
                        }
                        else MessageBox.Show("Введите количество билетов");

                        label1.Text = "Стоимость: " + cost.ToString();

                    }
                    else MessageBox.Show("Выберите сектор");
                }
                else MessageBox.Show("Выберите представление");
            }
            catch
            {
                MessageBox.Show("Error");
            }

        }

        private void create_bill_Click(object sender, EventArgs e)
        {
            try
            {
                //в файле хранится номер для заказа
                int code = Convert.ToInt32(File.ReadAllText(@"код.txt"));
                string date = DateTime.Now.ToShortDateString();
                File.Copy(@"чек.docx", @"чеки\" + code + "_" + date + "_" + main_cost + ".docx");
                //копируем образец чека и работаем с ним
                Word.Document doc = null;
                Word.Application app = new Word.Application();
                string sourse = @"C:\Users\Yaroslav\Desktop\Экзамен_ПМ02\Экзамен_ПМ02\bin\Debug\чеки\" + code + "_" + date + "_" + main_cost + ".docx";
                doc = app.Documents.Open(sourse);
                doc.Activate();
                Word.Bookmarks book = doc.Bookmarks;
                Word.Range range;
                int i = 0;
                string[] data = new string[] { code.ToString(), date.ToString(), comboBox1.Text, main_cost.ToString() };
                foreach (Word.Bookmark b in book)
                {
                    range = b.Range;
                    range.Text = data[i];
                    i++;
                }
                doc.Close();
                doc = null;
                code++;
                //перезаписываем код в файл
                File.WriteAllText(@"код.txt", code.ToString());
                MessageBox.Show("Успешно!");
            }
            catch
            {
                MessageBox.Show("Ошибка, попробуйте ещё раз");
            }
        }

        private void advert_Click(object sender, EventArgs e)
        {

        }
    }
}
