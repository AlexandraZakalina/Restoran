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

namespace MDK0101Ekz
{
    public partial class Form1 : Form
    {
        int sum = 0;//Сумма заказа
        int sumsal = 0;//Cумма салатов
        int sumvt = 0;//Сумма второго
        int sumnap = 0;//Сумма напитков
        int sdacha = 0;//Сдача
        int money = 0;//Внесенные
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Салаты
            if(comboBox1.Text == "Оливье")
            {
                if(textBox1.Text == "")
                {
                    MessageBox.Show("Не введено кол-во порций салатов!");
                }
                else
                {
                    sumsal = 100 * Convert.ToInt32(textBox1.Text);
                    sum += sumsal;
                }
            }
            else
            {
                if (comboBox1.Text == "Летний")
                {
                    if (textBox1.Text == "")
                    {
                        MessageBox.Show("Не введено кол-во порций салатов!");
                    }
                    else
                    {
                        sumsal = 50 * Convert.ToInt32(textBox1.Text);
                        sum += sumsal;
                    }
                }
            }

            //Второе
            if (comboBox2.Text == "Овощи гриль с мясом")
            {
                if (textBox2.Text == "")
                {
                    MessageBox.Show("Не введено кол-во порций второго!");
                }
                else
                {
                    sumvt = 240 * Convert.ToInt32(textBox2.Text);
                    sum += sumvt;
                }
            }
            else
            {
                if (comboBox2.Text == "Пюре с котлетой")
                {
                    if (textBox2.Text == "")
                    {
                        MessageBox.Show("Не введено кол-во порций второго!");
                    }
                    else
                    {
                        sumvt = 200 * Convert.ToInt32(textBox2.Text);
                        sum += sumvt;
                    }
                }
            }

            //Напитки
            if (comboBox3.Text == "Чай")
            {
                if (textBox3.Text == "")
                {
                    MessageBox.Show("Не введено кол-во порций напитков!");
                }
                else
                {
                    sumnap = 30 * Convert.ToInt32(textBox3.Text);
                    sum += sumnap;
                }
            }
            else
            {
                if (comboBox3.Text == "Кофе")
                {
                    if (textBox3.Text == "")
                    {
                        MessageBox.Show("Не введено кол-во порций напитков!");
                    }
                    else
                    {
                        sumnap = 60 * Convert.ToInt32(textBox3.Text);
                        sum += sumnap;
                    }
                }
            }

            //Вывод суммы
            textBox4.Text = Convert.ToString(sum)+" p.";

            if(Convert.ToInt32(textBox5.Text)<sum)
            {
                MessageBox.Show("Недостаточно средств");
            }
            else
            {
                //Сдача
                textBox6.Text = Convert.ToString(Convert.ToInt32(textBox5.Text) - sum);
            }

            money = Convert.ToInt32(textBox5.Text);
            sdacha = Convert.ToInt32(textBox6.Text);
        }

        //функция для Word
        private void Chek(string subToReplace, string text, Word.Document worddoc)
        {
            var range = worddoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: subToReplace, ReplaceWith: text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Квитанция сохранена");// вывод сообщения о сохранении
            var WordApp = new Word.Application();
            WordApp.Visible = false;
            //путь к шаблону
            var Worddoc = WordApp.Documents.Open(Application.StartupPath + @"\Бланк-квитанция1.docx");
            //заполнение
            Chek("{sumsal}", sumsal.ToString(), Worddoc);
            Chek("{sumvt}", sumvt.ToString(), Worddoc);
            Chek("{sumnap}", sumnap.ToString(), Worddoc);
            Chek("{sum}", sum.ToString(), Worddoc);
            Chek("{money}", money.ToString(), Worddoc);
            Chek("{sdacha}", sdacha.ToString(), Worddoc);
            Chek("{date}", DateTime.Now.ToLongDateString(), Worddoc);
            
            //сохранение документа
            Worddoc.SaveAs2(Application.StartupPath + $"\\Квитанция на сумму {sum} от {DateTime.Now.ToLongDateString()}" + ".docx");
            //открываем документ
            WordApp.Visible = true;
        }
    }
}
