using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Events
{
    public partial class Form1 : Form
    {
        SQLfunctions functions = new SQLfunctions();
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateEventsGrid();
        }
        #region EventsPan
        void UpdateEventsGrid()
        {
            dataGridView1.DataSource = functions.SelectAllFromEvents();
            DataTable dt = new DataTable();
            dt = functions.SelectInComboboxClient();
            comboBox1.DataSource = dt;
            comboBox1.ValueMember = "Clientid";
            comboBox1.DisplayMember = "FIO";

            dt = functions.SelectInComboboxResponsible();
            comboBox2.DataSource = dt;
            comboBox2.ValueMember = "Id";
            comboBox2.DisplayMember = "FIO";

            dt = functions.SelectInComboboxType();
            comboBox3.DataSource = dt;
            comboBox3.ValueMember = "Id";
            comboBox3.DisplayMember = "Type";

            dt = functions.SelectInComboboxAdress(int.Parse(comboBox2.SelectedValue.ToString()));
            comboBox4.DataSource = dt;
            comboBox4.ValueMember = "Id";
            comboBox4.DisplayMember = "Adress";

            dataGridView2.DataSource = functions.SelectAllClients_Responsible("Clients");
            dataGridView3.DataSource = functions.SelectAllClients_Responsible("Responsible");
            dataGridView4.DataSource = functions.SelectAllFromAdresses();

            dt = functions.SelectInComboboxType();
            comboBox5.DataSource = dt;
            comboBox5.ValueMember = "Id";
            comboBox5.DisplayMember = "Type";



        }
        private void AddEvent_Click(object sender, EventArgs e)
        {
            if(textBox1.Text =="")
            {
                MessageBox.Show("Заполните поле");
                return;
            }
            functions.AddEvent(textBox1.Text, int.Parse(comboBox1.SelectedValue.ToString()), int.Parse(comboBox2.SelectedValue.ToString()), int.Parse(comboBox4.SelectedValue.ToString()), textBox2.Text, dateTimePicker1.Value.ToString("yyyy-MM-dd"));
            MessageBox.Show("Запись успешно добавлена!");
            UpdateEventsGrid();
        }
        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();
                dt = functions.SelectInComboboxAdress(int.Parse(comboBox3.SelectedValue.ToString()));
                comboBox4.DataSource = dt;
                comboBox4.ValueMember = "Id";
                comboBox4.DisplayMember = "Adress";
            }
            catch { }
        }


        #endregion

        private void DeleteEvent_Click(object sender, EventArgs e)
        {
            if(functions.DeleteFromGridByIndex("Events",int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString())))
            MessageBox.Show("Запись успешно удалена");
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
            UpdateEventsGrid();
        }
        private void EditEvent_Click(object sender, EventArgs e)
        {
            functions.EditEvent(textBox1.Text, int.Parse(comboBox1.SelectedValue.ToString()), int.Parse(comboBox2.SelectedValue.ToString()), int.Parse(comboBox4.SelectedValue.ToString()), textBox2.Text, dateTimePicker1.Value.ToString("yyyy-MM-dd"),int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()));
            UpdateEventsGrid();
            MessageBox.Show("Запись успешно изменена");
        }

        private void AddClient_Click(object sender, EventArgs e)
        {
            if (SurnameIPF.Text == "" || NameIPF.Text == "" || LastNameIPF.Text == "" || maskedTextBox1.Text =="")
            {
                MessageBox.Show("Заполните поле");
                return;
            }
            functions.AddClient_responsible("Clients", SurnameIPF.Text, NameIPF.Text, LastNameIPF.Text, maskedTextBox1.Text);
            UpdateEventsGrid();
            MessageBox.Show("Клиент успешно добавлен");
        }

        private void AddRespon_Click(object sender, EventArgs e)
        {
            if (SurnameIPF.Text == "" || NameIPF.Text == "" || LastNameIPF.Text == "" || maskedTextBox1.Text == "")
            {
                MessageBox.Show("Заполните поле");
                return;
            }
            functions.AddClient_responsible("Responsible", SurnameIPF.Text, NameIPF.Text, LastNameIPF.Text, maskedTextBox1.Text);
            UpdateEventsGrid();
            MessageBox.Show("Сотрудник успешно добавлен");
        }

        private void EditClient_Click(object sender, EventArgs e)
        {
            functions.EditClientResponsible("Clients", SurnameIPF.Text, NameIPF.Text, LastNameIPF.Text, maskedTextBox1.Text, int.Parse(dataGridView2.CurrentRow.Cells[0].Value.ToString()));
            UpdateEventsGrid();
            MessageBox.Show("Запись о клиенте изменена");
        }

        private void EditRespon_Click(object sender, EventArgs e)
        {
            functions.EditClientResponsible("Responsible", SurnameIPF.Text, NameIPF.Text, LastNameIPF.Text, maskedTextBox1.Text, int.Parse(dataGridView3.CurrentRow.Cells[0].Value.ToString()));
            UpdateEventsGrid();
            MessageBox.Show("Запись о сотруднике изменена");
        }

        private void DeleteClient_Click(object sender, EventArgs e)
        {
            
            
            if(functions.DeleteFromGridByIndex("Clients", int.Parse(dataGridView2.CurrentRow.Cells[0].Value.ToString())))
            MessageBox.Show("Запись о клиенте удалена");
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
            UpdateEventsGrid();
        }

        private void DeleteRespon_Click(object sender, EventArgs e)
        {
            if(functions.DeleteFromGridByIndex("Responsible", int.Parse(dataGridView3.CurrentRow.Cells[0].Value.ToString())))
            MessageBox.Show("Запись о сотруднике удалена");
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
            UpdateEventsGrid();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            SurnameIPF.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
            NameIPF.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
            LastNameIPF.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
            maskedTextBox1.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            SurnameIPF.Text = dataGridView3.CurrentRow.Cells[1].Value.ToString();
            NameIPF.Text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
            LastNameIPF.Text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
            maskedTextBox1.Text = dataGridView3.CurrentRow.Cells[4].Value.ToString();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" || textBox4.Text == "")
            {
                MessageBox.Show("Заполните поле");
                return;
            }
            functions.AddAdress(textBox3.Text,textBox4.Text,int.Parse(comboBox5.SelectedValue.ToString()));
            UpdateEventsGrid();
            MessageBox.Show("Адрес добавлен");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            functions.EditAdress(textBox3.Text, textBox4.Text, int.Parse(comboBox5.SelectedValue.ToString()),int.Parse(dataGridView4.CurrentRow.Cells[0].Value.ToString()));
            UpdateEventsGrid();
            MessageBox.Show("Адрес изменен");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (functions.DeleteFromGridByIndex("Adresses", int.Parse(dataGridView4.CurrentRow.Cells[0].Value.ToString())))
                MessageBox.Show("Адрес удален");
            else MessageBox.Show("Невозможно удалить запись из-за наличия привязаных записей");
            UpdateEventsGrid();
        }


        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null) return;
            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document document =
                    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            //add_text
            document.Content.SetRange(0, 0);
            document.Content.Text =
                "ID записи - " + dataGridView1.CurrentRow.Cells[0].Value.ToString() + Environment.NewLine +
                "Название - " + dataGridView1.CurrentRow.Cells[1].Value.ToString() + Environment.NewLine +
                "Клиент - " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + Environment.NewLine +
                "Ответственный за проведение - " + dataGridView1.CurrentRow.Cells[3].Value.ToString() + Environment.NewLine +
                "Адрес - " + dataGridView1.CurrentRow.Cells[4].Value.ToString() + Environment.NewLine +
                "Доп информация - " + dataGridView1.CurrentRow.Cells[5].Value.ToString() + Environment.NewLine +
                "Дата проведения - " + dataGridView1.CurrentRow.Cells[6].Value.ToString() + Environment.NewLine;

            winword.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable All_products = new DataTable();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            //Книга.




            if (dateTimePicker2.Text == "" || dateTimePicker3.Text == "")
            {
                ExcelApp.Cells[1, 1] = "Данные об всех отчетах";
                All_products = functions.SelectAllFromEvents();
            }

            else
            {
                if (dateTimePicker2.Value < dateTimePicker3.Value)
                {
                    All_products = functions.Take_all_records_by_date(dateTimePicker2, dateTimePicker3);
                    ExcelApp.Cells[1, 1] = "Данные об отчетах в период:" + dateTimePicker2.Value + " - " + dateTimePicker3.Value;
                }
                else
                {
                    MessageBox.Show("Неверный промежуток дат"); return;
                }
            }
            ExcelApp.Columns.NumberFormat = "General";
            ExcelApp.Cells[3, 1] = "ID";
            ExcelApp.Columns[1].ColumnWidth = 5;

            ExcelApp.Cells[3, 2] = "Название";
            ExcelApp.Columns[2].ColumnWidth = 30;

            ExcelApp.Cells[3, 3] = "Клиент";
            ExcelApp.Columns[3].ColumnWidth = 30;

            ExcelApp.Cells[3, 4] = "Ответственный";
            ExcelApp.Columns[4].ColumnWidth = 30;
            ExcelApp.Columns.NumberFormat = "@";

            ExcelApp.Cells[3, 5] = "Адрес";
            ExcelApp.Columns[5].ColumnWidth = 30;

            ExcelApp.Cells[3, 6] = "Доп информация";
            ExcelApp.Columns[6].ColumnWidth = 20;

            ExcelApp.Cells[3, 7] = "Дата";
            ExcelApp.Columns[7].ColumnWidth = 20;
            for (int i = 0; i < All_products.Rows.Count; i++)
            {
                ExcelApp.Cells[i + 4, 1] = All_products.Rows[i][0].ToString(); //id
                ExcelApp.Cells[i + 4, 2] = All_products.Rows[i][1].ToString();//name
                ExcelApp.Cells[i + 4, 3] = All_products.Rows[i][2].ToString();//discr
                ExcelApp.Cells[i + 4, 4] = All_products.Rows[i][3].ToString();//cost
                ExcelApp.Cells[i + 4, 5] = All_products.Rows[i][4].ToString();//id author
                ExcelApp.Cells[i + 4, 6] = All_products.Rows[i][5].ToString();//id author
                ExcelApp.Cells[i + 4, 7] = All_products.Rows[i][6].ToString();//id author
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView4.CurrentRow.Cells[2].Value.ToString();
            textBox4.Text = dataGridView4.CurrentRow.Cells[3].Value.ToString();
            comboBox5.Text = dataGridView4.CurrentRow.Cells[1].Value.ToString();
        }
    }
}
