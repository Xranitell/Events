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
        }



        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateEventsGrid();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt =  new DataTable();
            dt = functions.SelectInComboboxAdress(int.Parse(comboBox2.SelectedValue.ToString()));
            comboBox4.DataSource = dt;
            comboBox4.ValueMember = "Id";
            comboBox4.DisplayMember = "Adress";
        }

        private void AddEvent_Click(object sender, EventArgs e)
        {
            functions.AddEvent(textBox1.Text, int.Parse(comboBox1.SelectedValue.ToString()), int.Parse(comboBox2.SelectedValue.ToString()),int.Parse(comboBox4.SelectedValue.ToString()),textBox2.Text,dateTimePicker1.Value.ToString("yyyy-MM-dd"));
            MessageBox.Show("Запись успешно добавлена!");
            UpdateEventsGrid();
        }
    }
}
