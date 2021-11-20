using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace grafobol
{

    public partial class Form1 : Form
    {

        string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = example.accdb";
        OleDbConnection connection;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            connection = new OleDbConnection(connectionString);
            connection.Open();
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            connection.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string commandString = "SELECT [причина ремонта] FROM [R1] WHERE [инвентарный номер] = '1к2'";
            OleDbCommand command = new OleDbCommand(commandString, connection);
            textBox1.Text = command.ExecuteScalar().ToString();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string commandString = "SELECT * FROM [R1]";
            OleDbCommand command = new OleDbCommand(commandString, connection);
            OleDbDataReader reader = command.ExecuteReader();


            while (reader.Read())
            {
                listBox1.Items.Add($"{reader[0].ToString()}\t {reader[1].ToString()}\t {reader[2].ToString()}\t +" +
                    $"{reader[3].ToString()}\t {reader[4].ToString()}\t {reader[0].ToString()}\t");

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {

                string commandString = $"INSERT INTO [R3] ([ФИО], [должность]) VALUES ('{textBox2.Text}', '{textBox3.Text}' )";
                OleDbCommand command = new OleDbCommand(commandString, connection);
                command.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Ошибка!");
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string commandString = $"UPDATE [R3] SET [должность] = '{textBox5.Text}' WHERE [ФИО] = '{textBox4.Text}'";
                OleDbCommand command = new OleDbCommand(commandString, connection);
                command.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }

         
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string commandString = $"DELETE FROM [R3]  WHERE [ФИО] = '{textBox6.Text}'";
                OleDbCommand command = new OleDbCommand(commandString, connection);
                command.ExecuteNonQuery();
            }


            catch
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string commandString = "SELECT * FROM [R1]";

            OleDbDataAdapter adapter = new OleDbDataAdapter(commandString, connection);

            DataTable table = new DataTable();

            adapter.Fill(table);

            dataGridView1.DataSource = table;


        }
    }
}
