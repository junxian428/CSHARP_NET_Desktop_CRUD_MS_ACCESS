using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms;
using System.Data.OleDb;


namespace C_Sharp_GUI_App
{
    public partial class Form1 : Form
    {
        private OleDbConnection connection = new OleDbConnection();

        public Form1()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\USer\source\repos\D210044A_Ho_Weng_Yin\D210044A_Ho_Weng_Yin\DatabaseForVB.mdb; Persist Security Info= False;";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                //MessageBox.Show("MS Access is Opened");
                connection.Close();
            } catch(Exception ex)
            {
                MessageBox.Show("Error :" + ex);
            }
            // Combo Box
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                //(Staff_Name, Staff_Age,Staff_Salary)
                //Staff_Name_Box, Staff_Age_Box, Staff_Salary_Box
                //string query = "Delete from Staff_Table " + " WHERE ID=" + ID_BOX.Text + ";";
                string query = "SELECT * FROM Staff_Table";
                command.CommandText = query;
                Console.WriteLine(query);
                //
                //command.ExecuteNonQuery();
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader["Staff_Name"].ToString());
                }
                MessageBox.Show("Data Read Successfully");
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = "INSERT INTO Staff_Table (Staff_Name, Staff_Age,Staff_Salary) VALUES ('" + Staff_Name_Box.Text + "','" + Staff_Age_Box.Text + "','" + Staff_Salary_Box.Text + "');";
                command.ExecuteNonQuery();
                MessageBox.Show("Data Saved");
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }
           
            //command.CommandText = "SELECT * from Staff_Table";
            //Staff_Name_Box, Staff_Age_Box, Staff_Salary_Box
            //Console.WriteLine(command.CommandText);
            //OleDbDataReader reader = command.ExecuteReader();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // Update
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                //(Staff_Name, Staff_Age,Staff_Salary)
                //Staff_Name_Box, Staff_Age_Box, Staff_Salary_Box
                string query = "Update Staff_Table set Staff_Name='" + Staff_Name_Box.Text + "',Staff_Age='" + Staff_Age_Box.Text+"',Staff_Salary='" + Staff_Salary_Box.Text + "' WHERE ID=" + ID_BOX.Text + ";" ;
                command.CommandText = query;
                Console.WriteLine(query);
                command.ExecuteNonQuery();
                MessageBox.Show("Data Updated Successfully");
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Delete
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                //(Staff_Name, Staff_Age,Staff_Salary)
                //Staff_Name_Box, Staff_Age_Box, Staff_Salary_Box
                string query = "Delete from Staff_Table " + " WHERE ID=" + ID_BOX.Text + ";";
                command.CommandText = query;
                Console.WriteLine(query);
                command.ExecuteNonQuery();
                MessageBox.Show("Data Delete Successfully");
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                //(Staff_Name, Staff_Age,Staff_Salary)
                //Staff_Name_Box, Staff_Age_Box, Staff_Salary_Box
                //string query = "Delete from Staff_Table " + " WHERE ID=" + ID_BOX.Text + ";";
                string query = "SELECT * FROM Staff_Table WHERE Staff_Name ='" + comboBox1.Text +"';";
                command.CommandText = query;
                Console.WriteLine(query);
                //
                //command.ExecuteNonQuery();
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    ID_BOX.Text = reader["ID"].ToString();
                    Staff_Name_Box.Text = reader["Staff_Name"].ToString();
                    Staff_Age_Box.Text = reader["Staff_Age"].ToString();
                    Staff_Salary_Box.Text = reader["Staff_Salary"].ToString();

                }
                MessageBox.Show("Data Read Successfully");
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex);
            }

        }
    }
}
