using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace Address_Book_DB_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<Person> people = new List<Person>();

        private void label5_Click_1(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Person p = new Person();
            p.Name = textBox1.Text;
            p.PhoneNumber = textBox2.Text;
            p.Email = textBox3.Text;
            p.StreetAddress = textBox4.Text;
            p.AdditionalNotes = textBox5.Text;
            p.Birthday = dateTimePicker1.Value;
            people.Add(p);
            listView1.Items.Add(p.Name);
            
            foreach (ListViewItem Item in listView1.Items)
            {
                try
                {
                    SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename='C:\Users\sreelatha\Documents\Visual Studio 2015\Projects\Address_Book(DB)\Address_Book(DB)\Address_Book.mdf';Integrated Security=True");
                    con.Open();
                    var swl = "INSERT INTO Contacts(Name,MobileNumber,Email,StreetAddress,Birthday,AdditionalNotes) VALUES('" + textBox1.Text + "','" + textBox2.Text + "', '" + textBox3.Text + "', '" + textBox4.Text + "',  '" + dateTimePicker1.Value + "','" + textBox5.Text + "')";
                    SqlCommand cmd = new SqlCommand(swl, con);
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("CONTACT ALREADY EXISTS!!!");
                }
            }
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            dateTimePicker1.Value = DateTime.Now;


        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            people[listView1.SelectedItems[0].Index].Name = textBox1.Text;
            people[listView1.SelectedItems[0].Index].PhoneNumber = textBox2.Text;
            people[listView1.SelectedItems[0].Index].Email = textBox3.Text;
            people[listView1.SelectedItems[0].Index].StreetAddress = textBox4.Text;
            people[listView1.SelectedItems[0].Index].AdditionalNotes = textBox5.Text;
            people[listView1.SelectedItems[0].Index].Birthday = dateTimePicker1.Value;
            listView1.SelectedItems[0].Text = textBox1.Text;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            dateTimePicker1.Value = DateTime.Now;

            
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                textBox1.Text = people[listView1.SelectedItems[0].Index].Name;
                textBox2.Text = people[listView1.SelectedItems[0].Index].PhoneNumber;
                textBox3.Text = people[listView1.SelectedItems[0].Index].Email;
                textBox4.Text = people[listView1.SelectedItems[0].Index].StreetAddress;
                textBox5.Text = people[listView1.SelectedItems[0].Index].AdditionalNotes;
                dateTimePicker1.Value = people[listView1.SelectedItems[0].Index].Birthday;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Remove();
        }
        void Remove()
        {
            try
            {
                listView1.Items.Remove(listView1.SelectedItems[0]);
                people.RemoveAt(listView1.SelectedItems[0].Index);
            }
            catch { }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Remove();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
        }
    }
    class Person
    {
        public string Name
        {
            get;
            set;
        }
        public string PhoneNumber
        {
            get;
            set;
        }
        public string Email
        {
            get;
            set;
        }
        public string StreetAddress
        {
            get;
            set;
        }
        public string AdditionalNotes
        {
            get;
            set;
        }
        public DateTime Birthday
        {
            get;
            set;
        }

    }
}
