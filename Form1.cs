using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel; //Referansdan elave etdikden sonra kitabxana daxil edirik

namespace Excel_Process
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application uygulama=new Microsoft.Office.Interop.Excel.Application();
            uygulama.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook kitap = uygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Microsoft.Office.Interop.Excel.Worksheet sayfa1=(Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
            Microsoft.Office.Interop.Excel.Range alan1=(Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1,1];  //Hansi setir ve sutuna daxil edeceyimizi yaziriq
            alan1.Value = textBox1.Text;
            

            Microsoft.Office.Interop.Excel.Range alan2 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1,2];  //Hansi setir ve sutuna daxil edeceyimizi yaziriq
            alan2.Value = textBox2.Text;


            Microsoft.Office.Interop.Excel.Range alan3 = (Microsoft.Office.Interop.Excel.Range)sayfa1.Cells[1,3];  //Hansi setir ve sutuna daxil edeceyimizi yaziriq
            alan3.Value = textBox3.Text;

            string ad = textBox1.Text;
            string soyad = textBox2.Text;
            string mail = textBox3.Text;
            string[] bilgiler = { ad, soyad, mail};
            listView1.Items.Add(new ListViewItem(bilgiler));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.Columns.Add("Ad", 150);
            listView1.Columns.Add("Soyad", 150);
            listView1.Columns.Add("Mail", 150);
        }
    }
}
