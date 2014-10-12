using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Migration
{
    public partial class Form1 : Form
    {
        public string file = "null";

        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.ApplicationClass();
            Excel.Workbook book;
            Excel.Worksheet sheet;
            Excel.Range range;

            if(file != "null")
            {
                book = app.Workbooks.Open(file);
               // sheet = (Excel.Worksheet)book.Worksheets.get_Item(3);
            }
            else {
                Console.WriteLine("File is null");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if(result == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                label8.Text = file;
            }

            Console.WriteLine(result);
        }

    }
}
