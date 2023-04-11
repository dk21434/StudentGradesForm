using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1._0
{
    public partial class Form1 : Form
    {
        localws.WebService1 ws = new localws.WebService1();


        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

       
        private void button2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = ws.GetAverage();
            dataGridView1.DataSource = dt;
        }

        
        private void button3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = ws.GetAverage();
            dataGridView1.SelectAll();
           
            DataObject obj = dataGridView1.GetClipboardContent();
            Clipboard.SetDataObject(obj);


            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlwb;
            Microsoft.Office.Interop.Excel.Worksheet xlws;
            Microsoft.Office.Interop.Excel.Range xlr;

            Object mv = System.Reflection.Missing.Value;

            
            xlwb = xls.Workbooks.Add(mv);
            xlws = xlwb.Worksheets.get_Item(1);
            xlr = xlws.Cells[1, 1];
            xlr.Select();
            xlws.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            xls.Columns.AutoFit();
            xls.Visible = true;




        }

       
    }
}
