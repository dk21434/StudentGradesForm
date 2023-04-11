using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System;
using System.Data;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1._0
{
    public partial class Form1 : Form
    {
        wService.WebService1 ws = new wService.WebService1();


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

            CopyDataGridViewToClipboard(dataGridView1);



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

        private void CopyDataGridViewToClipboard(DataGridView dataGridView)
        {
            dataGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView.SelectAll();

            DataObject dataObj = dataGridView.GetClipboardContent();
            if (dataObj != null)
            {
                Clipboard.SetDataObject(dataObj);
            }

            // Clear the selection if you don't want the cells to remain selected
            dataGridView.ClearSelection();
        }

        
    }
}
