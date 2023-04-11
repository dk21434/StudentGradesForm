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

       
        private void button2_Click(object sender, EventArgs e) // Button click event handler to display the average grades in a data grid view
        {
            System.Data.DataTable dt = new System.Data.DataTable(); // Call the web service method to get the average grades
            dt = ws.GetAverage();
            dataGridView1.DataSource = dt;  

        }


        private void button3_Click(object sender, EventArgs e)  // Button click event handler to export the data grid view to Microsoft Excel
        {
            System.Data.DataTable dt = new System.Data.DataTable(); // Call the web service method to get the average grades
            dt = ws.GetAverage();

            CopyDataGridViewToClipboard(dataGridView1); // Copy the data grid view to the clipboard for use in Excel



            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application(); // Create a new instance of Microsoft Excel and add a new workbook
            Microsoft.Office.Interop.Excel.Workbook xlwb;
            Microsoft.Office.Interop.Excel.Worksheet xlws;
            Microsoft.Office.Interop.Excel.Range xlr;
            Object mv = System.Reflection.Missing.Value;

            xlwb = xls.Workbooks.Add(mv);
            xlws = xlwb.Worksheets.get_Item(1);
            xlr = xlws.Cells[1, 1];
            xlr.Select();
            xlws.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true); // Paste the clipboard data into Excel
            xls.Columns.AutoFit();
            xls.Visible = true;

        }

        private void CopyDataGridViewToClipboard(DataGridView dataGridView) // Utility method to copy the data grid view to the clipboard
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
