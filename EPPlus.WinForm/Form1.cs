using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPPlus.WinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.AddExtension = true;
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";

            DialogResult result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var ep = new ExcelPackage(fs))
                {
                    var ws = ep.Workbook.Worksheets[1];
                    int startRowNumber = ws.Dimension.Start.Row;  //起始列編號，從1算起
                    int endRowNumber = ws.Dimension.End.Row;  //結束列編號，從1算起
                    int startColumn = ws.Dimension.Start.Column;  //開始欄編號，從1算起
                    int endColumn = ws.Dimension.End.Column;  //結束欄編號，從1算起

                    bool hasHeader = true;
                    //有包含標題
                    if (hasHeader)
                        startRowNumber += 1;

                    var cells = ws.Cells[startRowNumber, startColumn, endRowNumber, endColumn];
                }
            }
        }
    }
}
