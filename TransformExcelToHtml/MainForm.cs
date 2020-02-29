using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TransformExcelToHtml
{
    public partial class MainForm : Form
    {
        private string excelFilePath;
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            excelFilePath = openFileDialog1.FileName;
            textBox1.Text = openFileDialog1.SafeFileName;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;
            Workbook wb = excelApp.Workbooks.Open(excelFilePath);

            // Loop ws throught wb
            IEnumerator wsEnumerator = excelApp.ActiveWorkbook.Worksheets.GetEnumerator();
            object missing = Type.Missing;
            object format = Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml;
            //while (wsEnumerator.MoveNext())
            //{
                Workbook wsCurrent = wb;
                String outputFile = excelFilePath + ".html";
                //wsCurrent.SaveAs(outputFile, format, missing, missing, missing,
                //    missing, XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                PublishObjects publish = wsCurrent.PublishObjects;
                System.Diagnostics.Debug.WriteLine("outputFile: " + outputFile);
                System.Diagnostics.Debug.WriteLine("XlSourceType.xlSourcePrintArea: " + XlSourceType.xlSourcePrintArea);
                publish.Add(
                    XlSourceType.xlSourcePrintArea,
                    outputFile,
                    "default",
                    XlSourceType.xlSourcePrintArea,
                    XlHtmlType.xlHtmlStatic,
                    "",
                    ""
                ).Publish(true);
            //}
            //excelApp.Quit();
            MessageBox.Show("Done", "TETH", MessageBoxButtons.OK);
        }
    }
}
