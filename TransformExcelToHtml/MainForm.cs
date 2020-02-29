using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
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
            PublishObjects publish = wb.PublishObjects;
            object misValue = System.Reflection.Missing.Value;

            // Loop through excel worksheets
            foreach (Worksheet worksheet in wb.Worksheets)
            {
                // TODO config folder to export html file
                System.Diagnostics.Debug.WriteLine(Path.GetDirectoryName(excelFilePath));
                String outputFile = Path.GetDirectoryName(excelFilePath) + "/" + worksheet.Name + ".html";
                System.Diagnostics.Debug.WriteLine(outputFile);
                publish.Add(
                    XlSourceType.xlSourcePrintArea,
                    outputFile,
                    worksheet.Name,
                    XlSourceType.xlSourcePrintArea,
                    XlHtmlType.xlHtmlStatic,
                    worksheet.Name // id of div tag
                ).Publish(true);
            }
            wb.Close(false, misValue, misValue);
            excelApp.Quit();
            MessageBox.Show("Done", "TETH", MessageBoxButtons.OK);
        }
    }
}
