using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Programa
{
    public partial class Formulario : Form
    {
        public Formulario()
        {
            InitializeComponent();
        }

        private void BotonSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BotonCargar_Click(object sender, EventArgs e)
        {
            LoadFiles();
        }

        private void BotonGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (DataGridView1[0, 1].Value.ToString() != String.Empty)
                    CreateExcelFile();
            }
            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show(this, "File not upload.", "Error", MessageBoxButtons.OK);
            }
        }

        private void LoadFiles()
        {
            string[] dataArray = null;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    RutaArchivo.Text = Path.GetDirectoryName(openFileDialog.FileName);

                    foreach (String fileName in openFileDialog.FileNames)
                    {
                        StreamReader streamReader = new StreamReader(fileName);
                        string stringLine = "";

                        while (stringLine != null)
                        {
                            stringLine = streamReader.ReadLine();

                            if (stringLine != null)
                            {
                                dataArray = stringLine.Split(';');
                                DataGridView1.Rows.Add(dataArray[0], dataArray[1], dataArray[2], dataArray[3], dataArray[4], dataArray[5], dataArray[6], dataArray[7], dataArray[8]);
                            }  
                        }
                        streamReader.Close();
                    }
                }
                catch (IndexOutOfRangeException) 
                {
                    MessageBox.Show(this, "File content not supported.", "Error", MessageBoxButtons.OK);
                }
                
            } 
        }

        private void CreateExcelFile()
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application ExcelApplication = default(Microsoft.Office.Interop.Excel.Application);          
                Microsoft.Office.Interop.Excel.Workbook workBook = default(Microsoft.Office.Interop.Excel.Workbook);
                Microsoft.Office.Interop.Excel.Worksheet workSheet = default(Microsoft.Office.Interop.Excel.Worksheet);

                ExcelApplication = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true
                };

                workBook = ExcelApplication.Workbooks.Add();
                workBook.Worksheets[1].Name = "Report";
                workBook.Worksheets[2].Name = "Pivot Table";
                workBook.Worksheets[3].Name = "To check";

                workSheet = workBook.Worksheets[1];
                workSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;
                workSheet.Activate();

                ExportToExcel(workSheet, DataGridView1);

                Microsoft.Office.Interop.Excel.Range formatRange = workSheet.get_Range("A1", "I1");
                formatRange.Font.Bold = true;
                formatRange.Columns.ColumnWidth = 15;
                workSheet.UsedRange.Columns.Select();

                workBook.SaveAs(saveFileDialog.FileName, XlFileFormat.xlWorkbookNormal,
                                                System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                                                XlSaveAsAccessMode.xlShared, false, false,
                                                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            }
        }

        private void ExportToExcel(Microsoft.Office.Interop.Excel.Worksheet workSheet, DataGridView dataGridView)
        {
            int iCol = 0;

            foreach (DataGridViewColumn column in dataGridView.Columns)
                if (column.Visible)
                    workSheet.Cells[1, ++iCol] = column.HeaderText;

            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                }
            }

        }

        private void CreatePivotTable(Microsoft.Office.Interop.Excel.Worksheet workSheet)
        {

        }
    }
}
