using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ItrCalc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder str = new StringBuilder();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

            // Load the Excel file
            var fileInfo = new FileInfo(@"C:\Users\kiran\OneDrive\Desktop\ITR\kiran.xlsx");

            using (var package = new ExcelPackage(fileInfo))
            {
                // Get the first worksheet
                var worksheet = package.Workbook.Worksheets[0];

                // Loop through rows and columns
                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        str.Append($"{worksheet.Cells[row, col].Text}");
                    }
                }
            }
            MessageBox.Show(str.ToString());
        }

        private void npoi_Click(object sender, EventArgs e)
        {
            StringBuilder str = new StringBuilder();
            string filePath = "C:\\Users\\kiran\\OneDrive\\Desktop\\ITR\\Paybilldetails40.xls";  // Specify your Excel file path

            // Open the Excel file
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Create a workbook instance from the .xls file
                HSSFWorkbook workbook = new HSSFWorkbook(fileStream);  // HSSFWorkbook for .xls files

                // Get the first sheet in the workbook
                ISheet sheet = workbook.GetSheetAt(0);  // Sheet at index 0 (first sheet)

                // Iterate through rows and cells to read data
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null) continue;  // Skip empty rows

                    for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex);
                        if (cell == null) continue;  // Skip empty cells

                        // Print the value of each cell in the row
                        str.Append(cell.ToString() + "\t");
                    }
                    str.AppendLine();
                }
            }
            MessageBox.Show(str.ToString());
        }

        private void ms_Click(object sender, EventArgs e)
        {
            StringBuilder str = new StringBuilder();
            string filePath = "C:\\Users\\kiran\\OneDrive\\Desktop\\ITR\\Paybilldetails (38).xls";  // Specify your Excel file path
            var dataTable = CreateDataTableMetaData();
            // Initialize Excel application
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;  // Do not show Excel UI

            try
            {
                // Open the workbook
                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];  // Access the first sheet

                // Get the range of used cells
                Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                bool dataonly = false;
                int dataHeader;
                // Loop through rows and columns
                for (int row = 1; row <= rowCount; row++)
                {
                   
                    if(dataonly || ((Microsoft.Office.Interop.Excel.Range)usedRange.Cells[row, 1]).Value2 == "S No.")
                    {
                        if (!dataonly)
                        {
                            dataonly = true;
                            dataHeader = row;
                            continue;
                        }                       
                        DataRow dr = dataTable.NewRow();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cell = (Microsoft.Office.Interop.Excel.Range)usedRange.Cells[row, col];
                            dr[dataTable.Columns[col - 1].ColumnName] = cell.Value2;
                            str.Append(cell.Value2 + "\t");  // Print the cell value
                        }

                        if (dr.IsNull(0) && dr.IsNull(1) && dr.IsNull(3) && dr.IsNull(4))
                            continue;

                        dataTable.Rows.Add(dr);
                    }
                    // Newline after each row
                }

                // Close workbook
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Quit Excel application
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            MessageBox.Show(str.ToString());
        }

        private void LoadDatafromExcel()
        {
            var dataTable = CreateDataTableMetaData();
        }

        private DataTable CreateDataTableMetaData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("S No.");
            dt.Columns.Add("KGID No");
            dt.Columns.Add("Type");
            dt.Columns.Add("Month");
            dt.Columns.Add("Year");
            dt.Columns.Add("Paybill Generation Date");
            dt.Columns.Add("DDO Code");
            dt.Columns.Add("Paybill NO");
            dt.Columns.Add("Token No");
            dt.Columns.Add("Employee Name");
            dt.Columns.Add("Designation");
            dt.Columns.Add("Metal No");
            dt.Columns.Add("PAN No");
            dt.Columns.Add("TAN No");
            dt.Columns.Add("PayScale");
            dt.Columns.Add("Bill Unit No");
            dt.Columns.Add("Basic Pay");
            dt.Columns.Add("Stagnation Increment");
            dt.Columns.Add("DA");
            dt.Columns.Add("HRA");
            dt.Columns.Add("Special Pay");
            dt.Columns.Add("Uniform Allowance");
            dt.Columns.Add("Independent Charge Allowance");
            dt.Columns.Add("Medical Allowance");
            dt.Columns.Add("Personal Pay");
            dt.Columns.Add("Other Allowances");
            dt.Columns.Add("Gross Allowance");
            dt.Columns.Add("Income Tax");
            dt.Columns.Add("EGIS");
            dt.Columns.Add("PT");
            dt.Columns.Add("LIC");
            dt.Columns.Add("Nps Deduction Amount");
            dt.Columns.Add("Nps Recovery Amount");
            dt.Columns.Add("KGID");
            dt.Columns.Add("GPF");
            dt.Columns.Add("GPF Loan");
            dt.Columns.Add("KGID Loan");
            dt.Columns.Add("Festival Advance");
            dt.Columns.Add("Advance Pay");
            dt.Columns.Add("HBA");
            dt.Columns.Add("Motor Cycle Advance");
            dt.Columns.Add("Housing Development Finance Corporation");
            dt.Columns.Add("Recovery of Over Payment");
            dt.Columns.Add("Arogya Bhagya Yojana");
            dt.Columns.Add("Msil");
            dt.Columns.Add("Electricity");
            dt.Columns.Add("Co-operative Society");
            dt.Columns.Add("Gross Recovery");
            dt.Columns.Add("Gross Deduction");
            dt.Columns.Add("Gross Salary");
            dt.Columns.Add("Net Salary");
            dt.Columns.Add("Bank A/C No");
            dt.Columns.Add("Name Of The Bank");
            dt.Columns.Add("Name Of The Bank Branch");
            return dt;
        }        
    }
}
