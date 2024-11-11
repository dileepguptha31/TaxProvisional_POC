using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace ItrCalc
{
    public partial class Form1 : Form
    {
        List<string> selectedInputFiles = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder str = new StringBuilder();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;

            // Load the Excel file
            if(selectedInputFiles.Count < 1)
            {
                return;
            }

            var fileInfo = new FileInfo(selectedInputFiles.First());

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
            if (selectedInputFiles.Count < 1)
            {
                return;
            }

            StringBuilder str = new StringBuilder();
            string filePath = selectedInputFiles.First();  // Specify your Excel file path

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
            if (selectedInputFiles.Count < 1)
            {
                return;
            }

            StringBuilder str = new StringBuilder();
            string filePath = selectedInputFiles.First();  // Specify your Excel file path
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

                        dr["EntryType"] = "Received";

                        dataTable.Rows.Add(dr);
                    }
                    // Newline after each row
                }

                // Add provisional rows
                var receivedSalaryRows = dataTable.AsEnumerable()
                                .Where(row => row.Field<string>("EntryType") == "Received" && row.Field<string>("Type")=="SALARY");
                var receivedSalaryMonthsCount  = receivedSalaryRows.Count();
                if (receivedSalaryRows.Any() && receivedSalaryRows.Count() < 12 /*If credited for 12 months, no need to add provisional*/)
                {
                    var lastSalaryCreditRow  = receivedSalaryRows.Last();
                    var lastCreditedMonth = lastSalaryCreditRow["Month"].ToString();
                    var months = GetMonths();
                    for (int provisonalRowIndex = receivedSalaryMonthsCount+1; provisonalRowIndex <= 12; provisonalRowIndex++)
                    {
                        var previousMonthIndex = months.IndexOf(lastCreditedMonth);
                        var currentMonthIndex = previousMonthIndex == 11 ? 0 : previousMonthIndex + 1;
                        var currentMonth = months[currentMonthIndex];
                        var provRow = dataTable.NewRow();
                        provRow.ItemArray =  lastSalaryCreditRow.ItemArray.Clone() as object[];
                        provRow["S No."] = $"{dataTable.Rows.Count + 1}.";
                        provRow["Month"] = currentMonth;
                        provRow["EntryType"] = "Provisional";
                        dataTable.Rows.Add( provRow );

                        //prepare for next row
                        lastSalaryCreditRow = provRow;
                        lastCreditedMonth = currentMonth;
                    }
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

            //MessageBox.Show(str.ToString());
            using (provisonalDataDialog dialog = new provisonalDataDialog())
            {
                dialog.SetDataTable(dataTable); 
                dialog.ShowDialog();
            }
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
            dt.Columns.Add("EntryType");
            return dt;
        }

        private List<string> GetMonths()
        {
            List<string> list = new List<string>();
            list.Add("Jan");
            list.Add("Feb");
            list.Add("Mar");
            list.Add("Apr");
            list.Add("May");
            list.Add("June");
            list.Add("July");
            list.Add("Aug");
            list.Add("Sept");
            list.Add("Oct");
            list.Add("Nov");
            list.Add("Dec");
            return list;
        }

        private void importFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx";
            openFileDialog.FilterIndex = 0;
            openFileDialog.Multiselect = true;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                footerStatusValue.Text = "No files selected";
                return;
            }

            selectedInputFiles = openFileDialog.FileNames.ToList<string>();
            if (selectedInputFiles.Count > 1)
            {
                footerStatusValue.Text = "Multiple files selected1";
            }
            else
            {
                footerStatusValue.Text = selectedInputFiles.First();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            List<Person> people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30, Gender = "Male", City = "New York" },
            new Person { Name = "Jane Smith", Age = 25, Gender = "Female", City = "Los Angeles" },
            new Person { Name = "Samuel Green", Age = 35, Gender = "Male", City = "Chicago" }
        };
            string filePath = "C:\\Users\\kiran\\OneDrive\\Desktop\\ITR\\people.xlsx";

            // Call the method to write the data into Excel
            WriteDataToExcel(people, filePath);
        }
        //static void WriteDataToExcel(List<Person> people, string filePath)
        //{
        //    // Create an Excel application object
        //    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    if (excelApp == null)
        //    {
        //        Console.WriteLine("Excel is not installed properly!");
        //        return;
        //    }

        //    // Create a new workbook and worksheet
        //    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();
        //    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
        //    worksheet.Name = "People";  // Set the worksheet name

        //    // Set the column headers
        //    worksheet.Cells[1, 1] = "Name";
        //    worksheet.Cells[1, 2] = "Age";
        //    worksheet.Cells[1, 3] = "Gender";
        //    worksheet.Cells[1, 4] = "City";

        //    Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range["A1:F2"];
        //    headerRange.Merge();
        //    headerRange.Font.Size = 22;
        //    headerRange.Font.FontStyle = System.Drawing.FontStyle.Bold;
        //    headerRange.Font.Name = "Calibri";
        //    headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;  // Center align text horizontally
        //    headerRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        //    headerRange.Value = "Basic Information for Income Tax computation As per  HRMS";

        //    // Write data to Excel starting from the second row
        //    for (int i = 0; i < people.Count; i++)
        //    {
        //        var person = people[i];
        //        worksheet.Cells[i + 2, 1] = person.Name;
        //        worksheet.Cells[i + 2, 2] = person.Age;
        //        worksheet.Cells[i + 2, 3] = person.Gender;
        //        worksheet.Cells[i + 2, 4] = person.City;
        //    }

        //    // Save the workbook to the specified file path
        //    workbook.SaveAs(filePath);

        //    // Close the workbook and quit the Excel application
        //    workbook.Close(false);
        //    excelApp.Quit();

        //    // Release COM objects to prevent memory leaks
        //    Marshal.ReleaseComObject(worksheet);
        //    Marshal.ReleaseComObject(workbook);
        //    Marshal.ReleaseComObject(excelApp);
        //}

        static void WriteDataToExcel<T>(List<T> data, string filePath)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(typeof(T).Name);

                // Get class properties
                var properties = typeof(T).GetProperties();

                var range = worksheet.Cells["A1:F2"];
                range.Merge = true;  // Merge the cells

                // Set the font size for the merged cells
                range.Style.Font.Size = 22;  // Set the font size to 16 (or your desired size)
                range.Style.Font.Bold = true;
                range.Value = "Basic Information for Income Tax computation As per  HRMS";

                // Write headers
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[3, i + 1].Value = properties[i].Name;
                }

                // Write data
                for (int i = 0; i < data.Count; i++)
                {
                    for (int j = 0; j < properties.Length; j++)
                    {
                        worksheet.Cells[i + 3, j + 1].Value = properties[j].GetValue(data[i]);
                    }
                }

                // Save the file
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
        }
    }
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Gender { get; set; }
        public string City { get; set; }
    }
}
