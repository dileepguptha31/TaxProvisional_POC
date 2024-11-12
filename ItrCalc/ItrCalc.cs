using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace ItrCalc
{
    public partial class ItrCalc : Form
    {
        public ItrCalc()
        {
            InitializeComponent();
        }

        private void Load_Click(object sender, EventArgs e)
        {
            var errorCount = 0;
            if (string.IsNullOrEmpty(txtPath.Text))
            {
                MessageBox.Show("No Folder Path Selected or No files to Process");
            }
            else
            {
                var filesList = Directory.GetFiles(txtPath.Text);
                if (filesList.Count() == 0)
                {
                    MessageBox.Show("No files to Process");
                    return;
                }

                if(!Directory.Exists(txtPath.Text + "\\Processed"))
                {
                    Directory.CreateDirectory(txtPath.Text + "\\Processed");
                }
                if (!Directory.Exists(txtPath.Text + "\\Errors"))
                {
                    Directory.CreateDirectory(txtPath.Text + "\\Errors");
                }
                if (!Directory.Exists(txtPath.Text + "\\OutPut"))
                {
                    Directory.CreateDirectory(txtPath.Text + "\\OutPut");
                }

                var noFiles = filesList.Count();
                foreach (string inputFile in filesList)
                {
                    var fileExt = Path.GetExtension(inputFile);
                    var dataInput = LoadExcelfromMicrosoftInterop(inputFile);
                    if(dataInput != null)
                    {
                        var ProcessedfileName = txtPath.Text + "\\Processed\\" + Path.GetFileNameWithoutExtension(inputFile) + "_" + DateTime.Now.ToString("ddMMyyyy")  + fileExt;
                        File.Move(inputFile, ProcessedfileName);
                        ComputeAndCreateFinalAggregratedOutput(dataInput, txtPath.Text + "\\OutPut");
                    }
                    else
                    {
                        errorCount = errorCount + 1;
                        File.Move(inputFile, txtPath.Text + "\\Errors");
                    }
                }
            }
        }

        public DataTable LoadExcelfromMicrosoftInterop(string filePath)
        {
            var dtInputData = CreateDataTableMetaData();

            var excelApp = new Application();
            excelApp.Visible = false;  // Do not show Excel UI

            try
            {
                // Open the workbook
                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet worksheet = (Worksheet)workbook.Sheets[1];  // Access the first sheet

                // Get the range of used cells
                Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                bool dataonly = false;
                int dataHeader;
                // Loop through rows and columns
                for (int row = 1; row <= rowCount; row++)
                {

                    if (dataonly || ((Range)usedRange.Cells[row, 1]).Value2 == "S No.")
                    {
                        if (!dataonly)
                        {
                            dataonly = true;
                            dataHeader = row;
                            continue;
                        }
                        DataRow dr = dtInputData.NewRow();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cell = (Microsoft.Office.Interop.Excel.Range)usedRange.Cells[row, col];
                            dr[dtInputData.Columns[col - 1].ColumnName] = cell.Value2;
                        }

                        if ((dr.IsNull(0) && dr.IsNull(1) && dr.IsNull(3) && dr.IsNull(4)) || dr[0].Equals("Totals"))
                            continue;

                        dr["EntryType"] = "Received";

                        dtInputData.Rows.Add(dr);
                    }
                }
                // Close workbook
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                lblStatus.Text = lblStatus.Text + "Processed Sucess : " + Path.GetFileName(filePath);
                return dtInputData;
            }
            catch (Exception ex)
            {
                lblStatus.Text = lblStatus.Text + "Error occurred " + Path.GetFileName(filePath);
            }
            finally
            {
                // Quit Excel application
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            lblStatus.Visible = true;
            return null;
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
            return new List<string>
            {
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "June",
                "July",
                "Aug",
                "Sept",
                "Oct",
                "Nov",
                "Dec"
            };
        }
        private bool ComputeAndCreateFinalAggregratedOutput(DataTable dtInputData, string fileoutpath)
        {

            var distinctPanNo = dtInputData.AsEnumerable().Select(x => x.Field<string>("PAN No")).Distinct().ToList();

            foreach (var panNo in distinctPanNo) 
            {
                var dtPanSpecific = dtInputData.AsEnumerable().Where(x => x.Field<string>("PAN No") == panNo).CopyToDataTable();

                var updatedProvisionalData = UpdateProvisionalData(dtPanSpecific);

                AggregratedOutput(updatedProvisionalData, fileoutpath);
            }

            return true;
        }
        private decimal calculateSum(EnumerableRowCollection<DataRow> data, string paytype, string ddoNo = "")
        {
            if (string.IsNullOrEmpty(ddoNo))
            {
                return data.Sum(sal => Convert.ToDecimal(sal.Field<string>(paytype)));
            }
            else
            {
                return data.Where(row => row.Field<string>("DDO Code").Equals(ddoNo)).Sum(sal => Convert.ToDecimal(sal.Field<string>(paytype)));
            }
        }
        private void AggregratedOutput(DataTable updatedProvisionalData,string fileoutputPath)
        {
            var results = new Dictionary<string, (decimal total, decimal currentEmployerTotal, decimal cumulativeTotal)>();

            var data = updatedProvisionalData.AsEnumerable();
            var areMultipleEmployers = data.GroupBy(row => row.Field<string>("DDO Code")).Count();
            var lastData = data.Last();
            var currentEmployer = lastData["DDO Code"].ToString();
            var kgidno = lastData["KGID No"];
            var employeeName = lastData["Employee Name"];
            var panNo = lastData["PAN No"];
            var designation = lastData["Designation"];

            var finalFilename = fileoutputPath + "\\" + employeeName.ToString().Replace(" ", "") + "_" + kgidno + "_" + panNo+".xlsx";
            // Define an array with all field names
            string[] fields = new[]
            {
                "Basic Pay", "Stagnation Increment", "DA", "HRA", "Special Pay",
                "Uniform Allowance", "Independent Charge Allowance", "Medical Allowance",
                "Personal Pay", "Other Allowances", "Gross Allowance", "Income Tax",
                "EGIS", "PT", "LIC", "Nps Deduction Amount", "Nps Recovery Amount",
                "KGID", "GPF", "GPF Loan", "KGID Loan", "Festival Advance",
                "Advance Pay", "HBA", "Motor Cycle Advance",
                "Housing Development Finance Corporation", "Recovery of Over Payment",
                "Arogya Bhagya Yojana", "Msil", "Electricity",
                "Co-operative Society", "Gross Recovery", "Gross Deduction",
                "Gross Salary", "Net Salary"
            };


            foreach (var field in fields)
            {
                var total = calculateSum(data, field);

                decimal currentEmployerTotal = calculateSum(data, field, currentEmployer);

                decimal cumulativeTotal = total - currentEmployerTotal;

                results[field] = (total, currentEmployerTotal, cumulativeTotal);
            }

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("HRMS Data Input");


                //var range = worksheet.Cells["A1:F2"];
                //range.Merge = true;  // Merge the cells

                //// Set the font size for the merged cells
                //range.Style.Font.Size = 22;  // Set the font size to 16 (or your desired size)
                //range.Style.Font.Bold = true;
                //range.Value = "Basic Information for Income Tax computation As per  HRMS";

                // Write headers
                for (int i = 0; i < updatedProvisionalData.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = updatedProvisionalData.Columns[i].ColumnName;
                }
                // Make the first row bold
                using (var range = worksheet.Cells[1, 1, 1, updatedProvisionalData.Columns.Count])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Size = 12;
                }


                // Write data rows
                for (int row = 0; row < updatedProvisionalData.Rows.Count; row++)
                {
                    for (int col = 0; col < updatedProvisionalData.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = updatedProvisionalData.Rows[row][col];
                    }
                }
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                var worksheet2 = package.Workbook.Worksheets.Add("Basic Information Details");
                var sheet2range = worksheet2.Cells["A1:D2"];
                sheet2range.Merge = true;  // Merge the cells
                sheet2range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet2range.Style.Fill.BackgroundColor.SetColor(Color.LightCyan);

                // Set the font size for the merged cells
                sheet2range.Style.Font.Size = 16;  // Set the font size to 16 (or your desired size)
                sheet2range.Style.Font.Bold = true;
                sheet2range.Style.ShrinkToFit = true;
                sheet2range.Value = "Basic Information for Income Tax computation As per  HRMS";

                worksheet2.Cells["A3"].Value = "KGID No";
                worksheet2.Cells["A3"].Style.Font.Bold = true;
                var kgidRange = worksheet2.Cells["B3:D3"];
                kgidRange.Merge = true;  // Merge the cells
                kgidRange.Value = kgidno;

                worksheet2.Cells["A4"].Value = "Employee Name";
                worksheet2.Cells["A4"].Style.Font.Bold = true;
                
                var employeeNameRange = worksheet2.Cells["B4:D4"];
                employeeNameRange.Merge = true;  // Merge the cells
                employeeNameRange.Value = employeeName;

                worksheet2.Cells["A5"].Value = "Designation";
                worksheet2.Cells["A5"].Style.Font.Bold = true;
                worksheet2.Cells["B5"].Value = designation;


                var designationRange = worksheet2.Cells["B5:D5"];
                designationRange.Merge = true;  // Merge the cells
                designationRange.Value = designation;



                worksheet2.Cells["A6"].Value = "PAN No";
                worksheet2.Cells["A6"].Style.Font.Bold = true;
                worksheet2.Cells["B6"].Value = panNo;
                worksheet2.Cells["C6"].Value = "DDO "+currentEmployer;
                worksheet2.Cells["C6"].Style.Font.Bold = true;
                worksheet2.Cells["D6"].Value = "Other DDO";

                int startingrow = 7;
                foreach (var fieldValues in results)
                {
                    worksheet2.Cells[$"A{startingrow}"].Value = fieldValues.Key;
                    worksheet2.Cells[$"A{startingrow}"].Style.Font.Bold = true;
                    worksheet2.Cells[$"B{startingrow}"].Value = fieldValues.Value.total;
                    worksheet2.Cells[$"C{startingrow}"].Value = fieldValues.Value.currentEmployerTotal;
                    worksheet2.Cells[$"C{startingrow}"].Style.Font.Bold = true;
                    worksheet2.Cells[$"D{startingrow}"].Value = fieldValues.Value.cumulativeTotal;
                    startingrow = startingrow + 1;
                }


                var sheet2rangeBackground = worksheet2.Cells["A3:D41"];
                sheet2rangeBackground.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet2rangeBackground.Style.Fill.BackgroundColor.SetColor(Color.FloralWhite);

                // Set border styles for the range
                sheet2rangeBackground.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet2rangeBackground.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet2rangeBackground.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet2rangeBackground.Style.Border.Right.Style = ExcelBorderStyle.Thin;

                // Set border color (optional)
                Color borderColor = Color.Black;
                sheet2rangeBackground.Style.Border.Top.Color.SetColor(borderColor);
                sheet2rangeBackground.Style.Border.Bottom.Color.SetColor(borderColor);
                sheet2rangeBackground.Style.Border.Left.Color.SetColor(borderColor);
                sheet2rangeBackground.Style.Border.Right.Color.SetColor(borderColor);

                worksheet2.Cells[worksheet.Dimension.Address].AutoFitColumns();
                
                // Save the file
                FileInfo fi = new FileInfo(finalFilename);
                package.SaveAs(fi);
            }
        }

        private DataTable UpdateProvisionalData(DataTable dtPanSpecific)
        {
            var receivedSalaryRows = dtPanSpecific.AsEnumerable()
                               .Where(row => row.Field<string>("EntryType") == "Received" && row.Field<string>("Type") == "SALARY");

            var receivedSalaryMonthsCount = receivedSalaryRows.Count();

            if (receivedSalaryRows.Any() && receivedSalaryRows.Count() < 12)/*If credited for 12 months, no need to add provisional*/
            {
                var lastSalaryCreditRow = receivedSalaryRows.Last();

                var lastCreditedMonth = lastSalaryCreditRow["Month"].ToString();

                var months = GetMonths();

                for (int provisonalRowIndex = receivedSalaryMonthsCount + 1; provisonalRowIndex <= 12; provisonalRowIndex++)
                {
                    var previousMonthIndex = months.IndexOf(lastCreditedMonth);
                    var currentMonthIndex = previousMonthIndex == 11 ? 0 : previousMonthIndex + 1;
                    var currentMonth = months[currentMonthIndex];
                    var provRow = dtPanSpecific.NewRow();

                    provRow.ItemArray = lastSalaryCreditRow.ItemArray.Clone() as object[];
                    provRow["S No."] = $"{dtPanSpecific.Rows.Count + 1}.Provisional";
                    provRow["Month"] = currentMonth;
                    provRow["EntryType"] = "Provisional";
                    dtPanSpecific.Rows.Add(provRow);

                    //prepare for next row
                    lastSalaryCreditRow = provRow;
                    lastCreditedMonth = currentMonth;
                }
            }
            return dtPanSpecific;
        }

    }
}
