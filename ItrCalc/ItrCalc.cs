using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Color = System.Drawing.Color;
using DataTable = System.Data.DataTable;
using Timer = System.Windows.Forms.Timer;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ItrCalc
{
    public partial class ItrCalc : Form
    {
        private Timer timer;
        private int elapsedTime;
        public ItrCalc()
        {
            InitializeComponent();
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000; // Set the timer interval to 1 second (1000 milliseconds)
            timer.Tick += Timer_Tick;
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            elapsedTime++;
            label4.Text = $"Time: {elapsedTime / 60}";
        }
        private void Load_Click(object sender, EventArgs e)
        {
            label4.Visible = true;
            elapsedTime = 0; // Reset elapsed time
            timer.Start(); // Start the timer
            var errorCount = 0;
            var filesList = Directory.GetFiles(txtPath.Text, "*.xls*", SearchOption.AllDirectories);


            if (!Directory.Exists(txtoutput.Text + "\\Processed"))
            {
                Directory.CreateDirectory(txtoutput.Text + "\\Processed");
            }
            if (!Directory.Exists(txtoutput.Text + "\\Errors"))
            {
                Directory.CreateDirectory(txtoutput.Text + "\\Errors");
            }
            if (!Directory.Exists(txtoutput.Text + "\\OutPut"))
            {
                Directory.CreateDirectory(txtoutput.Text + "\\OutPut");                
            }
            if (!Directory.Exists(txtoutput.Text + "\\OutPut\\Consolidated"))
            {
                Directory.CreateDirectory(txtoutput.Text + "\\OutPut\\Consolidated");
            }
           
            int i = 0;
            var noFiles = filesList.Count();
            foreach (string inputFile in filesList)
            {
                processingStatus.Text = $"Processing  {i}";
                var fileExt = Path.GetExtension(inputFile);
                filestatus.Text = $"Working on File {Path.GetFileNameWithoutExtension(inputFile)}";
                var dataInput = LoadExcelfromMicrosoftInterop(inputFile);
                if (dataInput != null)
                {
                    var ProcessedfileName = txtoutput.Text + "\\Processed\\" + Path.GetFileNameWithoutExtension(inputFile) + fileExt;
                    File.Move(inputFile, ProcessedfileName);
                    ComputeAndCreateFinalAggregratedOutput(dataInput, txtoutput.Text + "\\OutPut");
                }
                else
                {
                    KillExcelProcesses();
                    errorCount = errorCount + 1;
                    File.Move(inputFile, txtoutput.Text + "\\Errors");
                }
                i = i + 1;
                filestatus.Text = $"File Completed :{Path.GetFileNameWithoutExtension(inputFile)}";                
            }
            timer.Stop();
            MessageBox.Show("Completed");
            
        }

        public DataTable LoadExcelfromMicrosoftInterop(string filePath)
        {
            var dtInputData = new DataTable();

            var excelApp = new Application();
            excelApp.Visible = false;  // Do not show Excel UI

            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath,
                                                                ReadOnly: true);
            Worksheet worksheet = (Worksheet)workbook.Sheets[1];
            try
            {
                
                // Get the range of used cells
                Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;
                bool dataonly = false;
                bool addSalaryType = false;
                bool addDDONumber = false;
                int dataHeader;
                string rootDDO = "";
                // Loop through rows and columns
                for (int row = 1; row <= rowCount; row++)
                {
                    tstripstatus.Text = $"Reading Row : {row}";
                    if (dataonly || ((Range)usedRange.Cells[row, 1]).Value2 == "S No.")
                    {
                        if (!dataonly)
                        {
                            dataonly = true;
                            dataHeader = row;
                            var ddoCell = (Microsoft.Office.Interop.Excel.Range)usedRange.Cells[row - 1, 1];
                            if (ddoCell != null && !String.IsNullOrEmpty(ddoCell.Value2.ToString()))
                            {
                                if (ddoCell.Value2.ToString().Contains("DDO Code"))
                                {
                                    rootDDO = ddoCell.Value2.ToString().Split(':')[1];
                                    rootDDO = rootDDO.Trim();
                                }
                            }
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cell = (Range)usedRange.Cells[row, col];
                                dtInputData.Columns.Add(cell.Value2.Trim());
                            }
                            dtInputData.Columns.Add("EntryType");

                            if (!dtInputData.Columns.Contains("Type"))
                            {
                                dtInputData.Columns.Add("Type");
                                addSalaryType = true;
                            }

                            if (!dtInputData.Columns.Contains("DDO Code"))
                            {
                                dtInputData.Columns.Add("DDO Code");
                                addDDONumber = true;
                            }
                            continue;
                        }
                        DataRow dr = dtInputData.NewRow();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cell = (Microsoft.Office.Interop.Excel.Range)usedRange.Cells[row, col];
                            var cellValue = cell.Value2;
                            if(cellValue!=null && dtInputData.Columns[col - 1].ColumnName.Trim() == "Paybill Generation Date")
                            {
                                if ( double.TryParse(cellValue.ToString(), out double oaDate))
                                {
                                    // If it's a serial number, convert it to DateTime
                                    DateTime dateValue = DateTime.FromOADate(oaDate);
                                    dr[dtInputData.Columns[col - 1].ColumnName.Trim()] = dateValue.ToShortDateString();
                                }
                            }
                            else
                            { 
                                dr[dtInputData.Columns[col - 1].ColumnName.Trim()] = cellValue;
                            }
                        }

                        if (addSalaryType)
                        {
                            dr["Type"] = "SALARY";
                        }

                        if (addDDONumber)
                        {
                            dr["DDO Code"] = rootDDO;
                        }

                        dr["EntryType"] = "Received";

                        if ((dr.IsNull(0) && dr.IsNull(1) && dr.IsNull(3) && dr.IsNull(4)) || dr[0].Equals("Totals"))
                            continue;

                        

                        dtInputData.Rows.Add(dr);
                    }
                }
                tstripstatus.Text = $"File Reading Completed";
                return dtInputData;
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                workbook.Close(false);
                excelApp.Quit();

                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            return null;
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
                "Jun",
                "Jul",
                "Aug",
                "Sep",
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
                tstripstatus.Text = $"Started Processing Pan No :{panNo}";
                var dtPanSpecific = dtInputData.AsEnumerable().Where(x => x.Field<string>("PAN No") == panNo).CopyToDataTable();

                var updatedProvisionalData = UpdateProvisionalData(dtPanSpecific);

                AggregratedOutput(updatedProvisionalData, fileoutpath);
                tstripstatus.Text = $"Processing Pan No :{panNo} Completed";
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
            var cnsFile = fileoutputPath + "\\Consolidated\\ConsolidatedData.csv";
            if (!File.Exists(cnsFile))
            {
                File.WriteAllText(cnsFile, "S No.,DDO Code,KGID No,From Month,To Month,Year,Employee Name,Designation,PAN No,Gross Salary from Current DDO,Gross Salary (Provisional Considered ),Gross Salary (From Other DDO's ),Basic Pay,Stagnation Increment,DA,HRA,Uniform Allowance,Independent Charge Allowance,Medical Allowance,Other Allowances,Income Tax from Current DDO,Income Tax from Other DDO,EGIS,PT,LIC,Nps Deduction Amount,KGID,GPF,HBA,Housing Development Finance Corporation,Arogya Bhagya Yojana");
            }
                
            var results = new Dictionary<string, (decimal total, decimal currentEmployerTotal, decimal cumulativeTotal)>();

            var data = updatedProvisionalData.AsEnumerable();
            var areMultipleEmployers = data.GroupBy(row => row.Field<string>("DDO Code")).Count();
            var lastData = data.Last();
            var firstData = data.First();
            var currentEmployer = lastData["DDO Code"].ToString();
            var kgidno = lastData["KGID No"];
            var employeeName = lastData["Employee Name"];
            var panNo = lastData["PAN No"];
            var designation = lastData["Designation"];
            var fromMonth = firstData["Month"];
            var toMonth = lastData["Month"];
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
                if (updatedProvisionalData.Columns.Contains(field))
                {
                    var total = calculateSum(data, field);

                    decimal currentEmployerTotal = calculateSum(data, field, currentEmployer);

                    decimal cumulativeTotal = total - currentEmployerTotal;

                    results[field] = (total, currentEmployerTotal, cumulativeTotal);
                }
                else
                {
                    results[field] = (0, 0, 0);
                }
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
                sheet2range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCyan);

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
                Color borderColor = System.Drawing.Color.Black;
                sheet2rangeBackground.Style.Border.Top.Color.SetColor(borderColor);
                sheet2rangeBackground.Style.Border.Bottom.Color.SetColor(borderColor);
                sheet2rangeBackground.Style.Border.Left.Color.SetColor(borderColor);
                sheet2rangeBackground.Style.Border.Right.Color.SetColor(borderColor);

                worksheet2.Cells[worksheet.Dimension.Address].AutoFitColumns();
                
                // Save the file
                FileInfo fi = new FileInfo(finalFilename);
                package.SaveAs(fi);
            }

            var lineNo = File.ReadAllLines(cnsFile).Length;

            StringBuilder str = new StringBuilder();
            string[] strdata = new string[31];
            
            strdata[0] = lineNo.ToString();
            strdata[1] = currentEmployer.ToString();
            strdata[2] = kgidno.ToString();
            strdata[3] = fromMonth.ToString();
            strdata[4] = toMonth.ToString();
            strdata[5] = firstData["Year"].ToString();
            strdata[6] = employeeName.ToString();
            strdata[7] = designation.ToString();
            strdata[8] = panNo.ToString();
            var grossSalaryfromCurrentDDo = data.Where(row => row.Field<string>("DDO Code").Equals(currentEmployer) && row.Field<string>("EntryType") == "Received")
                .Sum(sal => Convert.ToDecimal(sal.Field<string>("Gross Salary")));

            strdata[9] = grossSalaryfromCurrentDDo.ToString();

            var grossSalaryfromCurrentDDoprovisional = data.Where(row => row.Field<string>("DDO Code").Equals(currentEmployer) && row.Field<string>("EntryType") != "Received")
                .Sum(sal => Convert.ToDecimal(sal.Field<string>("Gross Salary")));
            strdata[10] = grossSalaryfromCurrentDDoprovisional.ToString();

            results.TryGetValue("Gross Salary", out var totalsalary);
            strdata[11] = (Convert.ToInt64(totalsalary.total) - (Convert.ToInt64(grossSalaryfromCurrentDDo) +Convert.ToInt64(grossSalaryfromCurrentDDoprovisional))).ToString();
            
            results.TryGetValue("Basic Pay", out var baseicpay);
            strdata[12] = baseicpay.total.ToString();

            results.TryGetValue("Stagnation Increment", out var StagnationIncrement);
            strdata[13] = StagnationIncrement.total.ToString();

            results.TryGetValue("DA", out var DA);
            strdata[14] = DA.total.ToString();

            results.TryGetValue("HRA", out var HRA);
            strdata[15] = HRA.total.ToString();

            results.TryGetValue("Uniform Allowance", out var UniformAllowance);
            strdata[16] = UniformAllowance.total.ToString();

            results.TryGetValue("Independent Charge Allowance", out var IndependentChargeAllowance);
            strdata[17] = IndependentChargeAllowance.total.ToString();

            results.TryGetValue("Medical Allowance", out var MedicalAllowance);
            strdata[18] = MedicalAllowance.total.ToString();

            results.TryGetValue("Other Allowances", out var OtherAllowances);
            strdata[19] = OtherAllowances.total.ToString();

            results.TryGetValue("Income Tax", out var incometax);
            strdata[20] = incometax.currentEmployerTotal.ToString();

            strdata[21] = incometax.cumulativeTotal.ToString();

            results.TryGetValue("EGIS", out var EGIS);
            strdata[22] = EGIS.total.ToString();

            results.TryGetValue("PT", out var PT);
            strdata[23] = PT.total.ToString();

            results.TryGetValue("LIC", out var LIC);
            strdata[24] = LIC.total.ToString();

            results.TryGetValue("Nps Deduction Amount", out var NpsDeductionAmount);
            strdata[25] = NpsDeductionAmount.total.ToString();

            results.TryGetValue("KGID", out var KGID);
            strdata[26] = KGID.total.ToString();

            results.TryGetValue("GPF", out var GPF);
            strdata[27] = GPF.total.ToString();

            results.TryGetValue("HBA", out var HBA);
            strdata[28] = HBA.total.ToString();

            results.TryGetValue("Housing Development Finance Corporation", out var HousingDevelopmentFinanceCorporation);
            strdata[29] = HousingDevelopmentFinanceCorporation.total.ToString();

            results.TryGetValue("Arogya Bhagya Yojana", out var ArogyaBhagyaYojana);
            strdata[30] = ArogyaBhagyaYojana.total.ToString();

            string finaldata = string.Join(",", strdata);

            File.AppendAllText(cnsFile,Environment.NewLine+finaldata);
        }

        

        private DataTable UpdateProvisionalData(DataTable dtPanSpecific)
        {
            var receivedSalaryRows = dtPanSpecific.AsEnumerable()
                               .Where(row => row.Field<string>("EntryType") == "Received" && row.Field<string>("Type") == "SALARY");

            var receivedSalaryMonthsCount = receivedSalaryRows.Count();

            if (receivedSalaryRows.Any() && receivedSalaryRows.Count() < 12)/*If credited for 12 months, no need to add provisional*/
            {
                var lastSalaryCreditRow = receivedSalaryRows.Last();

                var lastCreditedMonth = lastSalaryCreditRow["Month"].ToString().Substring(0, 3);

                var months = GetMonths();

                var lastCreditedMonthIndex= months.IndexOf(lastCreditedMonth);
                var provisionalMonthIndex = months.IndexOf(cmbProvisionalMonths.SelectedItem.ToString());

                if(provisionalMonthIndex > lastCreditedMonthIndex)
                    return dtPanSpecific;

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
                    if (provRow["Month"] == "Jan" || provRow["Month"] == "Feb")
                    {
                        provRow["Year"] = "2025";
                    }
                    dtPanSpecific.Rows.Add(provRow);

                    //prepare for next row
                    lastSalaryCreditRow = provRow;
                    lastCreditedMonth = currentMonth;

                    if(currentMonth == "Feb")
                        break;
                }
            }
            return dtPanSpecific;
        }

        private void lblStatus_Click(object sender, EventArgs e)
        {

        }

        private void btnloadfiles_Click(object sender, EventArgs e)
        {
            lstboxFiles.Items.Clear();
            if (string.IsNullOrEmpty(txtPath.Text) || string.IsNullOrEmpty(txtoutput.Text)
                || cmbProvisionalMonths.SelectedItem ==null)
            {
                MessageBox.Show("No Folder Path Selected", "Error");
            }
            else
            {
                var filesList = Directory.GetFiles(txtPath.Text, "*.xls*",SearchOption.AllDirectories);
                if (filesList.Count() == 0)
                {
                    MessageBox.Show("No files to Process");
                    return;
                }

                if (!Directory.Exists(txtoutput.Text + "\\Processed"))
                {
                    Directory.CreateDirectory(txtoutput.Text + "\\Processed");
                }
                if (!Directory.Exists(txtoutput.Text + "\\Errors"))
                {
                    Directory.CreateDirectory(txtoutput.Text + "\\Errors");
                }
                if (!Directory.Exists(txtoutput.Text + "\\OutPut"))
                {
                    Directory.CreateDirectory(txtoutput.Text + "\\OutPut");
                }

                foreach (var file in filesList)
                {
                    lstboxFiles.Items.Add(Path.GetFileName(file));
                }
                toolStripStatusLabel1.Text = "Toatl Files :" + filesList.Count();
                lstboxFiles.Visible = true;
                btnProcess.Visible = true;
            }
        }

        private void killOpenExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KillExcelProcesses();
        }
        public void KillExcelProcesses()
        {
            // Get all processes with the name "EXCEL"
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");

            // Loop through each process and kill it
            foreach (Process process in excelProcesses)
            {
                process.Kill();
            }

            MessageBox.Show("Completed");
        }

        private void ItrCalc_Load(object sender, EventArgs e)
        {

        }
    }
}
