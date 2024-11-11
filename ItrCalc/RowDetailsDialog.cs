using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ItrCalc
{
    public partial class provisonalDataDialog : Form
    {
        public provisonalDataDialog()
        {
            InitializeComponent();
        }

        public void SetDictionary(Dictionary<string, (decimal total, decimal currentEmployerTotal, decimal cumulativeTotal)> data)
        {
            dataGridView1.Columns.Add("Type", "Type");
            dataGridView1.Columns.Add("Total", "Total");
            dataGridView1.Columns.Add("CurrentDDO", "Current DDO");
            dataGridView1.Columns.Add("PreviousDDOs", "Cumulative");

            foreach (var item in data)
            {
                dataGridView1.Rows.Add(item.Key, item.Value.total, item.Value.currentEmployerTotal, item.Value.cumulativeTotal);
            }
            
        }

        public void SetDataTable(DataTable table)
        {
            foreach (DataColumn column in table.Columns)
            {
                dataGridView1.Columns.Add(column.ColumnName, column.ColumnName);
            }

            foreach (DataRow row in table.Rows)
            {
                //add columns for each field in the datarow
                //foreach (DataColumn column in row.Table.Columns)
                //    {
                //    dataGridView1.Columns.Add(column.ColumnName, column.ColumnName);
                //    }

                // add row data
                var values = row.ItemArray;
                dataGridView1.Rows.Add(values);
            }            
        }
    }
}
