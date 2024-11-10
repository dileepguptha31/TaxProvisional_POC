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
