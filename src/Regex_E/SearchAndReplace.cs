using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Regex_E
{
    public partial class SearchAndReplace : Form
    {
        public SearchAndReplace()
        {
            InitializeComponent();
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }
        private void UpdateResultGrid()
        {
            //clear old data.
            //add new results.
        }

        private void BtnReplaceAll_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text.Length == 0)
            {
                Debug.WriteLine("null search input.");
                MessageBox.Show("oops","type what you want to search.",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }

            var search = this.textBox1.Text;
            var replace = this.textBox2.Text;
            Range sheetRange = Globals.ThisAddIn.Application.ActiveSheet.UsedRange;
            var rangeRowCount = sheetRange.Rows.Count;
            var rangeColCount = sheetRange.Columns.Count;

            for (int row = 1; row <= rangeRowCount; row++)
            {
                for (int col = 1; col <= rangeColCount; col++)
                {
                    if (sheetRange.Cells[row, col].Value is string cellString)
                    {
                        var newString = string.Empty;
                        if (this.checkBoxRegexSwitchOn.Checked)
                        {
                            sheetRange.Cells[row, col] = Regex.Replace(cellString, search, replace);
                        }
                        else
                        {
                            sheetRange.Cells[row, col] = cellString.Replace(search, replace);
                        }
                    }

                }
            }
        }
    }
}
