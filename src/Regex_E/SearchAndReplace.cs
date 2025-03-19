using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Regex_E
{
    /// <summary>
    /// 1、可字符串方式和正则方式查找。
    /// 2、可忽略大小写。
    /// 3、可查找并替换。
    /// 4、查找结果可在单元格内定位显示，显示下一个。
    /// 5、所有查找结果信息可在扩展datagrid中显示，且被查找内容以高亮标记。
    /// 6、可一键替换所有查找结果。
    /// 7、可选择搜索范围是当前单元格还是整个工作表。
    /// </summary>
    public partial class SearchAndReplace : Form
    {
        private int DefaultWindowHeight = 0;
        private int ExtendWindowHeight = 750;
        public SearchAndReplace()
        {
            InitializeComponent();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.cellsCache.Clear();
            if (this.textBox1.Text.Length == 0)
            {
                Debug.WriteLine("null search input.");
                MessageBox.Show("oops", "type what you want to search.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var search = this.textBox1.Text;
            var replace = this.textBox2.Text;
            Range sheetRange = Globals.ThisAddIn.Application.ActiveSheet.UsedRange;
            var rangeRowCount = sheetRange.Rows.Count;
            var rangeColCount = sheetRange.Columns.Count;

            //
            for (int row = 1; row <= rangeRowCount; row++)
            {
                for (int col = 1; col <= rangeColCount; col++)
                {
                    if (sheetRange.Cells[row, col].Value is string cellString)
                    {
                        var regexOption = RegexOptions.None;
                        var stringCompare = StringComparison.CurrentCulture;
                        if (this.switchIgnoreCase)
                        {
                            regexOption = RegexOptions.IgnoreCase;
                            stringCompare = StringComparison.CurrentCultureIgnoreCase;
                        }
                        // match indexes should be cached here.
                        if (this.switchRegex && Regex.IsMatch(cellString, search, regexOption))
                            this.cellsCache.Add((row, col, cellString, (int)regexOption));
                        else if (cellString.IndexOf(search, stringCompare) != -1)
                            this.cellsCache.Add((row, col, cellString, (int)stringCompare));
                    }

                }
            }


            this.dataGridView1.Rows.Clear();
            this.Size = new Size(this.Size.Width, this.DefaultWindowHeight);
            if (this.cellsCache.Count == 0) return;

            this.Size = new Size(this.Size.Width, this.ExtendWindowHeight);
            foreach (var (row, col, val, option) in this.cellsCache)
                dataGridView1.Rows.Add("", "", "", $"({row}, {col})", val, "");
        }
        private void UpdateResultGrid()
        {
            //clear old data.
            //add new results.
        }

        private bool switchRegex = default;
        private bool switchIgnoreCase = default;

        private List<(int row, int col, string val, int option)> cellsCache = new List<(int row, int col, string val, int option)>();
        private void BtnReplaceAll_Click(object sender, EventArgs e)
        {
            var search = this.textBox1.Text;
            var replace = this.textBox2.Text;
            Range sheetRange = Globals.ThisAddIn.Application.ActiveSheet.UsedRange;

            foreach (var (row, col, val, option) in this.cellsCache)
            {
                if (this.switchRegex)
                    sheetRange.Cells[row, col] = Regex.Replace(val, search, replace, (RegexOptions)option);
                else
                    sheetRange.Cells[row, col] = ReplaceAndTrackOccurrences(val, search, replace, (StringComparison)option);
            }
        }
        static string ReplaceAndTrackOccurrences(string input, string search, string replacement, StringComparison compare)
        {
            int index = input.IndexOf(search, compare);

            while (index != -1)
            {
                input = input.Substring(0, index) + replacement + input.Substring(index + search.Length);

                index = input.IndexOf(search, index + replacement.Length, compare);
            }

            return input;
        }

        private void checkBoxIgnoreCaseSwitchOn_CheckedChanged(object sender, EventArgs e)
        {
            this.switchIgnoreCase = (sender as System.Windows.Forms.CheckBox).Checked;
        }

        private void checkBoxRegexSwitchOn_CheckedChanged(object sender, EventArgs e)
        {
            this.switchRegex = (sender as System.Windows.Forms.CheckBox).Checked;
        }
        private string TitleText = string.Empty;
        private void SearchAndReplace_Load(object sender, EventArgs e)
        {
            this.DefaultWindowHeight = this.Size.Height;
            this.TitleText = this.Text;
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.TitleText} v{version.Major}.{version.Minor}.{version.Build}";
        }
    }
}







