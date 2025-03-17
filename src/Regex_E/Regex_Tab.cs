using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms.VisualStyles;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace Regex_E
{
    public partial class Regex_Tab
    {
        private void Regex_Tab_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void BtnMain_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new SearchAndReplace();
            form.ShowIcon = false;
            form.TopMost = true;
            form.Show();
            //var ggg = Globals.ThisAddIn.Application;
            //var sheet = ggg.ActiveSheet;
            //Range rr = sheet.UsedRange;
            //for (int row = 0; row < rr.Rows.Count; row++)
            //{
            //    for (int col = 0; col < rr.Columns.Count; col++)
            //    {
            //        var cellrr = (rr.Cells[row, col] as Range).Value;
            //        var ee = cellrr as string;
            //        if (ee!=null)
            //        {
            //            rr.Cells[row, col] = Regex.Replace(ee,"","");
            //        }
            //    }
            //}
        }
    }
}
