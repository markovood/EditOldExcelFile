using System.Collections.Generic;
using System.Drawing;

using ExpertXls.ExcelLib;

namespace ChangeTableLineBackgroundColor
{
    public class Program
    {
        public static void Main()
        {
            // Opens a workbook from the specified Excel file
            string path = @"..\..\imagine.xls";
            ExcelWorkbook excelFile = new ExcelWorkbook(path);

            var worksheet = excelFile.Worksheets[0];
            worksheet.Activate();

            List<int> Rs = new List<int>(1220);
            List<int> Gs = new List<int>(1220);
            List<int> Bs = new List<int>(1220);

            var allRows = worksheet.UsedRangeRows;
            for (int i = 1; i < allRows.Count; i++)
            {
                var currentRow = allRows[i];
                Rs.Add(int.Parse(currentRow.Cells[3].Value.ToString()));
                Gs.Add(int.Parse(currentRow.Cells[4].Value.ToString()));
                Bs.Add(int.Parse(currentRow.Cells[5].Value.ToString()));
            }

            // sets each row's background color to the specified one in the .xls file
            for (int i = 1; i <= Rs.Count; i++)
            {
                allRows[i].Style.Fill.FillType = ExcelCellFillType.SolidFill;
                allRows[i].Style.Fill.SolidFillOptions.BackColor = Color.FromArgb(255, Rs[i - 1], Gs[i - 1], Bs[i - 1]);
            }

            excelFile.Save(@"..\..\Imagine(Modified).xls");
            excelFile.Close();
        }
    }
}