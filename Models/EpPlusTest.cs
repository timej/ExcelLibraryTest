using System.IO;
using System.Text;
using OfficeOpenXml;

namespace ExcelLibraryTest.Models
{
    public class EpPlusTest
    {
        public static void EpPlusReader(string infile, string outFile)
        {
            FileInfo newFile = new FileInfo(infile);
            ExcelPackage pck = new ExcelPackage(newFile);
            ExcelWorksheet sheet = pck.Workbook.Worksheets[1];

            var sb = new StringBuilder();
            int rows = sheet.Dimension.Rows;

            for (int r = 1; r <= rows; r++)
            {
                {
                    for (int c = 1; c <= sheet.Dimension.Columns; c++)
                        sb.Append(sheet.Cells[r, c].Text + "\t");
                    sb.Append("\n");
                }
            }
            File.WriteAllText(outFile, sb.ToString());
        }

        public static void EpPlusWriter(string inFile, string outFile)
        {
            FileInfo inFileInfo = new FileInfo(inFile);
            ExcelPackage pck = new ExcelPackage(inFileInfo);
            ExcelWorksheet sheet = pck.Workbook.Worksheets[1];

            for (int r = 1; r <= sheet.Dimension.Rows; r++)
            {
                {
                    for (int c = 1; c <= sheet.Dimension.Columns; c++)
                        sheet.Cells[r, c].Value = r + c;
                }
            }
            FileInfo outFileInfo = new FileInfo(outFile);
            pck.SaveAs(outFileInfo);
        }
    }
}