using System.IO;
using System.Linq;
using System.Text;
using NPOI.OpenXmlFormats.Dml.Diagram;
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
                    {
                        if(sheet.Cells[r, c].Text.Contains('\n'))
                            sb.Append("\"" + sheet.Cells[r, c].Text + "\"\t");
                        else
                            sb.Append(sheet.Cells[r, c].Text + "\t");
                    }

                    sb.Remove(sb.Length - 1, 1);
                    sb.Append("\n");
                }
            }
            File.WriteAllText(outFile, sb.ToString());
        }

        ////tsvに改行文字がある場合に対応
        public static void EpPlusWriter(string inFile, string outFile)
        {
            using (var book = new ExcelPackage())
            {
                var sheet = book.Workbook.Worksheets.Add("Sheet1");
                using (var sr = new StreamReader(File.OpenRead(inFile)))
                {
                    int row = 1;
                    string line = "";
                    string s;
                    bool ifInner = false;
                    while ((s = sr.ReadLine()) != null)
                    {
                        if (ifInner)
                        {
                            line += "\n" + s;
                            if (s.Contains("\"\t"))
                                ifInner = false;
                            else
                                continue;
                        }
                        else
                            line = s;
                        
                        int pos;
                        if ((pos = s.LastIndexOf('\t')) > -1 && pos != s.Length - 1)
                        {
                            if (s[pos + 1] == '"' && s[s.Length - 1] != '"')
                            {
                                ifInner = true;
                                continue;
                            }
                        }

                        string[] columns = line.Split('\t');
                        for (var c = 0; c < columns.Length; c++)
                        {
                            if (double.TryParse(columns[c], out double value))
                                sheet.Cells[row, c + 1].Value = value;
                            else
                                sheet.Cells[row, c + 1].Value = columns[c].Trim('"');
                        }
                        row++;
                    }
                }

                book.File = new FileInfo(outFile);
                book.Save();
            }
        }
    }
}