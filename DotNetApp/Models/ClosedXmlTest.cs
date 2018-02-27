using System.IO;
using System.Linq;
using System.Text;
using ClosedXML;

namespace DotNetApp.Models
{
    public class ClosedXmlTest
    {
        public static void ClosedXmlReader(string inFile, string outFile)
        {
            var workbook = new ClosedXML.Excel.XLWorkbook(inFile);
            ClosedXML.Excel.IXLWorksheet sheet = workbook.Worksheets.Worksheet(1);
            var sb = new StringBuilder();
            var rows = sheet.LastRowUsed().RangeAddress.FirstAddress.RowNumber;
            var columns = sheet.LastColumnUsed().RangeAddress.FirstAddress.ColumnNumber;

            foreach(var row in sheet.Rows())
            {
                foreach (var cell in row.Cells())
                {
                    string s = cell.GetString();
                    if (s.Contains('\n'))
                        sb.Append("\"" + s + "\"\t");
                    else
                        sb.Append(s + "\t");
                }
                sb.Remove(sb.Length - 1, 1);
                sb.Append("\n");
            }
            File.WriteAllText(outFile, sb.ToString());
        }

        public static void ClosedXmlWriter(string inFile, string outFile)
        {
            var workbook = new ClosedXML.Excel.XLWorkbook();
            ClosedXML.Excel.IXLWorksheet sheet = workbook.Worksheets.Add("Sheet1");
            using (var sr = new StreamReader(File.OpenRead(inFile)))
            {
                int r = 1;
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

                    var row = sheet.Row(r);
                    string[] columns = line.Split('\t');
                    for (var c = 0; c < columns.Length; c++)
                    {
                        if (double.TryParse(columns[c], out double value))
                            row.Cell(c + 1).Value = value;
                        else
                            row.Cell(c + 1).Value = columns[c].Trim('"');
                    }
                    r++;
                }
            }
            workbook.SaveAs(outFile);
        }
    }
}
