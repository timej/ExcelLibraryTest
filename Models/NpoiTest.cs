using System.IO;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelLibraryTest.Models
{
    public class NpoiTest
    {
        public static void NpoiXlsReader(string inFile, string outFile)
        {
            FileStream stream = File.OpenRead(inFile);
            var book = new HSSFWorkbook(stream);
            stream.Close();
            ISheet sheet = book.GetSheetAt(0);
            int lastRowNum = sheet.LastRowNum;

            var sb = new StringBuilder();

            for (int r = 0; r < sheet.LastRowNum; r++)
            {
                IRow datarow = sheet.GetRow(r);
                if (datarow != null)
                {
                    int c = 0;
                    foreach (var cell in datarow.Cells)
                    {
                        for (int i = c; i < cell.ColumnIndex; i++)
                        {
                            sb.Append("\t");
                            c++;
                        }
                        if (cell.CellType == CellType.Numeric)
                            sb.Append(cell.NumericCellValue + "\t");
                        else
                        {
                            if (cell.StringCellValue.Contains("\n"))
                                sb.Append("\"" + cell.StringCellValue + "\"\t");
                            else
                                sb.Append(cell.StringCellValue + "\t");
                        }
                        c++;
                    }
                }

                sb.Remove(sb.Length - 1, 1);
                sb.Append("\n");
            }

            File.WriteAllText(outFile, sb.ToString());
        }

        public static void NpoiXlsxReader(string inFile, string outFile)
        {
            FileStream stream = File.OpenRead(inFile);
            var book = new XSSFWorkbook(stream);
            stream.Close();
            ISheet sheet = book.GetSheetAt(0);
            int lastRowNum = sheet.LastRowNum;

            var sb = new StringBuilder();

            for (int r = 0; r < sheet.LastRowNum; r++)
            {
                var datarow = sheet.GetRow(r);
                {
                    foreach (ICell cell in datarow.Cells)
                    { 
                        if(cell.CellType == CellType.Numeric)
                            sb.Append(cell.NumericCellValue + "\t");
                        else
                        {
                            if(cell.StringCellValue.Contains("\n"))
                                sb.Append("\"" + cell.StringCellValue + "\"\t");
                            else
                                sb.Append(cell.StringCellValue + "\t");
                        }
                    }
                    sb.Remove(sb.Length - 1, 1);
                    sb.Append("\n");
                }
            }
            File.WriteAllText(outFile, sb.ToString());
        }


        public static void NpoiXlsWriter(string inFile, string outFile)
        {
            var book = new HSSFWorkbook();
            var sheet = book.CreateSheet("sheet1");
            WriteSheet(sheet, inFile, outFile);
            using (var ws = File.OpenWrite(outFile))
                book.Write(ws);
        }

        public static void NpoiXlsxWriter(string inFile, string outFile)
        {
            var book = new XSSFWorkbook();
            var sheet = book.CreateSheet("sheet1");
            WriteSheet(sheet, inFile, outFile);
            using (var ws = File.OpenWrite(outFile))
                book.Write(ws);
        }

        //tsvに改行文字がある場合に対応
        private static void WriteSheet(ISheet sheet, string inFile, string outFile)
        {
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

                    string[] columns = line.Split('\t');
                    for (var c = 0; c < columns.Length; c++)
                    {
                        var row = sheet.GetRow(r) ?? sheet.CreateRow(r);
                        var cell = row.GetCell(c) ?? row.CreateCell(c);

                        if (double.TryParse(columns[c], out double value))
                        {
                            cell.SetCellValue(value);
                        }
                        else
                            cell.SetCellValue(columns[c].Trim('"'));
                    }

                    r++;
                }
            }
        }
    }
}