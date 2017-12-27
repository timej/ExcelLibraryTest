using System.IO;
using System.Text;

namespace ExcelLibraryTest.Models
{
    public class NpoiTest
    {
        public static void NpoiXlsReader(string inFile, string outFile)
        {
            FileStream stream = File.OpenRead(inFile);
            var book = new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
            stream.Close();
            NPOI.SS.UserModel.ISheet sheet = book.GetSheetAt(0);
            int lastRowNum = sheet.LastRowNum;

            var sb = new StringBuilder();

            for (int r = 0; r < sheet.LastRowNum; r++)
            {
                NPOI.SS.UserModel.IRow datarow = sheet.GetRow(r);
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
                        if (cell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                            sb.Append(cell.NumericCellValue + "\t");
                        else
                            sb.Append(cell.StringCellValue + "\t");
                        c++;
                    }
                }
                sb.Append("\n");
            }

            File.WriteAllText(outFile, sb.ToString());
        }

        public static void NpoiXlsxReader(string inFile, string outFile)
        {
            FileStream stream = File.OpenRead(inFile);
            var book = new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
            stream.Close();
            NPOI.SS.UserModel.ISheet sheet = book.GetSheetAt(0);
            int lastRowNum = sheet.LastRowNum;

            var sb = new StringBuilder();

            for (int r = 0; r < sheet.LastRowNum; r++)
            {
                var datarow = sheet.GetRow(r);
                {
                    foreach (NPOI.SS.UserModel.ICell t in datarow.Cells)
                        sb.Append(t.StringCellValue + "\t");
                    sb.Append("\n");
                }
            }
            File.WriteAllText(outFile, sb.ToString());
        }


        public static void NpoiXlsWriter(string inFile, string outFile)
        {
            FileStream stream = File.OpenRead(inFile);
            var book = new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
            stream.Close();
            NPOI.SS.UserModel.ISheet sheet = book.GetSheetAt(0);

            for (int r = 0; r < sheet.LastRowNum; r++)
            {
                var datarow = sheet.GetRow(r);
                {
                    for (int c = 0; c < datarow.Cells.Count; c++)
                        datarow.Cells[c].SetCellValue(r + c);
                }
            }
            var ws = File.OpenWrite(outFile);
            book.Write(ws);
        }

        public static void NpoiXlsxWriter(string inFile, string outFile)
        {
            FileStream stream = File.OpenRead(inFile);
            var book = new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
            stream.Close();
            NPOI.SS.UserModel.ISheet sheet = book.GetSheetAt(0);

            for (int r = 0; r < sheet.LastRowNum; r++)
            {
                var datarow = sheet.GetRow(r);
                {
                    for (int c = 0; c < datarow.Cells.Count; c++)
                        datarow.Cells[c].SetCellValue(r + c);
                }
            }

            FileStream streamw = File.OpenWrite(outFile);
            book.Write(streamw);
        }

    }
}