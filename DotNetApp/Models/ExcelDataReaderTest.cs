using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace DotNetApp.Models
{
    public class ExcelDataReaderTest
    {

        public static void ExcelDataReader(string inFile, string outFile)
        {
            using (var stream = File.Open(inFile, FileMode.Open, FileAccess.Read))           
            using (var excelReader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = excelReader.AsDataSet();

                var sb = new StringBuilder();
  
                foreach (DataRow datarow in result.Tables[0].Rows)
                {
                    sb.Append(datarow.ItemArray.Aggregate((s, x) => s + "\t" + (x.ToString().Contains("\n")?"\"" + x.ToString() + "\"" : x.ToString())));
                    sb.Append("\n");
                }
                File.WriteAllText(outFile, sb.ToString());
            }
        }
    }
}