﻿using DotNetApp.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetApp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!Directory.Exists("App_Data/DataOut"))
                Directory.CreateDirectory("App_Data/DataOut");
            if (args.Contains("-d") && args.Contains("-r") && args.Contains("-x"))
            {
                ExcelDataReaderTest.ExcelDataReader("App_Data/DataIn/am0411.xlsx", "App_Data/DataOut/dx.txt");
            }
            else if (args.Contains("-d") && args.Contains("-r"))
            {
                ExcelDataReaderTest.ExcelDataReader("App_Data/DataIn/am0411.xls", "App_Data/DataOut/ds.txt");
            }
            else if (args.Contains("-n") && args.Contains("-r") && args.Contains("-x"))
            {
                NpoiTest.NpoiXlsxReader("App_Data/DataIn/am0411.xlsx", "App_Data/DataOut/nx.txt");
            }
            else if (args.Contains("-n") && args.Contains("-r"))
            {
                NpoiTest.NpoiXlsReader("App_Data/DataIn/am0411.xls", "App_Data/DataOut/ns.txt");
            }
            else if (args.Contains("-n") && args.Contains("-w") && args.Contains("-x"))
            {
                NpoiTest.NpoiXlsxWriter("App_Data/DataIn/am0411.txt", "App_Data/DataOut/nx.xlsx");
            }
            else if (args.Contains("-n") && args.Contains("-w"))
            {
                NpoiTest.NpoiXlsWriter("App_Data/DataIn/am0411.txt", "App_Data/DataOut/ns.xls");
            }
            else if (args.Contains("-e") && args.Contains("-r") && args.Contains("-x"))
            {
                EpPlusTest.EpPlusReader("App_Data/DataIn/am0411.xlsx", "App_Data/DataOut/ex.txt");
            }
            else if (args.Contains("-e") && args.Contains("-w") && args.Contains("-x"))
            {
                EpPlusTest.EpPlusWriter("App_Data/DataIn/am0411.txt", "App_Data/DataOut/ex.xlsx");
            }
            else if (args.Contains("-c") && args.Contains("-r") && args.Contains("-x"))
            {
                ClosedXmlTest.ClosedXmlReader("App_Data/DataIn/am0411.xlsx", "App_Data/DataOut/cx.txt");
            }
            else if (args.Contains("-c") && args.Contains("-w") && args.Contains("-x"))
            {
                ClosedXmlTest.ClosedXmlWriter("App_Data/DataIn/am0411.txt", "App_Data/DataOut/cx.xlsx");
            }
        }
    }
}
