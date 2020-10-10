using System;
using System.Linq;
using ConsoleAppCore.Models;
using DocumentFormat.OpenXml.Drawing.Charts;
using KissExcel.Core;

namespace ConsoleAppCore
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = $@"{Environment.CurrentDirectory}\Docs\sample_parser.xlsx";
            var orderModels = ExcelHub.Reader.Open(path)
                .SheetAs("Sheet1").MapTo<OrderModel>().ToList();

            Console.Read();
        }
    }
}
