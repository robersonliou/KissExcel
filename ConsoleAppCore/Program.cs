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
            var path = @"C:\Users\Roberson\Desktop\sample_column_name.xlsx";
            var orderModels = ExcelHub.Reader.Open(path)
                .SheetAs("Sheet1").
                IncludeHeader(true).MapTo<OrderModel>().ToList();

            Console.Read();

        }
    }
}
