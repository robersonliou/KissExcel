using System;
using ConsoleAppCore.Attributes;
using KissExcel.Attributes;

namespace ConsoleAppCore.Models
{
    internal class OrderModel
    {

        [MyIndexColumnParser]
        [ColumnIndex(0)]
        public int Id { get; set; }
        
        //[ColumnName("項目")]
        [ColumnIndex(1)]
        public string Item { get; set; }
        
        //[ColumnName("價格")]
        public int? Price { get; set; }

        //[ColumnName("數量")]
        public int? Amount { get; set; }

        //[ColumnName("總價")]
        [ColumnIndex(4)]
        public int Total { get; set; }
        
    }
}