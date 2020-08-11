using KissExcel.Attributes;

namespace ConsoleAppCore.Models
{
    internal class OrderModel
    {
        
        [ColumnName("編號")]
        public int Id { get; set; }
        
        [ColumnName("項目")]
        public string Item { get; set; }
        
        [ColumnName("價格")]
        public int Price { get; set; }

        [ColumnName("數量")]
        public int Amount { get; set; }

        [ColumnName("總價")]
        public int Total { get; set; }
        
    }
}