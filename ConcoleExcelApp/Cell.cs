
namespace ConcoleExcelApp
{
    public enum CellDataType
    {
        Empty,
        Numeric,
        Text,
        Formula,
        Error
    };

    public class Cell
    {
        public string Index { get; set; }
        public string Data { get; set; }
        public CellDataType DataType { get; set; }

        public Cell (string data)
        {
            Data = data;
        }
    }
}
